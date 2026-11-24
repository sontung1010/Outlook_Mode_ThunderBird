"use strict";

/* global ExtensionAPI, ExtensionCommon, ChromeUtils, Services, Cc, Ci, Components, IOUtils */

// Experiment: the Outlook-style reading pane and compose dock.
//
// Thunderbird Conversations shows a whole thread in the message pane, in its
// page chrome://conversations/content/experiment-api/stub.html. This
// experiment adds reading-pane/pane.js and pane.css to that page (newest
// message on top, replies and forwards written inline in the pane) and does
// the privileged work for the page script through a bridge object: reply
// data, address lookup, in-memory drafts, and sending with nsIMsgSend, the
// same sender the compose window uses (so the Sent copy, the outgoing server
// and Microsoft Graph accounts work as usual). It also routes Thunderbird's
// own Reply / Reply all / Forward commands (toolbar, context menu, Ctrl+R) to
// the inline composer while a conversation is shown.
//
// New messages (Write, Ctrl+N, File > New > Message) open in a dock at the
// bottom of the 3-pane tab instead of a compose window, like Outlook: one tab
// per message, minimize / maximize / resize. The dock is a chrome page
// (reading-pane/dock.html) running the same pane.js composer.
//
// The page script never puts message HTML into the (privileged) page: the
// quoted original is only added to the outgoing message here.

var { MailServices } = ChromeUtils.importESModule("resource:///modules/MailServices.sys.mjs");
var { MailUtils } = ChromeUtils.importESModule("resource:///modules/MailUtils.sys.mjs");
var { MsgHdrToMimeMessage } = ChromeUtils.importESModule(
  "resource:///modules/gloda/MimeMessage.sys.mjs"
);
var { ExtensionParent } = ChromeUtils.importESModule(
  "resource://gre/modules/ExtensionParent.sys.mjs"
);
var { setTimeout, clearTimeout } = ChromeUtils.importESModule(
  "resource://gre/modules/Timer.sys.mjs"
);
var { ExtensionSupport } = ChromeUtils.importESModule(
  "resource:///modules/ExtensionSupport.sys.mjs"
);
var { AddonManager } = ChromeUtils.importESModule("resource://gre/modules/AddonManager.sys.mjs");
var { MessageSend } = ChromeUtils.importESModule("resource:///modules/MessageSend.sys.mjs");
var { ExtensionUtils } = ChromeUtils.importESModule("resource://gre/modules/ExtensionUtils.sys.mjs");

const PANE_CONVERSATIONS_ID = "gconversation@xulforum.org";
const PANE_STUB_URL = "chrome://conversations/content/experiment-api/stub.html";
const PANE_PACKAGE = "outlookmode";
// Busts the script and style caches when the add-on is updated.
const PANE_LOAD_ID = Date.now();
const PANE_DOCK_URL = `chrome://${PANE_PACKAGE}/content/dock.html`;
const PANE_DOCK_BAR = 34; // px: the dock's tab bar, all that shows when minimized
const PANE_DOCK_HEIGHT_PREF = "extensions.outlookmode.dockHeight";
const PANE_WINDOW_LISTENER = "outlookPane-compose";
const PANE_COMP_TYPES = {
  new: Ci.nsIMsgCompType.New,
  reply: Ci.nsIMsgCompType.ReplyToSender,
  replyAll: Ci.nsIMsgCompType.ReplyAll,
  forward: Ci.nsIMsgCompType.ForwardInline,
};
const PANE_MAX_DRAFTS = 50;
const PANE_DARK_READER_PREF = "mail.dark-reader.enabled";
const PANE_IMAGE_TYPES = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  webp: "image/webp",
};
const PANE_MAX_IMAGE_BYTES = 2 * 1024 * 1024;
// Outlook writes mail in Calibri/Aptos 11 pt and says so in the message;
// without it, Outlook shows the text in Times New Roman 12 pt.
const PANE_DEFAULT_FONT = "Calibri, Arial, Helvetica, sans-serif";
const PANE_FONT_SIZE = "11pt";

function paneFont() {
  return Services.prefs.getStringPref("msgcompose.font_face", "") || PANE_DEFAULT_FONT;
}

const paneState = {
  extension: null,
  config: {
    enabled: false,
    newestFirst: true,
    expandAll: true,
    foldQuotes: true,
    attachmentsTop: true,
    latestOnly: true,
    inlineReply: true,
    inlineNew: true,
    outlookHeader: true,
    lightPane: true,
    pinButton: true,
    pinKeyword: "olpinned",
    undoSendSeconds: 10,
    replyHeader: true,
    removeQuoteBar: true,
  },
  chromeHandle: null,
  observing: false,
  windows: new Set(), // Conversations pages running pane.js
  routed: new Map(), // about:3pane window -> original command functions
  composeRouted: new Map(), // messenger window -> original ComposeMessage
  windowListener: false,
  dock: null, // the compose dock (one per Thunderbird session)
  newOrder: [], // keys ("new:...") of the new messages, in dock tab order
  newSeq: 0,
  drafts: new Map(), // message URI or new-message key -> draft fields
  files: new Map(), // token of a file attached in the pane -> its local path
  pending: new Map(), // message URI -> { timer, until, fields } (Undo Send)
  listeners: { popOut: new Set(), sendResult: new Set(), pin: new Set(), openSettings: new Set() },
  singleSentCopy: false, // Graph accounts: no Fcc copy of our own (paneFccIsSecondCopy)
  sentCopyListener: false,
  sentCopyWindows: new Set(), // compose windows with the compose-send-message listener
  sentCopiesSkipped: 0, // sends that filed no second copy
  addressCache: null,
  lastError: "",
};

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */

// Page objects must not be kept after the page is gone.
function paneClone(value) {
  return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
}

function paneEscape(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function paneFold(text) {
  return String(text || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .toLowerCase();
}

function paneNote(where, error) {
  paneState.lastError = `${new Date().toISOString()} ${where}: ${error && error.message ? error.message : error}`;
  console.error(`[outlook-mode] reading pane: ${where}`, error);
}

function hdrFromUri(uri) {
  try {
    return MailServices.messageServiceFromURI(uri).messageURIToMsgHdr(uri);
  } catch (error) {
    return null;
  }
}

function uriOfHdr(hdr) {
  return hdr && hdr.folder ? hdr.folder.getUriForMsg(hdr) : null;
}

function conversationsExtension() {
  return ExtensionParent.GlobalManager.getExtension(PANE_CONVERSATIONS_ID) || null;
}

// Outlook style: "Saturday, November 7, 2026 3:05 PM"
function paneOutlookDate(date) {
  if (Number.isNaN(date.getTime())) {
    return "";
  }
  const day = date.toLocaleDateString("en-US", {
    weekday: "long",
    year: "numeric",
    month: "long",
    day: "numeric",
  });
  const time = date.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
  return `${day} ${time}`;
}

function paneParse(raw) {
  if (!raw) {
    return [];
  }
  return MailServices.headerParser.parseEncodedHeaderW(raw).filter((a) => a.email);
}

// "A <a@x>, B <b@x>" typed by the user -> ["A <a@x>", "B <b@x>"]
function paneAddressList(text) {
  const addresses = MailServices.headerParser.parseDecodedHeader(
    String(text || "").replace(/;/g, ",")
  );
  return addresses
    .filter((a) => a.email || a.name)
    .map((a) => {
      if (!/^[^@\s]+@[^@\s]+$/.test(a.email || "")) {
        throw new Error(`Not an email address: ${a.name || a.email}`);
      }
      return a.toString();
    });
}

function paneMyAddresses() {
  return new Set(
    MailServices.accounts.allIdentities.map((i) => (i.email || "").toLowerCase()).filter(Boolean)
  );
}

/* ------------------------------------------------------------------ */
/* The original message                                                */
/* ------------------------------------------------------------------ */

const paneMimeCache = new Map(); // uri -> Promise<MimeMessage|null>

function paneMime(hdr) {
  const uri = uriOfHdr(hdr);
  if (paneMimeCache.has(uri)) {
    return paneMimeCache.get(uri);
  }
  const promise = new Promise((resolve) => {
    const timer = setTimeout(() => resolve(null), 30000);
    try {
      MsgHdrToMimeMessage(
        hdr,
        null,
        (aHdr, mime) => {
          clearTimeout(timer);
          resolve(mime || null);
        },
        true,
        { examineEncryptedParts: false, partsOnDemand: false }
      );
    } catch (error) {
      clearTimeout(timer);
      paneNote("read message", error);
      resolve(null);
    }
  });
  paneMimeCache.set(uri, promise);
  if (paneMimeCache.size > 10) {
    paneMimeCache.delete(paneMimeCache.keys().next().value);
  }
  return promise;
}

function paneOriginalHeader(hdr) {
  const list = (raw) =>
    paneParse(raw)
      .map((a) => a.toString())
      .join("; ");
  const hadRe = Boolean(hdr.flags & Ci.nsMsgMessageFlags.HasRe);
  return {
    from: hdr.mime2DecodedAuthor || "",
    sent: paneOutlookDate(new Date(hdr.date / 1000)),
    to: list(hdr.recipients),
    cc: list(hdr.ccList),
    subject: `${hadRe ? "RE: " : ""}${hdr.mime2DecodedSubject || ""}`,
  };
}

// The body of the original as HTML, for the quote.
function paneOriginalBody(mime) {
  let html = null;
  let text = null;
  const visit = (part) => {
    if (!part || html !== null) {
      return;
    }
    const type = String(part.contentType || "").toLowerCase();
    if (typeof part.body == "string") {
      if (type.startsWith("text/html")) {
        html = part.body;
      } else if (type.startsWith("text/plain") && text === null) {
        text = part.body;
      }
    }
    for (const child of part.parts || []) {
      visit(child);
    }
  };
  visit(mime);
  if (html !== null) {
    html = html.replace(/<script[\s\S]*?<\/script>/gi, "");
    const styles = (html.match(/<style[\s\S]*?<\/style>/gi) || []).join("");
    const body = html.match(/<body[^>]*>([\s\S]*)<\/body>/i);
    return styles + (body ? body[1] : html);
  }
  if (text !== null) {
    return `<div style="white-space: pre-wrap">${paneEscape(text)}</div>`;
  }
  return "";
}

function paneReplyRecipients(hdr, mime, type) {
  const mine = paneMyAddresses();
  const author = paneParse(hdr.author);
  const replyToHeader = mime && mime.headers && mime.headers["reply-to"];
  const replyTo = paneParse(replyToHeader ? replyToHeader[0] : "");
  const fromMe = author.some((a) => mine.has(a.email.toLowerCase()));
  let to;
  let cc = [];
  if (fromMe) {
    // Replying to my own message goes to the people I wrote to.
    to = paneParse(hdr.recipients);
    if (type == "replyAll") {
      cc = paneParse(hdr.ccList);
    }
  } else {
    to = replyTo.length ? replyTo : author;
    if (type == "replyAll") {
      to = [...to, ...paneParse(hdr.recipients)];
      cc = paneParse(hdr.ccList);
    }
  }
  const seen = new Set();
  const keep = (address) => {
    const email = address.email.toLowerCase();
    if (seen.has(email) || (type == "replyAll" && mine.has(email))) {
      return false;
    }
    seen.add(email);
    return true;
  };
  let toList = to.filter(keep);
  if (!toList.length && to.length) {
    toList = [to[0]]; // only my own address was there
  }
  return {
    to: toList.map((a) => a.toString()),
    cc: cc.filter(keep).map((a) => a.toString()),
  };
}

function paneSubject(hdr, type) {
  const base = hdr.mime2DecodedSubject || "";
  const hadRe = Boolean(hdr.flags & Ci.nsMsgMessageFlags.HasRe);
  return type == "forward" ? `FW: ${hadRe ? "RE: " : ""}${base}` : `RE: ${base}`;
}

function paneUserAttachments(mime) {
  if (!mime) {
    return [];
  }
  return mime.allUserAttachments
    .filter((a) => a.url)
    .map((a) => ({ url: a.url, name: a.name || "attachment", contentType: a.contentType || "", size: a.size || 0 }));
}

// Signatures can come from a mail (signature-from-mail), so they are
// untrusted: Thunderbird's sanitizer drops scripts, event handler
// attributes, javascript: URLs, frames, objects and meta/base; forms go too.
function paneSanitize(html) {
  const flags = Ci.nsIParserUtils.SanitizerAllowStyle | Ci.nsIParserUtils.SanitizerDropForms;
  const out = Cc["@mozilla.org/parserutils;1"].getService(Ci.nsIParserUtils).sanitize(html, flags);
  const body = out.match(/<body[^>]*>([\s\S]*)<\/body>/i);
  return body ? body[1] : out;
}

async function paneSignature(identity, type) {
  if (type == "forward" ? !identity.sigOnForward : type != "new" && !identity.sigOnReply) {
    return "";
  }
  let text = "";
  let isHtml = false;
  try {
    const file = identity.attachSignature ? identity.signature : null;
    if (file && file.exists()) {
      text = await IOUtils.readUTF8(file.path);
      isHtml = /\.x?html?$/i.test(file.leafName);
    } else {
      text = identity.htmlSigText || "";
      isHtml = identity.htmlSigFormat;
    }
    if (isHtml) {
      text = await paneEmbedImages(text, file && file.exists() ? file : null);
    }
  } catch (error) {
    paneNote("signature", error);
    return "";
  }
  if (!text.trim()) {
    return "";
  }
  const body = isHtml ? paneSanitize(text) : paneEscape(text).replace(/\r?\n/g, "<br>");
  return (identity.suppressSigSep ? "" : "-- <br>") + body;
}

/* ------------------------------------------------------------------ */
/* Images in signatures and in the message                             */
/* ------------------------------------------------------------------ */

function paneWindow() {
  return Services.wm.getMostRecentWindow("mail:3pane") || Services.appShell.hiddenDOMWindow;
}

// An inert document: nothing in it loads or runs.
function paneParseHtml(html) {
  return new (paneWindow().DOMParser)().parseFromString(html, "text/html");
}

function paneBase64(bytes) {
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return paneWindow().btoa(binary); // the experiment scope has no btoa()
}

// Prepares an HTML signature: its local image files become data: URLs, so
// they show in the composer and are sent inside the message (file: URLs, and
// paths relative to the signature file: Outlook keeps a signature's images
// in "<name>_files/" next to its .htm file). Returns the whole document.
async function paneEmbedImages(html, baseFile) {
  if (!/<img|cellpadding/i.test(html)) {
    return html;
  }
  const doc = paneParseHtml(html);
  // Conversations' page sets every padding to 0, which beats cellpadding:
  // the same spacing as inline styles shows there too (and in the mail).
  for (const table of doc.querySelectorAll("table[cellpadding]")) {
    const padding = parseInt(table.getAttribute("cellpadding"), 10);
    for (const row of padding > 0 ? table.rows : []) {
      for (const cell of row.cells) {
        if (!cell.style.padding) {
          cell.style.padding = `${padding}px`;
        }
      }
    }
  }
  for (const img of doc.querySelectorAll("img[src]")) {
    const src = img.getAttribute("src").trim();
    if (/^(data|https?|cid|blob):/i.test(src)) {
      continue;
    }
    let file = null;
    try {
      if (/^file:/i.test(src)) {
        file = Services.io.newURI(src).QueryInterface(Ci.nsIFileURL).file;
      } else if (baseFile && !/^[a-z][\w+.-]*:/i.test(src)) {
        file = baseFile.parent.clone();
        for (const part of decodeURIComponent(src).split(/[\\/]+/)) {
          if (part == "..") {
            file = file.parent;
          } else if (part && part != ".") {
            file.append(part);
          }
        }
      }
    } catch (error) {
      file = null;
    }
    const type = file && PANE_IMAGE_TYPES[file.leafName.split(".").pop().toLowerCase()];
    if (!type || !file.exists() || file.fileSize > PANE_MAX_IMAGE_BYTES) {
      continue;
    }
    img.setAttribute("src", `data:${type};base64,${paneBase64(await IOUtils.read(file.path))}`);
  }
  return `<!DOCTYPE html>\n${doc.documentElement.outerHTML}`;
}

// The page shows data: images through blob: URLs (Conversations' page only
// allows those) and keeps the data in data-ol-src: put it back. Images
// whose data is gone are dropped rather than sent broken.
function paneRestoreImages(doc) {
  for (const img of doc.querySelectorAll("img")) {
    const original = img.getAttribute("data-ol-src");
    if (original) {
      img.setAttribute("src", original);
      img.removeAttribute("data-ol-src");
    } else if (/^blob:/i.test(img.getAttribute("src") || "")) {
      img.remove();
    }
  }
}

// What MessageSend reads from the compose window's editor: its document
// (data: images in it become cid: parts of a multipart/related message, as
// in the compose window) and its HTML afterwards.
function paneEditorFor(body) {
  const doc = paneParseHtml(body);
  paneRestoreImages(doc);
  return {
    document: doc,
    // As the editor writes it: a no-break space as the character, since
    // MessageSend's link scan (mozITXTToHTMLConv) would escape "&nbsp;".
    outputToString: () => `<!DOCTYPE html>\n${doc.documentElement.outerHTML.replace(/&nbsp;/g, "\u00a0")}`,
  };
}

// Imports an HTML signature saved by Outlook (Windows: %APPDATA%\Microsoft\
// Signatures\<name>.htm with its <name>_files folder) into an identity's
// signature text, with its images inside, so it also works in the compose
// window and syncs between devices.
async function paneImportSignature(identityKey) {
  const identity = paneIdentityByKey(identityKey);
  if (!identity) {
    throw new Error("Unknown identity");
  }
  const win = paneWindow();
  const picker = Cc["@mozilla.org/filepicker;1"].createInstance(Ci.nsIFilePicker);
  picker.init(win.browsingContext, `Outlook signature for ${identity.email}`, Ci.nsIFilePicker.modeOpen);
  picker.appendFilter("Outlook signature (*.htm)", "*.htm; *.html");
  picker.appendFilters(Ci.nsIFilePicker.filterAll);
  const result = await new Promise((resolve) => picker.open(resolve));
  if (result != Ci.nsIFilePicker.returnOK || !picker.file) {
    return { cancelled: true };
  }
  const file = picker.file;
  const bytes = await IOUtils.read(file.path);
  // Outlook writes the charset in a <meta> tag (often windows-1252).
  const head = String.fromCharCode(...bytes.subarray(0, 4096));
  let text;
  try {
    text = new win.TextDecoder((head.match(/charset=["']?([\w-]+)/i) || [])[1] || "utf-8").decode(bytes);
  } catch (error) {
    text = new win.TextDecoder("utf-8").decode(bytes);
  }
  return paneSetSignature(identity, await paneEmbedImages(text, file));
}

// Makes `html` (an Outlook signature: a file's or a sent mail's, pictures
// as data: URLs) the identity's HTML signature.
function paneSetSignature(identity, html) {
  const doc = paneParseHtml(html);
  // Outlook's stylesheet (dropped with <head>) sets no margins on its paragraphs.
  for (const node of doc.querySelectorAll(".MsoNormal")) {
    if (!node.style.margin) {
      node.style.margin = "0";
    }
  }
  const images = doc.querySelectorAll("img[src^='data:']").length;
  const missing = doc.querySelectorAll("img:not([src^='data:']):not([src^='http'])").length;
  const clean = paneSanitize(`<html><body>${doc.body.innerHTML}</body></html>`).trim();
  if (!clean) {
    throw new Error("There is no signature in it.");
  }
  identity.htmlSigFormat = true;
  identity.htmlSigText = clean;
  identity.attachSignature = false;
  identity.suppressSigSep = true; // Outlook puts no "-- " line above it
  Services.prefs.savePrefFile(null);
  return { email: identity.email, images, missing, size: clean.length };
}

function paneIdentityByKey(key) {
  return MailServices.accounts.allIdentities.find((i) => i.key == key) || null;
}

// The identity Thunderbird's ComposeMessage picks for a new message.
function paneIdentityForFolder(folder) {
  let identity = null;
  try {
    if (folder) {
      identity = folder.customIdentity || MailUtils.getIdentityForServer(folder.server)[0];
    }
  } catch (error) {
    // no server identity
  }
  return identity || MailServices.accounts.defaultAccount?.defaultIdentity || null;
}

function paneIdentityLabel(identity) {
  return identity.fullName ? `${identity.fullName} <${identity.email}>` : identity.email;
}

/* ------------------------------------------------------------------ */
/* Bridge: reply data and addresses                                    */
/* ------------------------------------------------------------------ */

async function panePrepare(uri, type) {
  if (!Object.hasOwn(PANE_COMP_TYPES, type)) {
    throw new Error(`Unknown reply type: ${type}`);
  }
  const isNew = type == "new";
  const hdr = isNew ? null : hdrFromUri(uri);
  if (!isNew && !hdr) {
    throw new Error("The message is no longer available.");
  }
  const mime = hdr ? await paneMime(hdr) : null;
  let [identity] = isNew
    ? [paneIdentityByKey(paneState.drafts.get(uri)?.identityKey)]
    : MailUtils.getIdentityForHeader(hdr, PANE_COMP_TYPES[type]);
  if (!identity) {
    identity = MailServices.accounts.defaultAccount?.defaultIdentity || null;
  }
  const identities = [];
  for (const each of MailServices.accounts.allIdentities) {
    if (each.email && !identities.some((i) => i.key == each.key)) {
      identities.push({
        key: each.key,
        label: paneIdentityLabel(each),
        signature: await paneSignature(each, type),
      });
    }
  }
  const recipients =
    type == "forward" || isNew ? { to: [], cc: [] } : paneReplyRecipients(hdr, mime, type);
  return {
    uri,
    type,
    identityKey: identity ? identity.key : identities[0]?.key || "",
    identities,
    to: recipients.to.join(", "),
    cc: recipients.cc.join(", "),
    bcc: "",
    subject: isNew ? "" : paneSubject(hdr, type),
    attachments: type == "forward" ? paneUserAttachments(mime) : [],
    original: hdr ? paneOriginalHeader(hdr) : null,
    font: paneFont(),
    fontSize: PANE_FONT_SIZE,
  };
}

function paneAddressBook() {
  const cache = paneState.addressCache;
  if (cache && Date.now() - cache.at < 5 * 60 * 1000) {
    return cache.list;
  }
  const list = [];
  const seen = new Set();
  for (const directory of MailServices.ab.directories) {
    if (directory.isRemote) {
      continue;
    }
    let cards = [];
    try {
      cards = directory.childCards;
    } catch (error) {
      continue;
    }
    for (const card of cards) {
      if (card.isMailList) {
        continue;
      }
      const name =
        card.displayName || [card.firstName, card.lastName].filter(Boolean).join(" ") || "";
      const emails = card.emailAddresses?.length ? card.emailAddresses : [card.primaryEmail];
      for (const email of emails) {
        const lower = (email || "").toLowerCase();
        if (!lower || seen.has(lower)) {
          continue;
        }
        seen.add(lower);
        list.push({ name, email, key: paneFold(`${name} ${email}`) });
      }
    }
  }
  paneState.addressCache = { at: Date.now(), list };
  return list;
}

function paneSearchAddresses(query) {
  const words = paneFold(query).split(/\s+/).filter(Boolean);
  if (!words.length) {
    return [];
  }
  const results = [];
  for (const entry of paneAddressBook()) {
    const parts = entry.key.split(/[\s.@<>_-]+/);
    if (words.every((w) => parts.some((p) => p.startsWith(w)) || entry.key.includes(w))) {
      results.push(MailServices.headerParser.makeMimeAddress(entry.name, entry.email));
      if (results.length >= 8) {
        break;
      }
    }
  }
  return results;
}

// @mentions in the composer: { name, email, label, address }, where label
// is the name without a trailing "(Department)", as background/mention.js
// does for the compose window. Names starting with the typed word come first.
function paneSearchContacts(query) {
  const words = paneFold(query).split(/\s+/).filter(Boolean);
  if (!words.length) {
    return [];
  }
  const labelOf = (name) => name.replace(/\s*\([^)]*\)\s*$/, "") || name;
  const rank = (entry) =>
    paneFold(labelOf(entry.name)).split(/\s+/).some((part) => part.startsWith(words[0])) ? 0 : 1;
  return paneAddressBook()
    .filter((entry) => words.every((w) => entry.key.includes(w)))
    .sort((a, b) => rank(a) - rank(b))
    .slice(0, 8)
    .map((entry) => {
      const name = entry.name || entry.email;
      return {
        name,
        email: entry.email,
        label: labelOf(name),
        address: MailServices.headerParser.makeMimeAddress(entry.name, entry.email),
      };
    });
}

/* ------------------------------------------------------------------ */
/* Sending                                                             */
/* ------------------------------------------------------------------ */

function paneQuote(hdr, type, mime) {
  const original = paneOriginalBody(mime);
  const head = paneOriginalHeader(hdr);
  const rows = [
    ["From", head.from],
    ["Sent", head.sent],
    ["To", head.to],
    ["Cc", head.cc],
    ["Subject", head.subject],
  ]
    .filter(([, value]) => value)
    .map(([label, value]) => `<b>${label}:</b> ${paneEscape(value)}`)
    .join("<br>");
  if (paneState.config.replyHeader) {
    return (
      '<div id="divRplyFwdMsg" style="border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0 0 0;margin-top:6px">' +
      `<p style="margin:0">${rows}</p></div><div>${original}</div>`
    );
  }
  if (type == "forward") {
    return `<div class="moz-forward-container"><br>-------- Forwarded Message --------<br>${rows}<br><br>${original}</div>`;
  }
  const prefix = `<div class="moz-cite-prefix">On ${paneEscape(head.sent)}, ${paneEscape(head.from)} wrote:<br></div>`;
  return paneState.config.removeQuoteBar
    ? `${prefix}<div>${original}</div>`
    : `${prefix}<blockquote type="cite">${original}</blockquote>`;
}

function paneProgress() {
  return {
    processCanceledByUser: false,
    msgWindow: null,
    openProgressDialog() {},
    closeProgressDialog() {},
    registerListener() {},
    unregisterListener() {},
    onStateChange() {},
    onProgressChange() {},
    onLocationChange() {},
    onStatusChange() {},
    onSecurityChange() {},
    onContentBlockingEvent() {},
    showStatusString() {},
    startMeteors() {},
    stopMeteors() {},
    showProgress() {},
    setStatusString() {},
    setWrappedStatusFeedback() {},
    QueryInterface: ChromeUtils.generateQI([
      "nsIMsgProgress",
      "nsIMsgStatusFeedback",
      "nsIWebProgressListener",
    ]),
  };
}

function paneSendMessage(identity, fields, body, originalUri, compType) {
  return new Promise((resolve, reject) => {
    const listener = {
      onStartSending() {},
      onProgress() {},
      onStatus() {},
      onStopSending(msgId, status, message) {
        if (Components.isSuccessCode(status)) {
          resolve();
        } else {
          reject(new Error(message || `Sending failed (0x${(status >>> 0).toString(16)})`));
        }
      },
      onTransportSecurityError() {},
      onGetDraftFolderURI() {},
      onSendNotPerformed() {
        reject(new Error("The message was not sent."));
      },
      onStartCopy() {},
      setMessageKey() {},
      getMessageId() {
        return null;
      },
      onStopCopy() {},
      QueryInterface: ChromeUtils.generateQI(["nsIMsgSendListener", "nsIMsgCopyServiceListener"]),
    };
    try {
      // The JS class itself, not the XPCOM service: it takes the editor
      // stand-in, which makes the message's data: images cid: parts.
      const sender = new MessageSend();
      Promise.resolve(
        sender.createAndSendMessage(
          paneEditorFor(body),
          identity,
          "",
          fields,
          false,
          false,
          Ci.nsIMsgSend.nsMsgDeliverNow,
          null,
          "text/html",
          body,
          null,
          paneProgress(),
          listener,
          null,
          originalUri,
          compType
        )
      ).catch(reject);
    } catch (error) {
      reject(error);
    }
  });
}

function paneAttachment(url, name, contentType) {
  const attachment = Cc["@mozilla.org/messengercompose/attachment;1"].createInstance(
    Ci.nsIMsgAttachment
  );
  attachment.url = url;
  attachment.name = name;
  if (contentType) {
    attachment.contentType = contentType;
  }
  return attachment;
}

async function paneDeliver(uri, draft) {
  const isNew = draft.type == "new";
  const hdr = isNew ? null : hdrFromUri(uri);
  if (!isNew && !hdr) {
    throw new Error("The original message is no longer available.");
  }
  const identity = MailServices.accounts.getIdentity(draft.identityKey);
  if (!identity || !identity.email) {
    throw new Error("Choose a From address.");
  }
  const fields = Cc["@mozilla.org/messengercompose/composefields;1"].createInstance(
    Ci.nsIMsgCompFields
  );
  fields.from = MailServices.headerParser.makeMimeAddress(identity.fullName || "", identity.email);
  fields.to = paneAddressList(draft.to).join(", ");
  fields.cc = paneAddressList(draft.cc).join(", ");
  const autoBcc = identity.doBcc ? identity.doBccList || "" : "";
  fields.bcc = paneAddressList([draft.bcc, autoBcc].filter(Boolean).join(", ")).join(", ");
  if (!fields.to && !fields.cc && !fields.bcc) {
    throw new Error("Add at least one recipient.");
  }
  fields.subject = draft.subject || "";
  fields.replyTo = identity.replyTo || "";
  fields.organization = identity.organization || "";
  fields.useMultipartAlternative = true;
  if (hdr && draft.type != "forward") {
    const references = [];
    for (let i = 0; i < hdr.numReferences; i++) {
      references.push(`<${hdr.getStringReference(i)}>`);
    }
    references.push(`<${hdr.messageId}>`);
    fields.references = references.join(" ");
  }
  for (const a of draft.attachments || []) {
    fields.addAttachment(paneAttachment(a.url, a.name, a.contentType));
  }
  for (const f of draft.files || []) {
    const path = paneState.files.get(f.id);
    const file = Cc["@mozilla.org/file/local;1"].createInstance(Ci.nsIFile);
    if (path) {
      file.initWithPath(path);
    }
    if (!path || !file.exists()) {
      throw new Error(`File not found: ${f.name}`);
    }
    fields.addAttachment(paneAttachment(Services.io.newFileURI(file).spec, f.name, f.type));
  }

  if (paneFccIsSecondCopy(identity, fields.fcc)) {
    fields.fcc = "nocopy://";
    paneState.sentCopiesSkipped++;
  }

  const own = `<div style="font-family: ${paneEscape(paneFont())}; font-size: ${PANE_FONT_SIZE};">${draft.html || ""}</div>`;
  const body =
    '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head><body>' +
    own +
    (hdr ? paneQuote(hdr, draft.type, await paneMime(hdr)) : "") +
    "</body></html>";

  await paneSendMessage(identity, fields, body, hdr ? uri : "", PANE_COMP_TYPES[draft.type]);
  if (!hdr) {
    return;
  }
  try {
    hdr.folder.addMessageDispositionState(
      hdr,
      draft.type == "forward"
        ? Ci.nsIMsgFolder.nsMsgDispositionState_Forwarded
        : Ci.nsIMsgFolder.nsMsgDispositionState_Replied
    );
  } catch (error) {
    paneNote("mark replied", error);
  }
}

/* ------------------------------------------------------------------ */
/* One copy of sent mail on Microsoft Graph accounts                   */
/* ------------------------------------------------------------------ */

// Thunderbird sends through a Microsoft Graph outgoing server with
// POST /me/messages and /send, and Exchange then keeps the sent message in
// Sent Items itself. Thunderbird still files its own copy (Fcc) as well,
// uploaded as a new message: every sent mail was in Sent Items twice, and
// Exchange stores the uploaded copy as an unsent draft, which Outlook shows
// as "[Draft]" in the conversation. With singleSentCopy on, a send whose Fcc
// folder is on a Graph account files no copy of its own: the reading pane's
// and the dock's sends (paneDeliver) and the compose window's
// (paneSentCopyComposeSend).

function paneOutgoingIsGraph(identity) {
  try {
    const server = MailServices.outgoingServer.getServerByIdentity(identity);
    return String((server && server.type) || "").toLowerCase() == "graph";
  } catch (error) {
    return false;
  }
}

// Whether Thunderbird's copy of a message the identity sends, with the
// compose fields' fcc (empty: the identity's own Sent folder), would be the
// second copy on the server. Same choice of folder as MessageSend's getFcc.
function paneFccIsSecondCopy(identity, fcc) {
  if (!paneState.singleSentCopy || !identity || !paneOutgoingIsGraph(identity)) {
    return false;
  }
  fcc = String(fcc || "");
  if (fcc.startsWith("nocopy://")) {
    return false;
  }
  let folder = fcc ? MailUtils.getExistingFolder(fcc) : null;
  if (!folder) {
    if (!identity.doFcc || !identity.fccFolderURI) {
      return false;
    }
    folder = MailUtils.getExistingFolder(identity.fccFolderURI);
  }
  return Boolean(folder && folder.server && String(folder.server.type).toLowerCase() == "graph");
}

const PANE_SENT_COPY_LISTENER = "outlookPane-sent-copy";
const PANE_SEND_MODES = new Set(
  [Ci.nsIMsgCompDeliverMode.Now, Ci.nsIMsgCompDeliverMode.Later, Ci.nsIMsgCompDeliverMode.Background].filter(
    (mode) => mode !== undefined
  )
);

// The compose window's "compose-send-message" event comes just before
// gMsgCompose.sendMsg(), which builds the message from these compose fields.
function paneSentCopyComposeSend(event) {
  try {
    const win = event.currentTarget;
    if (!PANE_SEND_MODES.has(event.detail && event.detail.msgType)) {
      return; // saving a draft or template
    }
    const fields = win.gMsgCompose && win.gMsgCompose.compFields;
    if (fields && paneFccIsSecondCopy(win.gCurrentIdentity, fields.fcc)) {
      fields.fcc = "nocopy://";
      paneState.sentCopiesSkipped++;
    }
  } catch (error) {
    paneNote("sent copy", error);
  }
}

function paneSentCopyWatch(win) {
  if (!paneState.sentCopyWindows.has(win)) {
    win.addEventListener("compose-send-message", paneSentCopyComposeSend);
    paneState.sentCopyWindows.add(win);
    win.addEventListener("unload", () => paneState.sentCopyWindows.delete(win), { once: true });
  }
}

function paneSetSingleSentCopy(enabled) {
  paneState.singleSentCopy = Boolean(enabled);
  if (paneState.singleSentCopy && !paneState.sentCopyListener) {
    ExtensionSupport.registerWindowListener(PANE_SENT_COPY_LISTENER, {
      chromeURLs: ["chrome://messenger/content/messengercompose/messengercompose.xhtml"],
      onLoadWindow: paneSentCopyWatch,
    });
    paneState.sentCopyListener = true;
  } else if (!paneState.singleSentCopy) {
    paneStopSentCopy();
  }
}

function paneStopSentCopy() {
  if (paneState.sentCopyListener) {
    ExtensionSupport.unregisterWindowListener(PANE_SENT_COPY_LISTENER);
    paneState.sentCopyListener = false;
  }
  for (const win of paneState.sentCopyWindows) {
    try {
      win.removeEventListener("compose-send-message", paneSentCopyComposeSend);
    } catch (error) {
      // the window is closing
    }
  }
  paneState.sentCopyWindows.clear();
}

// Tell the open panes, then the background (which notifies on failure when
// no pane showed it).
function paneReport(uri, result) {
  let shown = false;
  for (const win of paneState.windows) {
    try {
      if (win.outlookModePane && win.outlookModePane.onSendResult(uri, paneClone(result))) {
        shown = true;
      }
    } catch (error) {
      // the page is going away
    }
  }
  for (const fire of paneState.listeners.sendResult) {
    fire.async({ ...result, shown });
  }
}

async function paneDeliverAndReport(uri, draft) {
  try {
    await paneDeliver(uri, draft);
    paneState.drafts.delete(uri);
    paneForgetNew(uri);
    paneReport(uri, { ok: true, subject: draft.subject || "" });
  } catch (error) {
    paneNote("send", error);
    const text = String(error && error.message ? error.message : error);
    paneState.drafts.set(uri, { ...draft, dirty: true });
    paneReport(uri, { ok: false, error: text, subject: draft.subject || "" });
    throw new Error(text);
  }
}

function paneSend(uri, draft) {
  draft = paneClone(draft);
  // Fail fast on a bad address instead of after the Undo Send countdown.
  paneAddressList(draft.to);
  paneAddressList(draft.cc);
  paneAddressList(draft.bcc);
  if (!draft.to.trim() && !draft.cc.trim() && !draft.bcc.trim()) {
    throw new Error("Add at least one recipient.");
  }
  paneCancelSend(uri);
  paneState.drafts.set(uri, { ...draft, dirty: true });
  const seconds = Math.min(Math.max(Math.floor(Number(paneState.config.undoSendSeconds)) || 0, 0), 60);
  if (seconds <= 0) {
    return paneDeliverAndReport(uri, draft).then(() => ({ sent: true }));
  }
  const until = Date.now() + seconds * 1000;
  const timer = setTimeout(() => {
    paneState.pending.delete(uri);
    paneDeliverAndReport(uri, draft).catch(() => {});
  }, seconds * 1000);
  paneState.pending.set(uri, { timer, until, draft });
  return Promise.resolve({ pending: true, until });
}

function paneCancelSend(uri) {
  const pending = paneState.pending.get(uri);
  if (!pending) {
    return null;
  }
  clearTimeout(pending.timer);
  paneState.pending.delete(uri);
  return pending.draft;
}

function paneSendNow(uri) {
  const draft = paneCancelSend(uri);
  if (!draft) {
    return Promise.resolve({ sent: false });
  }
  return paneDeliverAndReport(uri, draft).then(() => ({ sent: true }));
}

async function paneFlushPending() {
  const uris = [...paneState.pending.keys()];
  await Promise.allSettled(uris.map((uri) => paneSendNow(uri)));
}

/* ------------------------------------------------------------------ */
/* Bridge object (used by reading-pane/pane.js)                        */
/* ------------------------------------------------------------------ */

const paneBridge = {
  config() {
    return {
      ...paneState.config,
      // Conversations' ◐ button (Thunderbird's dark message reader).
      darkReader: Services.prefs.getBoolPref(PANE_DARK_READER_PREF, false),
      loadId: PANE_LOAD_ID,
    };
  },
  uriForConvId(id) {
    const ext = conversationsExtension();
    try {
      return uriOfHdr(ext && ext.messageManager.get(Number(id)));
    } catch (error) {
      return null;
    }
  },
  // The date of a message of the conversation (ms), to order the messages:
  // Conversations appends a message that comes in while the conversation is
  // shown at the end of its list, whatever its date.
  dateForConvId(id) {
    try {
      const hdr = hdrFromUri(this.uriForConvId(id));
      return hdr ? Math.floor(hdr.date / 1000) : null;
    } catch (error) {
      return null;
    }
  },
  convIdForUri(uri) {
    const ext = conversationsExtension();
    const hdr = hdrFromUri(uri);
    try {
      return ext && hdr ? ext.messageManager.convert(hdr).id : null;
    } catch (error) {
      return null;
    }
  },
  // HTML for the pane's editor (a signature, an earlier draft), cleaned by
  // the sanitizer. Pictures shown through blob: URLs are put back as their
  // data: URLs first, so they survive; the pane makes blob: URLs again.
  sanitize(html) {
    const doc = paneParseHtml(`<html><body>${String(html || "")}</body></html>`);
    paneRestoreImages(doc);
    return paneSanitize(`<html><body>${doc.body.innerHTML}</body></html>`);
  },
  prepare(uri, type) {
    return panePrepare(uri, type).catch((error) => {
      paneNote("prepare", error);
      throw new Error(String(error && error.message ? error.message : error));
    });
  },
  searchAddresses(query) {
    try {
      return paneSearchAddresses(query);
    } catch (error) {
      paneNote("address search", error);
      return [];
    }
  },
  searchContacts(query) {
    try {
      return paneSearchContacts(query);
    } catch (error) {
      paneNote("mention search", error);
      return [];
    }
  },
  getDraft(uri) {
    return paneClone(paneState.drafts.get(uri) || null);
  },
  // A file attached in the pane: its path is kept here, the pane and its
  // drafts get a token for it.
  addFile(file) {
    const path = file && file.mozFullPath;
    if (!path) {
      return "";
    }
    for (const [id, known] of paneState.files) {
      if (known === path) {
        return id;
      }
    }
    const id = `file-${paneState.files.size + 1}`;
    paneState.files.set(id, path);
    return id;
  },
  saveDraft(uri, draft) {
    paneState.drafts.delete(uri);
    paneState.drafts.set(uri, paneClone(draft));
    if (paneState.drafts.size > PANE_MAX_DRAFTS) {
      // New messages stay while their dock tab does.
      const oldest = [...paneState.drafts.keys()].find((key) => !key.startsWith("new:"));
      paneState.drafts.delete(oldest);
    }
  },
  discardDraft(uri) {
    paneState.drafts.delete(uri);
  },
  send(uri, draft) {
    try {
      return paneSend(uri, draft);
    } catch (error) {
      return Promise.reject(new Error(String(error && error.message ? error.message : error)));
    }
  },
  pendingSend(uri) {
    const pending = paneState.pending.get(uri);
    return pending ? { until: pending.until, draft: paneClone(pending.draft) } : null;
  },
  cancelSend(uri) {
    return paneClone(paneCancelSend(uri));
  },
  sendNow(uri) {
    return paneSendNow(uri);
  },
  popOut(uri, draft) {
    draft = paneClone(draft);
    const hdr = draft.type == "new" ? null : hdrFromUri(uri);
    if ((draft.type != "new" && !hdr) || !paneState.extension || !paneState.listeners.popOut.size) {
      return false;
    }
    const messageId = hdr ? paneState.extension.messageManager.convert(hdr).id : null;
    const split = (text) => {
      try {
        return paneAddressList(text);
      } catch (error) {
        return String(text || "")
          .split(/[;,]/)
          .map((s) => s.trim())
          .filter(Boolean);
      }
    };
    for (const fire of paneState.listeners.popOut) {
      fire.async({
        messageId,
        type: draft.type,
        identityId: draft.identityKey,
        to: split(draft.to),
        cc: split(draft.cc),
        bcc: split(draft.bcc),
        subject: draft.subject || "",
        // The text written in the pane may hold pasted HTML: sanitized
        // like a signature before it goes into the compose window.
        html: draft.html ? paneSanitize(`<html><body>${draft.html}</body></html>`) : "",
      });
    }
    paneState.drafts.delete(uri);
    paneForgetNew(uri);
    return true;
  },
  // The composer's Signature menu → Preferences, at that section.
  openSettings(section) {
    for (const fire of paneState.listeners.openSettings) {
      fire.async(String(section || ""));
    }
  },
  // The Outlook-style To / Cc lines of a message in the conversation.
  messageDetails(convId) {
    const hdr = hdrFromUri(this.uriForConvId(convId));
    if (!hdr || !this.isShown(hdr)) {
      return null;
    }
    const list = (raw) => {
      try {
        return MailServices.headerParser
          .parseEncodedHeaderW(raw || "")
          .filter((address) => address.email)
          .map((address) => ({ name: address.name || "", email: address.email }));
      } catch (error) {
        return [];
      }
    };
    return {
      to: list(hdr.recipients),
      cc: list(hdr.ccList),
      bcc: list(hdr.bccList),
      time: Math.round(hdr.date / 1000),
    };
  },
  // Conversations' star is the pin button (pin.js pins the conversation).
  isPinned(convId) {
    const hdr = hdrFromUri(this.uriForConvId(convId));
    const keywords = hdr ? hdr.getStringProperty("keywords") || "" : "";
    return keywords.split(" ").includes(paneState.config.pinKeyword);
  },
  // The pin button pins the conversation the page shows: the one selected
  // in the message list of the mail tab the page is in. Any other id is
  // refused, so the page (which shows mail content) can pin nothing else.
  togglePin(convId) {
    const hdr = hdrFromUri(this.uriForConvId(convId));
    const shown = this.shownConversation();
    if (!hdr || !shown || !paneState.extension || !paneState.listeners.pin.size) {
      return false;
    }
    if (!this.isShown(hdr)) {
      return false;
    }
    const messageId = paneState.extension.messageManager.convert(shown[0]).id;
    for (const fire of paneState.listeners.pin) {
      fire.async(messageId);
    }
    return true;
  },
  reportError(where, text) {
    paneNote(`page: ${where}`, text);
  },
  // Whether hdr is mail of the conversation this page shows.
  isShown(hdr) {
    const shown = this.shownConversation();
    return Boolean(
      hdr && shown && shown.some((selected) => selected == hdr || sameThread(selected, hdr) || selected.messageId == hdr.messageId)
    );
  },
  // Saves the conversation this page shows as one .eml file, at a place
  // the user picks ("<subject> - <today>.eml"): a summary of its messages
  // (date, sender, subject; newest first), then every message in full as an
  // attached message (message/rfc822), oldest first. Thunderbird shows them
  // all in the message; Outlook shows the summary and each as an attachment.
  // Only messages of that conversation are saved: in the selected thread, or
  // linked to it by Message-ID and References (your replies in Sent Items).
  // Returns { saved, file } (saved 0: cancelled or nothing to save).
  async saveConversation(convIds) {
    const shown = this.shownConversation();
    const about3Pane = this.page?.browsingContext?.embedderElement?.ownerDocument?.defaultView;
    if (!shown || !about3Pane || !Array.isArray(convIds)) {
      return { saved: 0, file: "" };
    }
    const seen = new Set();
    const hdrs = [];
    for (const id of convIds) {
      const hdr = hdrFromUri(this.uriForConvId(id));
      if (hdr && !seen.has(hdr.messageId) && (this.isShown(hdr) || paneRelated(hdr, shown))) {
        seen.add(hdr.messageId);
        hdrs.push(hdr);
      }
    }
    if (!hdrs.length) {
      return { saved: 0, file: "" };
    }
    hdrs.sort((a, b) => a.date - b.date);
    const subject = paneThreadSubject(hdrs);
    const now = new Date();
    const path = await paneAskSavePath(about3Pane.browsingContext, paneConversationFileName(subject, now));
    if (!path) {
      return { saved: 0, file: "" };
    }
    const raws = [];
    for (const hdr of hdrs) {
      raws.push(await paneMessageBytes(uriOfHdr(hdr)));
    }
    await IOUtils.write(path, await paneConversationEml(hdrs, raws, subject, now));
    return { saved: hdrs.length, file: path };
  },
  // The mail selected in the list of the mail tab this page is in (one row:
  // one conversation), or null.
  shownConversation() {
    const browser = this.page?.browsingContext?.embedderElement;
    const about3Pane = browser?.ownerDocument?.defaultView;
    if (!about3Pane || about3Pane.multiMessageBrowser != browser || about3Pane.threadTree?.selectedIndices.length != 1) {
      return null;
    }
    const hdrs = about3Pane.gDBView?.getSelectedMsgHdrs() || [];
    return hdrs.length ? hdrs : null;
  },
};

// Whether a message belongs to the same conversation as one of the given
// ones by Message-ID and References (a reply in another folder, e.g. Sent
// Items).
function paneRelated(hdr, others) {
  const ids = (h) => {
    const set = new Set([h.messageId]);
    for (let i = 0; i < h.numReferences; i++) {
      set.add(h.getStringReference(i));
    }
    set.delete("");
    return set;
  };
  const mine = ids(hdr);
  return others.some((other) => [...ids(other)].some((id) => mine.has(id)));
}

/* ------------------------------------------------------------------ */
/* A conversation as one .eml file                                     */
/* ------------------------------------------------------------------ */

const PANE_SUBJECT_PREFIX = /^\s*((re|fw|fwd|aw|wg|tr|sv|vs|rv|antw|ynt|trả lời|chuyển tiếp|回复|答复|转发|返信|転送)\s*(\[\d+\])?\s*[:：]\s*)+/i;

// The conversation's subject: its oldest message's, without Re: / Fw:.
function paneThreadSubject(hdrs) {
  const subject = hdrs.map((hdr) => hdr.mime2DecodedSubject || "").find(Boolean) || "";
  return subject.replace(PANE_SUBJECT_PREFIX, "").trim() || "(no subject)";
}

function paneFileSafe(text, max = 150) {
  return String(text)
    .replace(/[\\/:*?"<>|\u0000-\u001f\u007f]+/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, max)
    .replace(/[. ]+$/, "");
}

// "<subject> - 2026-11-24.eml" (the day it is saved).
function paneConversationFileName(subject, now) {
  const pad = (n) => String(n).padStart(2, "0");
  const day = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
  return `${paneFileSafe(subject, 180) || "Conversation"} - ${day}.eml`;
}

// Thunderbird's file picker in save mode, in the folder last used to save
// mail. Returns the path, or "" when cancelled.
async function paneAskSavePath(browsingContext, name) {
  const fp = Cc["@mozilla.org/filepicker;1"].createInstance(Ci.nsIFilePicker);
  fp.init(browsingContext, "Save the conversation", Ci.nsIFilePicker.modeSave);
  fp.defaultString = name;
  fp.defaultExtension = "eml";
  fp.appendFilter("Mail (*.eml)", "*.eml");
  fp.appendFilters(Ci.nsIFilePicker.filterAll);
  try {
    fp.displayDirectory = Services.prefs.getComplexValue("messenger.save.dir", Ci.nsIFile);
  } catch (error) {
    // not set yet
  }
  const result = await new Promise((resolve) => fp.open(resolve));
  if (result == Ci.nsIFilePicker.returnCancel || !fp.file) {
    return "";
  }
  try {
    Services.prefs.setComplexValue("messenger.save.dir", Ci.nsIFile, fp.file.parent);
  } catch (error) {
    // remembering the folder is not needed
  }
  return /\.eml$/i.test(fp.file.path) ? fp.file.path : `${fp.file.path}.eml`;
}

// UTF-8 bytes of a string. The experiment scope has no TextEncoder (and
// one of a window would make arrays of that window).
function paneUtf8(text) {
  const out = [];
  for (const char of String(text)) {
    let code = char.codePointAt(0);
    if (code >= 0xd800 && code <= 0xdfff) {
      code = 0xfffd; // a lone surrogate, as TextEncoder does
    }
    if (code < 0x80) {
      out.push(code);
    } else if (code < 0x800) {
      out.push(0xc0 | (code >> 6), 0x80 | (code & 63));
    } else if (code < 0x10000) {
      out.push(0xe0 | (code >> 12), 0x80 | ((code >> 6) & 63), 0x80 | (code & 63));
    } else {
      out.push(0xf0 | (code >> 18), 0x80 | ((code >> 12) & 63), 0x80 | ((code >> 6) & 63), 0x80 | (code & 63));
    }
  }
  return new Uint8Array(out);
}

// A message as it is stored (downloaded first if needed), through
// Thunderbird's own "save as .eml" into a temporary file that is removed.
async function paneMessageBytes(uri) {
  const { MessageSaver } = ChromeUtils.importESModule("moz-src:///comm/mail/modules/MessageSaver.sys.mjs");
  const temp = Services.dirsvc.get("TmpD", Ci.nsIFile);
  temp.append("outlookmode-message.eml");
  temp.createUnique(Ci.nsIFile.NORMAL_FILE_TYPE, 0o600);
  const path = temp.path;
  try {
    await MessageSaver.saveToFile(uri, path, "eml");
    let bytes = await IOUtils.read(path);
    // An mbox envelope line ("From - ...") is not part of the message.
    if (bytes.length > 5 && String.fromCharCode(...bytes.subarray(0, 5)) == "From ") {
      const end = bytes.indexOf(10);
      bytes = end >= 0 ? bytes.subarray(end + 1) : bytes;
    }
    return bytes;
  } finally {
    await IOUtils.remove(path, { ignoreAbsent: true }).catch(() => {});
  }
}

// Header text, RFC 2047 encoded when not plain ASCII. fold: false keeps
// the encoded words on one line (inside a quoted parameter).
function paneEncodeHeader(text, fold = true) {
  text = String(text || "").replace(/[\r\n]+/g, " ");
  if (/^[\x20-\x7e]*$/.test(text)) {
    return text;
  }
  const words = [];
  let chunk = "";
  for (const char of text) {
    if (paneUtf8(chunk + char).length > 45) {
      words.push(chunk);
      chunk = "";
    }
    chunk += char;
  }
  words.push(chunk);
  return words.map((word) => `=?UTF-8?B?${paneBase64(paneUtf8(word))}?=`).join(fold ? "\r\n " : " ");
}

function paneWrap76(text) {
  return text.replace(/.{1,76}/g, "$&\r\n");
}

// A stored message, parsed with Thunderbird's MIME parser: its headers, its
// text as HTML, its pictures by Content-ID and its attachments. A message
// attached to it is kept whole, as one attachment (decodeSubMessages off).
function paneParseMessage(raw) {
  const { MimeParser } = ChromeUtils.importESModule("resource:///modules/mimeParser.sys.mjs");
  let binary = "";
  for (let i = 0; i < raw.length; i += 0x8000) {
    binary += String.fromCharCode(...raw.subarray(i, i + 0x8000));
  }
  const parts = new Map(); // part number -> { headers, chunks }
  let top = null;
  MimeParser.parseSync(
    binary,
    {
      startPart(partNum, headers) {
        if (partNum === "") {
          top = headers;
        }
        parts.set(partNum, { partNum, headers, chunks: [] });
      },
      deliverPartData(partNum, data) {
        parts.get(partNum)?.chunks.push(data);
      },
    },
    { bodyformat: "decode", strformat: "unicode", decodeSubMessages: false }
  );
  const message = { headers: top, html: "", text: "", images: new Map(), attachments: [] };
  for (const part of parts.values()) {
    const { partNum, headers } = part;
    const type = headers.contentType || { mediatype: "text", subtype: "plain", type: "text/plain", get: () => "" };
    if (type.mediatype == "multipart") {
      continue;
    }
    const disposition = String(headers.get("content-disposition") || "");
    const isAttachment = /^\s*attachment/i.test(disposition);
    const name =
      MimeParser.getParameter(disposition, "filename") || (type.get && type.get("name")) || "";
    const join = () =>
      typeof part.chunks[0] == "string" ? part.chunks.join("") : paneConcat(part.chunks);
    if (type.mediatype == "message" && partNum !== "") {
      // Named after its own subject when it has no file name.
      const data = paneBytesOf(join());
      let title = name;
      if (!title) {
        try {
          let head = "";
          for (let i = 0; i < Math.min(data.length, 65536); i += 0x8000) {
            head += String.fromCharCode(...data.subarray(i, Math.min(i + 0x8000, 65536)));
          }
          const subject = MimeParser.extractHeaders(head).get("Subject");
          title = subject ? `${paneFileSafe(subject, 120)}.eml` : "";
        } catch (error) {
          // no usable subject
        }
      }
      message.attachments.push({ type: type.type, name: title || "message.eml", data });
      continue;
    }
    const cid = String(headers.get("content-id") || "").replace(/^\s*<|>\s*$/g, "");
    if (!isAttachment && type.type == "text/html" && !message.html) {
      message.html = join();
    } else if (!isAttachment && type.type == "text/plain" && !message.text && !name) {
      message.text = join();
    } else if (cid && !isAttachment) {
      message.images.set(cid, { type: type.type, name, data: paneBytesOf(join()) });
    } else if (partNum !== "") {
      message.attachments.push({ type: type.type, name: name || "attachment", data: paneBytesOf(join()) });
    }
  }
  return message;
}

function paneConcat(chunks) {
  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const out = new Uint8Array(total);
  let at = 0;
  for (const chunk of chunks) {
    out.set(chunk, at);
    at += chunk.length;
  }
  return out;
}

// Bytes of a part: text parts come as strings (Unicode), others as bytes.
function paneBytesOf(data) {
  return typeof data == "string" ? paneUtf8(data) : data;
}

// The first line of Outlook's quoted header ("From:" in the mail's language).
const PANE_FROM_LABEL = /^\s*(From|Từ|De|Von|Da|Van|Od|От|Från|Fra|Kimden|发件人|發件人|寄件者|差出人|보낸 사람)\s*[:：]/;

// Where the earlier mail a message quotes starts, as in the reading pane
// (reading-pane/pane.js, quoteStart): [node, toEnd] or null.
function paneQuoteStart(doc) {
  const body = doc.body;
  const found = [];
  for (const sel of ["#appendonsend", "#divRplyFwdMsg"]) {
    const node = body.querySelector(sel);
    if (node) {
      found.push([node, true]);
    }
  }
  const header = [...body.querySelectorAll("div[style*='border-top']")].find((node) => PANE_FROM_LABEL.test(node.textContent));
  if (header) {
    found.push([header, true]);
  }
  const rule = [...body.querySelectorAll("hr")].find(
    (node) => node.nextElementSibling && PANE_FROM_LABEL.test(node.nextElementSibling.textContent)
  );
  if (rule) {
    found.push([rule, true]);
  }
  for (const sel of [".gmail_quote", ".yahoo_quoted", ".moz-cite-prefix", ".moz-forward-container", "blockquote[type='cite']"]) {
    const node = body.querySelector(sel);
    if (node) {
      found.push([node, false]);
    }
  }
  found.sort(([a], [b]) => (a === b ? 0 : a.compareDocumentPosition(b) & 4 ? -1 : 1));
  return found[0] || null;
}

// A message's own part as HTML: its text without the earlier mail it quotes
// (keepQuote: the oldest message keeps it), without scripts, frames, forms
// or event handlers, its pictures renamed to `${prefix}-...` Content-IDs.
// Returns { html, styles, cids } (cids: original Content-IDs still shown).
function paneOwnPart(message, keepQuote, prefix) {
  let html = message.html;
  if (!html) {
    const text = keepQuote ? message.text || "" : paneOwnText(message.text || "");
    html = `<pre style="white-space: pre-wrap; font-family: inherit;">${paneEscape(text)}</pre>`;
  }
  const doc = paneParseHtml(html);
  for (const node of doc.querySelectorAll("script, iframe, frame, frameset, object, embed, form, base, meta, link, title")) {
    node.remove();
  }
  if (!keepQuote) {
    const start = paneQuoteStart(doc);
    if (start) {
      const [node, toEnd] = start;
      const range = doc.createRange();
      range.setStart(doc.body, 0);
      range.setEndBefore(node);
      const own = range.toString().trim() || range.cloneContents().querySelector("img");
      if (own) {
        // A forward with no text of its own keeps what it forwards.
        if (toEnd) {
          const cut = doc.createRange();
          cut.setStartBefore(node);
          cut.setEndAfter(doc.body.lastChild);
          cut.deleteContents();
        } else {
          const last = node.matches(".moz-cite-prefix") && node.nextElementSibling ? node.nextElementSibling : node;
          const cut = doc.createRange();
          cut.setStartBefore(node);
          cut.setEndAfter(last);
          cut.deleteContents();
        }
      }
    }
  }
  const styles = [...doc.querySelectorAll("style")].map((node) => {
    const css = node.textContent;
    node.remove();
    return css;
  });
  const cids = new Set();
  for (const node of doc.querySelectorAll("*")) {
    for (const attr of [...node.attributes]) {
      const value = attr.value.trim();
      if (/^on/i.test(attr.name) || /^\s*(javascript|vbscript):/i.test(value)) {
        node.removeAttribute(attr.name);
      } else if (/^cid:/i.test(value) && ["src", "background", "href"].includes(attr.name.toLowerCase())) {
        const cid = decodeURIComponent(value.slice(4));
        if (message.images.has(cid)) {
          cids.add(cid);
          node.setAttribute(attr.name, `cid:${prefix}-${cid}`);
        }
      }
    }
  }
  return { html: doc.body.innerHTML, styles, cids };
}

// Plain text without the earlier mail it quotes: from Outlook's underscore
// line followed by "From:", or a closing block of "> " lines with the
// "... wrote:" line above it. Text that is nothing but a quote stays.
function paneOwnText(text) {
  let own = text;
  const outlook = /^[ \t]*_{20,}[ \t]*\r?\n(?=[ \t]*\S)/m.exec(text);
  if (outlook && PANE_FROM_LABEL.test(text.slice(outlook.index + outlook[0].length))) {
    own = text.slice(0, outlook.index);
  } else {
    const quoted = /(?:^.*(?:wrote|viết|schrieb|écrit|escribió):[ \t]*\r?\n)?(?:^>.*(?:\r?\n|(?![\s\S]))|^[ \t]*\r?\n)*(?![\s\S])/im.exec(text);
    if (quoted && quoted[0].includes(">")) {
      own = text.slice(0, quoted.index);
    }
  }
  return own.trim() ? own.replace(/\s+$/, "") : text;
}

function paneAddressText(list) {
  const flat = [];
  for (const item of list || []) {
    if (item.group) {
      flat.push(...item.group);
    } else {
      flat.push(item);
    }
  }
  return flat.map((a) => (a.name ? `${a.name} <${a.email}>` : a.email)).join("; ");
}

// A part of the .eml file: headers, then the data in base64.
function panePartBytes(headerLines, data) {
  return [paneUtf8(`${headerLines.join("\r\n")}\r\nContent-Transfer-Encoding: base64\r\n\r\n`), paneUtf8(paneWrap76(paneBase64(data)))];
}

function paneNameParams(name) {
  return `name="${paneEncodeHeader(name, false)}"`;
}

function paneDispositionParams(kind, name) {
  return `${kind}; filename="${paneEncodeHeader(name, false)}"; filename*=UTF-8''${encodeURIComponent(name)}`;
}

// The .eml file: the whole conversation as the message's own text, newest
// first, each message under an Outlook-style header (From, Sent, To, Cc,
// Subject) and without the earlier mail it quotes (the oldest keeps its
// quote), with the pictures it shows; then the messages' attachments (the
// same file once).
//   multipart/mixed
//     multipart/related: text/html + pictures (Content-ID)
//     attachments
async function paneConversationEml(hdrs, raws, subject, now) {
  const boundary = `----=_OutlookMode_${Date.now().toString(36)}_${Math.random().toString(36).slice(2)}`;
  const related = `${boundary}_rel`;
  let from = "";
  try {
    const identity = MailServices.accounts.getFirstIdentityForServer(hdrs[0].folder.server);
    // Only the name is encoded; the address stays as it is.
    const name = identity && identity.fullName ? identity.fullName : "";
    from = !identity
      ? ""
      : /^[\x20-\x7e]*$/.test(name)
        ? MailServices.headerParser.makeMimeAddress(name, identity.email)
        : `${paneEncodeHeader(name)} <${identity.email}>`;
  } catch (error) {
    // no From line
  }

  const sections = [];
  const styles = new Set();
  const pictures = [];
  const attachments = [];
  const seenFiles = new Set();
  const usedNames = new Set();
  for (let index = hdrs.length - 1; index >= 0; index--) {
    // Newest first; a pause between messages keeps Thunderbird responsive.
    await new Promise((resolve) => setTimeout(resolve, 0));
    const hdr = hdrs[index];
    let message;
    try {
      message = paneParseMessage(raws[index]);
    } catch (error) {
      paneNote("save: parse message", error);
      message = { headers: null, html: "", text: "(This message could not be read.)", images: new Map(), attachments: [] };
    }
    const prefix = `m${index + 1}`;
    const own = paneOwnPart(message, index === 0, prefix);
    own.styles.forEach((css) => styles.add(css));
    for (const cid of own.cids) {
      const image = message.images.get(cid);
      pictures.push({ cid: `${prefix}-${cid}`, ...image });
    }
    for (const attachment of message.attachments) {
      const key = `${attachment.name}\u0000${attachment.data.length}`;
      if (!seenFiles.has(key)) {
        seenFiles.add(key);
        // Different files of the same name: "name (2).ext".
        let name = attachment.name;
        const dot = name.lastIndexOf(".");
        const [stem, ext] = dot > 0 ? [name.slice(0, dot), name.slice(dot)] : [name, ""];
        for (let n = 2; usedNames.has(name.toLowerCase()); n++) {
          name = `${stem} (${n})${ext}`;
        }
        usedNames.add(name.toLowerCase());
        attachments.push({ ...attachment, name });
      }
    }
    const headers = message.headers;
    const get = (name) => {
      try {
        return headers ? headers.get(name) : null;
      } catch (error) {
        return null;
      }
    };
    const fromText = paneAddressText(get("From")) || hdr.mime2DecodedAuthor || "";
    const toText = paneAddressText(get("To"));
    const ccText = paneAddressText(get("Cc"));
    const line = (label, value) => (value ? `<b>${label}:</b> ${paneEscape(value)}<br>` : "");
    sections.push(
      '<div style="border:none;border-top:solid #E1E1E1 1.0pt;padding:6pt 0 0 0;margin:18pt 0 8pt 0">' +
        '<p style="margin:0 0 8pt 0;font-size:11pt">' +
        line("From", fromText) +
        line("Sent", paneOutlookDate(new Date(hdr.date / 1000))) +
        line("To", toText) +
        line("Cc", ccText) +
        line("Subject", hdr.flags & Ci.nsMsgMessageFlags.HasRe ? `RE: ${hdr.mime2DecodedSubject}` : hdr.mime2DecodedSubject) +
        "</p></div>" +
        `<div>${own.html}</div>`
    );
  }

  const html =
    '<!DOCTYPE html><html><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">' +
    [...styles].map((css) => `<style>${css}</style>`).join("") +
    "</head>" +
    `<body style="font-family: ${paneEscape(PANE_DEFAULT_FONT)}; font-size: 11pt;">` +
    `<p style="font-size:14pt;margin:0"><b>${paneEscape(subject)}</b></p>` +
    `<p style="color:#666;margin:2pt 0 0 0">${hdrs.length} message${hdrs.length == 1 ? "" : "s"}, newest first. Saved ${paneEscape(paneOutlookDate(now))}.</p>` +
    sections.join("") +
    "</body></html>";

  const out = [
    paneUtf8(
      [
        from ? `From: ${from}` : "",
        `Subject: ${paneEncodeHeader(subject)}`,
        `Date: ${now.toUTCString().replace("GMT", "+0000")}`,
        "MIME-Version: 1.0",
        "X-Outlook-Mode: conversation saved from Thunderbird",
        `Content-Type: multipart/mixed; boundary="${boundary}"`,
        "",
        "This is a conversation saved as one message.",
        "",
        `--${boundary}`,
        `Content-Type: multipart/related; type="text/html"; boundary="${related}"`,
        "",
        `--${related}`,
        "",
      ]
        .filter((text, index) => text || index > 0)
        .join("\r\n")
    ),
  ];
  out.push(...panePartBytes(["Content-Type: text/html; charset=UTF-8"], paneUtf8(html)));
  for (const picture of pictures) {
    out.push(paneUtf8(`\r\n--${related}\r\n`));
    out.push(
      ...panePartBytes(
        [
          `Content-Type: ${picture.type || "application/octet-stream"}${picture.name ? `; ${paneNameParams(picture.name)}` : ""}`,
          `Content-ID: <${picture.cid}>`,
          picture.name ? `Content-Disposition: ${paneDispositionParams("inline", picture.name)}` : "Content-Disposition: inline",
        ],
        picture.data
      )
    );
  }
  out.push(paneUtf8(`\r\n--${related}--\r\n`));
  for (const attachment of attachments) {
    out.push(paneUtf8(`\r\n--${boundary}\r\n`));
    if (/^message\//i.test(attachment.type)) {
      // An attached mail goes in as it is: MIME allows no base64 there.
      const eightBit = attachment.data.some((byte) => byte > 127);
      out.push(
        paneUtf8(
          `Content-Type: ${attachment.type}; ${paneNameParams(attachment.name)}\r\n` +
            `Content-Transfer-Encoding: ${eightBit ? "8bit" : "7bit"}\r\n` +
            `Content-Disposition: ${paneDispositionParams("attachment", attachment.name)}\r\n\r\n`
        ),
        attachment.data
      );
      if (attachment.data.length && attachment.data[attachment.data.length - 1] != 10) {
        out.push(paneUtf8("\r\n"));
      }
      continue;
    }
    out.push(
      ...panePartBytes(
        [
          `Content-Type: ${attachment.type || "application/octet-stream"}; ${paneNameParams(attachment.name)}`,
          `Content-Disposition: ${paneDispositionParams("attachment", attachment.name)}`,
        ],
        attachment.data
      )
    );
  }
  out.push(paneUtf8(`\r\n--${boundary}--\r\n`));
  return paneConcat(out);
}

// Whether two messages are in the same thread of the same folder.
function sameThread(a, b) {
  try {
    if (a.folder != b.folder) {
      return false;
    }
    const db = a.folder.msgDatabase;
    return db.getThreadContainingMsgHdr(a)?.threadKey == db.getThreadContainingMsgHdr(b)?.threadKey;
  } catch (error) {
    return false;
  }
}

/* ------------------------------------------------------------------ */
/* Pages: inject, configure, remove                                    */
/* ------------------------------------------------------------------ */

function paneInject(win) {
  try {
    if (!win || win.closed || win.outlookModePane || !paneState.config.enabled) {
      return;
    }
    if (!String(win.document.documentURI).startsWith(PANE_STUB_URL)) {
      return;
    }
    // A bridge of its own, which knows its page (see togglePin).
    const bridge = Object.create(paneBridge, { page: { value: win, writable: false, enumerable: false } });
    win.outlookModeBridge = bridge;
    Services.scriptloader.loadSubScript(
      `chrome://${PANE_PACKAGE}/content/pane.js?${PANE_LOAD_ID}`,
      win
    );
    if (!win.outlookModePane) {
      return;
    }
    paneState.windows.add(win);
    win.addEventListener("unload", () => paneState.windows.delete(win), { once: true });
    // multiMessageBrowser is a content browser, so frameElement is null here.
    const browser = win.browsingContext?.embedderElement || win.frameElement;
    paneRouteCommands(browser && browser.ownerDocument.defaultView);
  } catch (error) {
    paneNote("inject", error);
  }
}

function paneScanOpen() {
  for (const win of Services.wm.getEnumerator("mail:3pane")) {
    const tabmail = win.document.getElementById("tabmail");
    for (const tab of (tabmail && tabmail.tabInfo) || []) {
      const about3Pane = tab.chromeBrowser && tab.chromeBrowser.contentWindow;
      const browser = about3Pane && about3Pane.multiMessageBrowser;
      const stub = browser && browser.contentWindow;
      if (stub && !stub.closed) {
        // Still loading at startup: created before the observer was added.
        if (stub.document.readyState == "loading") {
          stub.addEventListener("DOMContentLoaded", () => paneInject(stub), { once: true });
        } else {
          paneInject(stub);
        }
      }
    }
  }
}

const paneObserver = {
  observe(subject, topic) {
    if (topic == "document-element-inserted") {
      const doc = subject;
      const url = String(doc.documentURI || "");
      const win = doc.defaultView;
      if (!win) {
        return;
      }
      if (url.startsWith(PANE_STUB_URL)) {
        win.addEventListener("DOMContentLoaded", () => paneInject(win), { once: true });
      } else if (url == "about:message") {
        win.addEventListener("DOMContentLoaded", () => paneStyleMessageView(win), { once: true });
      } else if (url.startsWith(PANE_DOCK_URL)) {
        // The iframe's load event does not reach chrome listeners here.
        win.addEventListener(
          "DOMContentLoaded",
          () => paneState.dock?.frame.contentWindow == win && paneInjectDock(paneState.dock),
          { once: true }
        );
      }
    } else if (topic == "quit-application-requested") {
      paneQuitRequested(subject);
    }
  },
};

// Unsent inline replies are not lost silently when Thunderbird quits.
function paneQuitRequested(subject) {
  const cancel = subject.QueryInterface(Ci.nsISupportsPRBool);
  if (cancel.data) {
    return;
  }
  if (paneState.pending.size) {
    // Messages in their Undo Send countdown go out now, then quit again.
    cancel.data = true;
    paneFlushPending().finally(() => Services.startup.quit(Ci.nsIAppStartup.eAttemptQuit));
    return;
  }
  const unsent = [...paneState.drafts.values()].filter((d) => d.dirty).length;
  if (
    unsent &&
    !Services.prompt.confirm(
      null,
      "Outlook Mode",
      `You have ${unsent} unsent ${unsent == 1 ? "message" : "messages"} in the reading pane or the compose dock. Quit anyway and discard ${unsent == 1 ? "it" : "them"}?`
    )
  ) {
    cancel.data = true;
  }
}

// Without Conversations, Thunderbird shows a message in about:message. Its
// header takes its colours from light-dark() pairs (layout.css), while the
// message text follows the ◐ toggle (messageBody.css: dark only with
// mail.dark-reader.enabled). So in a dark theme with ◐ off, the header is
// dark above white mail. Like the Conversations page, the header follows ◐:
// white while it is off, dark while it is on.
function paneStyleMessageView(win, remove = false) {
  try {
    const root = win.document.documentElement;
    if (!remove && paneState.config.enabled && paneState.config.lightPane) {
      root.style.colorScheme = Services.prefs.getBoolPref(PANE_DARK_READER_PREF, false) ? "dark" : "light";
    } else {
      root.style.removeProperty("color-scheme");
    }
  } catch (error) {
    // the window is closing
  }
}

// Every about:message: the 3-pane tabs' message pane, message tabs and
// message windows.
function paneStyleMessageViews(remove = false) {
  for (const win of Services.wm.getEnumerator(null)) {
    const views = [];
    if (win.document.documentURI == "about:message") {
      views.push(win);
    }
    if (win.messageBrowser?.contentWindow) {
      views.push(win.messageBrowser.contentWindow); // a message window
    }
    for (const tab of win.document.getElementById("tabmail")?.tabInfo || []) {
      const content = tab.chromeBrowser?.contentWindow;
      if (content?.document.documentURI == "about:message") {
        views.push(content); // a message tab
      } else if (content?.messageBrowser?.contentWindow) {
        views.push(content.messageBrowser.contentWindow); // a 3-pane tab
      }
    }
    for (const view of views) {
      if (view.document.documentURI == "about:message") {
        paneStyleMessageView(view, remove);
      }
    }
  }
}

// The ◐ button (Conversations' or Thunderbird's own): the open panes follow it.
const paneDarkReaderObserver = {
  observe() {
    for (const win of paneState.windows) {
      try {
        win.outlookModePane?.configChanged();
      } catch (error) {
        // the page is going away
      }
    }
    paneStyleMessageViews();
  },
};

function paneStartObserving() {
  if (!paneState.observing) {
    Services.obs.addObserver(paneObserver, "document-element-inserted");
    Services.obs.addObserver(paneObserver, "quit-application-requested");
    Services.prefs.addObserver(PANE_DARK_READER_PREF, paneDarkReaderObserver);
    paneState.observing = true;
  }
}

function paneStopObserving() {
  if (paneState.observing) {
    Services.obs.removeObserver(paneObserver, "document-element-inserted");
    Services.obs.removeObserver(paneObserver, "quit-application-requested");
    Services.prefs.removeObserver(PANE_DARK_READER_PREF, paneDarkReaderObserver);
    paneState.observing = false;
  }
}

function paneRemoveAll() {
  for (const win of paneState.windows) {
    try {
      win.outlookModePane?.destroy();
      delete win.outlookModePane;
      delete win.outlookModeBridge;
    } catch (error) {
      // the page is going away
    }
  }
  paneState.windows.clear();
  for (const about3Pane of [...paneState.routed.keys()]) {
    paneUnrouteCommands(about3Pane);
  }
  paneRemoveDock();
  paneStopComposeRouting();
  paneStyleMessageViews(true); // back to Thunderbird's colours
}

/* ------------------------------------------------------------------ */
/* Thunderbird's Reply / Reply all / Forward commands                  */
/* ------------------------------------------------------------------ */

// Opens the inline composer for the selected message when the conversation
// view is showing it. Returns false to let Thunderbird open a compose window.
function paneOpenInline(about3Pane, type) {
  if (!paneState.config.enabled || !paneState.config.inlineReply) {
    return false;
  }
  try {
    const browser = about3Pane.multiMessageBrowser;
    const pane = browser && !browser.hidden && browser.contentWindow?.outlookModePane;
    const view = about3Pane.gDBView;
    if (!pane || !view || view.numSelected != 1) {
      return false;
    }
    const hdr = view.hdrForFirstSelectedMessage;
    if (!hdr) {
      return false;
    }
    // A collapsed thread row: Outlook replies to the newest message.
    let newest = false;
    try {
      const index = about3Pane.threadTree.selectedIndex;
      newest = view.isContainer(index) && !view.isContainerOpen(index);
    } catch (error) {
      // keep the selected message
    }
    return Boolean(pane.openFor(uriOfHdr(hdr), type, { newest }));
  } catch (error) {
    paneNote("route command", error);
    return false;
  }
}

function paneRouteCommands(about3Pane) {
  if (!about3Pane || paneState.routed.has(about3Pane)) {
    return;
  }
  const controller = about3Pane.commandController;
  if (!controller || !controller._composeMsgByType || !controller._callbackCommands?.cmd_forward) {
    return;
  }
  const original = {
    controller,
    compose: controller._composeMsgByType,
    forward: controller._callbackCommands.cmd_forward,
  };
  controller._composeMsgByType = function (compType, ...args) {
    const type =
      compType == Ci.nsIMsgCompType.ReplyToSender
        ? "reply"
        : compType == Ci.nsIMsgCompType.ReplyAll
          ? "replyAll"
          : compType == Ci.nsIMsgCompType.ForwardInline
            ? "forward"
            : null;
    if (type && paneOpenInline(about3Pane, type)) {
      return undefined;
    }
    return original.compose.call(this, compType, ...args);
  };
  // Forward follows mail.forward_message_mode (attachment by default); in
  // the reading pane it forwards inline, like Outlook.
  controller._callbackCommands.cmd_forward = function (...args) {
    if (paneOpenInline(about3Pane, "forward")) {
      return undefined;
    }
    return original.forward.apply(this, args);
  };
  paneState.routed.set(about3Pane, original);
  about3Pane.addEventListener("unload", () => paneState.routed.delete(about3Pane), { once: true });
}

function paneUnrouteCommands(about3Pane) {
  const original = paneState.routed.get(about3Pane);
  paneState.routed.delete(about3Pane);
  if (!original) {
    return;
  }
  try {
    original.controller._composeMsgByType = original.compose;
    original.controller._callbackCommands.cmd_forward = original.forward;
  } catch (error) {
    // the window is closing
  }
}

/* ------------------------------------------------------------------ */
/* New messages: the compose dock                                      */
/* ------------------------------------------------------------------ */

function paneCreateNew(identity) {
  const uri = `new:${Date.now().toString(36)}-${++paneState.newSeq}`;
  paneState.drafts.set(uri, {
    type: "new",
    identityKey: identity ? identity.key : "",
    to: "",
    cc: "",
    bcc: "",
    subject: "",
    html: null,
    attachments: [],
    files: [],
    dirty: false,
  });
  paneState.newOrder.push(uri);
  return uri;
}

function paneForgetNew(uri) {
  paneState.newOrder = paneState.newOrder.filter((key) => key != uri);
}

function paneDockHeight(dock) {
  const full = dock.about3Pane.innerHeight;
  if (dock.hidden) {
    return 0;
  }
  if (dock.minimized) {
    return PANE_DOCK_BAR;
  }
  if (dock.maximized) {
    return full;
  }
  return Math.max(PANE_DOCK_BAR + 160, Math.min(dock.height, full - 80));
}

// The dock is fixed to the bottom of about:3pane; the panes above it get the
// rest of the height, so nothing is covered.
function paneDockLayout(dock) {
  const height = paneDockHeight(dock);
  dock.box.style.display = dock.hidden ? "none" : "block";
  dock.box.style.height = `${height}px`;
  dock.about3Pane.document.body.style.height = height ? `calc(100vh - ${height}px)` : "";
}

function paneDockApi(dock) {
  const relayout = () => {
    paneDockLayout(dock);
  };
  return {
    // The message to show when the page starts or is told something changed.
    takeFocus() {
      const uri = dock.focusUri;
      dock.focusUri = null;
      return uri;
    },
    messages() {
      return paneState.newOrder.map((uri) => {
        const draft = paneState.drafts.get(uri) || {};
        return {
          uri,
          subject: draft.subject || "",
          dirty: Boolean(draft.dirty),
          sending: paneState.pending.has(uri),
        };
      });
    },
    create() {
      return paneCreateNew(paneIdentityForFolder(dock.about3Pane.gFolder));
    },
    remove(uri) {
      if (!paneState.pending.has(uri)) {
        paneState.drafts.delete(uri);
      }
      paneForgetNew(uri);
    },
    state() {
      return { minimized: dock.minimized, maximized: dock.maximized };
    },
    height() {
      return paneDockHeight(dock);
    },
    setMinimized(on) {
      dock.minimized = Boolean(on);
      if (dock.minimized) {
        dock.maximized = false;
      }
      relayout();
    },
    setMaximized(on) {
      dock.maximized = Boolean(on);
      if (dock.maximized) {
        dock.minimized = false;
      }
      relayout();
    },
    setHeight(px) {
      dock.height = Math.round(Number(px) || 0);
      dock.minimized = false;
      dock.maximized = false;
      relayout();
      Services.prefs.setIntPref(PANE_DOCK_HEIGHT_PREF, paneDockHeight(dock));
    },
    hide() {
      dock.hidden = true;
      relayout();
    },
  };
}

function paneInjectDock(dock) {
  try {
    const win = dock.frame.contentWindow;
    if (!win || win.outlookModePane || !String(win.document.documentURI).startsWith(PANE_DOCK_URL)) {
      return;
    }
    const bridge = Object.create(paneBridge);
    bridge.dock = paneDockApi(dock);
    win.outlookModeBridge = bridge;
    Services.scriptloader.loadSubScript(`chrome://${PANE_PACKAGE}/content/pane.js?${PANE_LOAD_ID}`, win);
    if (!win.outlookModePane) {
      return;
    }
    dock.page = win;
    paneState.windows.add(win);
    win.addEventListener(
      "unload",
      () => {
        paneState.windows.delete(win);
        if (dock.page == win) {
          dock.page = null;
        }
      },
      { once: true }
    );
    win.outlookModePane.dockChanged();
  } catch (error) {
    paneNote("inject dock", error);
  }
}

function paneCreateDock(about3Pane) {
  const doc = about3Pane.document;
  const box = doc.createElement("div");
  box.id = "outlookModeDock";
  box.style.cssText =
    "position: fixed; inset-inline: 0; bottom: 0; z-index: 10; overflow: hidden;" +
    "border-top: 1px solid var(--splitter-color, ThreeDShadow);" +
    "box-shadow: 0 -2px 8px rgba(0, 0, 0, 0.15); background: Canvas;";
  const frame = doc.createElement("iframe");
  frame.style.cssText = "display: block; width: 100%; height: 100%; border: 0;";
  box.appendChild(frame);
  const dock = {
    about3Pane,
    box,
    frame,
    page: null,
    hidden: false,
    minimized: false,
    maximized: false,
    height:
      Services.prefs.getIntPref(PANE_DOCK_HEIGHT_PREF, 0) || Math.round(about3Pane.innerHeight * 0.55),
    focusUri: null,
    onResize: null,
  };
  dock.onResize = () => paneDockLayout(dock);
  about3Pane.addEventListener("resize", dock.onResize);
  about3Pane.addEventListener(
    "unload",
    () => {
      if (paneState.dock == dock) {
        paneState.dock = null;
      }
    },
    { once: true }
  );
  doc.body.appendChild(box);
  frame.src = `${PANE_DOCK_URL}?${PANE_LOAD_ID}`;
  return dock;
}

function paneRemoveDock() {
  const dock = paneState.dock;
  paneState.dock = null;
  if (!dock) {
    return;
  }
  try {
    dock.page?.outlookModePane?.destroy();
    paneState.windows.delete(dock.page);
    dock.about3Pane.removeEventListener("resize", dock.onResize);
    dock.about3Pane.document.body.style.height = "";
    dock.box.remove();
  } catch (error) {
    // the window is closing
  }
}

// The 3-pane tab holding the dock, or the one to put it in.
function paneDockTab(win) {
  const dock = paneState.dock;
  if (dock) {
    for (const each of Services.wm.getEnumerator("mail:3pane")) {
      const tabmail = each.document.getElementById("tabmail");
      const tab = tabmail?.tabInfo.find((t) => t.chromeBrowser?.contentWindow == dock.about3Pane);
      if (tab && dock.box.isConnected) {
        return { win: each, tabmail, tab };
      }
    }
    paneRemoveDock();
  }
  const tabmail = win.document.getElementById("tabmail");
  if (!tabmail) {
    return null;
  }
  const isMail = (t) => t?.mode?.name == "mail3PaneTab";
  const tab = isMail(tabmail.currentTabInfo) ? tabmail.currentTabInfo : tabmail.tabInfo.find(isMail);
  return tab ? { win, tabmail, tab } : null;
}

// Opens a new message in the dock. Returns false to let Thunderbird open a
// compose window.
function paneNewMessage(win, folder) {
  const { config } = paneState;
  if (!config.enabled || !config.inlineNew) {
    return false;
  }
  try {
    if (folder && !folder.isServer && folder.server.type == "nntp") {
      return false; // a news post
    }
    const target = paneDockTab(win);
    if (!target) {
      return false;
    }
    if (target.tabmail.currentTabInfo != target.tab) {
      target.tabmail.switchToTab(target.tab);
    }
    const about3Pane = target.tab.chromeBrowser.contentWindow;
    if (!about3Pane || about3Pane.document.readyState != "complete") {
      return false;
    }
    if (target.win != win) {
      target.win.focus();
    }
    const uri = paneCreateNew(paneIdentityForFolder(folder || about3Pane.gFolder));
    const dock = paneState.dock || (paneState.dock = paneCreateDock(about3Pane));
    dock.hidden = false;
    dock.minimized = false;
    dock.focusUri = uri;
    paneDockLayout(dock);
    dock.page?.outlookModePane?.dockChanged();
    return true;
  } catch (error) {
    paneNote("new message", error);
    return false;
  }
}

// Every way to write a new message in the main window (Write button, Ctrl+N,
// File > New > Message, the folder pane's button) ends in ComposeMessage.
function paneRouteCompose(win) {
  if (!win || paneState.composeRouted.has(win) || typeof win.ComposeMessage != "function") {
    return;
  }
  const original = win.ComposeMessage;
  win.ComposeMessage = function (type, format, folder, ...rest) {
    if (type == Ci.nsIMsgCompType.New && paneNewMessage(win, folder)) {
      return Promise.resolve();
    }
    return original.call(this, type, format, folder, ...rest);
  };
  paneState.composeRouted.set(win, original);
  win.addEventListener("unload", () => paneState.composeRouted.delete(win), { once: true });
}

function paneStartComposeRouting() {
  if (paneState.windowListener) {
    return;
  }
  ExtensionSupport.registerWindowListener(PANE_WINDOW_LISTENER, {
    chromeURLs: ["chrome://messenger/content/messenger.xhtml"],
    onLoadWindow: paneRouteCompose,
  });
  paneState.windowListener = true;
}

function paneStopComposeRouting() {
  if (paneState.windowListener) {
    ExtensionSupport.unregisterWindowListener(PANE_WINDOW_LISTENER);
    paneState.windowListener = false;
  }
  for (const [win, original] of paneState.composeRouted) {
    try {
      win.ComposeMessage = original;
    } catch (error) {
      // the window is closing
    }
  }
  paneState.composeRouted.clear();
}

/* ------------------------------------------------------------------ */
/* Thunderbird Conversations state (read only)                         */
/* ------------------------------------------------------------------ */

// The add-on never installs or enables Thunderbird Conversations itself: the
// user installs it through their usual channel for add-ons. This only reports it.
async function paneConversationsState() {
  const addon = await AddonManager.getAddonByID(PANE_CONVERSATIONS_ID);
  return {
    installed: Boolean(addon),
    enabled: Boolean(addon && addon.isActive),
    version: addon ? addon.version : "",
    appVersion: Services.appinfo.version,
  };
}

/* ------------------------------------------------------------------ */
/* API                                                                 */
/* ------------------------------------------------------------------ */

function paneStatus() {
  const ext = conversationsExtension();
  return {
    conversations: Boolean(ext),
    conversationsVersion: ext ? ext.version : "",
    panes: paneState.windows.size,
    routedWindows: paneState.routed.size,
    drafts: [...paneState.drafts.values()].filter((d) => d.dirty).length,
    pendingSends: paneState.pending.size,
    newMessages: paneState.newOrder.length,
    dock: Boolean(paneState.dock && !paneState.dock.hidden),
    composeRoutedWindows: paneState.composeRouted.size,
    singleSentCopy: paneState.singleSentCopy,
    sentCopiesSkipped: paneState.sentCopiesSkipped,
    lastError: paneState.lastError,
  };
}

function paneRegisterChrome(extension) {
  if (paneState.chromeHandle) {
    return;
  }
  const aomStartup = Cc["@mozilla.org/addons/addon-manager-startup;1"].getService(
    Ci.amIAddonManagerStartup
  );
  const manifestURI = Services.io.newURI("manifest.json", null, extension.rootURI);
  paneState.chromeHandle = aomStartup.registerChrome(manifestURI, [
    ["content", PANE_PACKAGE, "reading-pane/"],
  ]);
}

this.outlookPane = class extends ExtensionCommon.ExtensionAPI {
  onShutdown(isAppShutdown) {
    if (isAppShutdown) {
      return;
    }
    // Disabled or updated: messages in their countdown are sent now.
    paneFlushPending();
    paneStopSentCopy();
    paneStopObserving();
    paneRemoveAll();
    if (paneState.chromeHandle) {
      paneState.chromeHandle.destruct();
      paneState.chromeHandle = null;
    }
    Services.obs.notifyObservers(null, "startupcache-invalidate");
  }

  getAPI(context) {
    const event = (name, set) =>
      new ExtensionCommon.EventManager({
        context,
        name: `outlookPane.${name}`,
        register: (fire) => {
          set.add(fire);
          return () => set.delete(fire);
        },
      }).api();

    return {
      outlookPane: {
        async configure(config) {
          paneState.extension = context.extension;
          Object.assign(paneState.config, paneClone(config));
          if (!paneState.config.enabled) {
            paneStopObserving();
            paneRemoveAll();
            return paneStatus();
          }
          paneRegisterChrome(context.extension);
          paneStartObserving();
          paneStartComposeRouting();
          for (const win of paneState.windows) {
            try {
              win.outlookModePane.configChanged();
            } catch (error) {
              paneNote("configure page", error);
            }
          }
          paneScanOpen();
          paneStyleMessageViews();
          return paneStatus();
        },
        async status() {
          return paneStatus();
        },
        async setSingleSentCopy(enabled) {
          paneSetSingleSentCopy(enabled);
          return { enabled: paneState.singleSentCopy, skipped: paneState.sentCopiesSkipped };
        },
        async conversationsState() {
          return paneConversationsState();
        },
        async sanitizeHtml(html) {
          return paneSanitize(`<html><body>${html}</body></html>`);
        },
        async setSignature(identityId, html) {
          const identity = paneIdentityByKey(identityId);
          if (!identity) {
            throw new ExtensionUtils.ExtensionError("Unknown identity");
          }
          // cellpadding as inline padding, as for a signature file.
          return paneSetSignature(identity, await paneEmbedImages(html, null));
        },
        async getSignature(identityId) {
          const identity = paneIdentityByKey(identityId);
          return identity ? paneSignature(identity, "new") : "";
        },
        async importSignature(identityId) {
          try {
            return await paneImportSignature(identityId);
          } catch (error) {
            paneNote("import signature", error);
            throw new ExtensionUtils.ExtensionError(String(error && error.message ? error.message : error));
          }
        },
        onPopOut: event("onPopOut", paneState.listeners.popOut),
        onSendResult: event("onSendResult", paneState.listeners.sendResult),
        onPin: event("onPin", paneState.listeners.pin),
        onOpenSettings: event("onOpenSettings", paneState.listeners.openSettings),
      },
    };
  }
};
