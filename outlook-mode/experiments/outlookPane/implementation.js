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
