"use strict";

/* global ExtensionAPI, ExtensionCommon, ChromeUtils, Services, Ci, IOUtils */

// Experiment: reads and writes the Thunderbird settings that settings-sync.js
// keeps in step between devices.
//
// Keys:
//   pref:<name>                    a global preference
//   server:<sid>:<attribute>       mail.server.<server key>.<attribute>
//   identity:<email>:<attribute>   mail.identity.<identity key>.<attribute>
//   filters:<sid>                  the account's message filters (msgFilterRules.dat)
// <sid> is "type:user@host" and <email> the identity's address: they name the
// same account or identity on every device, while the keys in prefs.js
// (server2, id1) differ between devices. A value of null means "not set".
//
// Only settings that mean the same on every device are synced. Never synced:
// connection details, passwords and tokens, encryption settings, file paths,
// references to other accounts by key, and Thunderbird's own bookkeeping
// (versions, timestamps, migration flags).
//
// collectContacts() adds the people you have written to (recipients of your
// own mail) and replied to into Collected Addresses, which settings sync
// keeps the same on every device. A contact's id comes from its address, so
// devices that collect the same person create the same contact. Thunderbird
// also collects the recipients of mail you send, with random ids, so two
// devices can each add the same person: duplicates in Collected Addresses are
// merged into the contact with the smallest id, which every device picks.

var { MailServices } = ChromeUtils.importESModule("resource:///modules/MailServices.sys.mjs");
var { AddrBookCard } = ChromeUtils.importESModule("resource:///modules/AddrBookCard.sys.mjs");

const COLLECTED_BOOK_URI = "jsaddrbook://history.sqlite";
const COLLECT_MAX_ADDS = 5000;
const COLLECT_SKIPPED_FOLDERS =
  Ci.nsMsgFolderFlags.Trash | Ci.nsMsgFolderFlags.Junk | Ci.nsMsgFolderFlags.Virtual;
// Automated senders nobody writes to.
const COLLECT_SKIPPED_ADDRESS =
  /^(no-?reply|do-?not-?reply|donotreply|mailer-daemon|postmaster|bounces?|notifications?)([+._-]|@)/i;

const PREF_BRANCHES = [
  "mail.",
  "mailnews.",
  "msgcompose.",
  "font.",
  "calendar.",
  "intl.date_time.pattern_override.",
  "intl.regional_prefs.",
  "spellchecker.",
  "layout.spellcheckDefault",
  "browser.display.",
  "general.autoScroll",
  "general.smoothScroll",
  "ldap_2.autoComplete.",
];

const PREF_BLOCKED = [
  // Accounts are synced attribute by attribute (server:, identity:).
  /^mail\.(server|identity|account|accountmanager|smtpserver|smtp|outgoingserver|root)\./,
  /^mail\.(smtpservers|outgoingservers)$/,
  // Device or install specific.
  /^mail\.(winsearch|spotlight|shell|taskbar|startup|rights|cloud_files|openpgp|e2ee|provider|setup)\./,
  /^mailnews\.(database|start_page_override)\./,
  /^mailnews\.headers\.extraAddonHeaders$/,
  /^calendar\.(registry|alarms\.ignored|caldav|timezone\.local|list\.sortOrder|integration)\b/,
  /^font\.internaluseonly\./,
  // Changes with every move; would sync all the time.
  /^mail\.last_msg_movecopy_/,
  // Protections that stay as set on each device (a synced change could
  // turn off phishing or junk detection, or load remote content).
  /^mail\.(phishing|spam|smime)\./,
  /^mailnews\.message_display\.disable_remote_image$/,
  /^mailnews\.display\.(html_as|disallow_mime_handlers|prefer_plaintext)$/,
  // Bookkeeping and secrets.
  /(^|\.)(version|lastKey|uuid|guid|profileId)$/,
  /token|password|secret|oauth|timestamp|datestamp|lastUpdate|last_?check|migrat|first_?run|telemetry/i,
];

const SERVER_BLOCKED =
  /^(hostname|realhostname|port|type|userName|realuserName|directory|directory-rel|authMethod|socketType|clientid|clientidEnabled|ews_url|ews.*|graph.*|storeContractID|nextFilterTime|lastFilterTime|namespace\..*|capability|serverIDResponse|uidvalidity|deferred_to_account|oauth.*|password)$/i;

// Automatic Cc/Bcc and Reply-To stay per device: synced, they could send a
// copy of every mail, or the replies, to another address.
const IDENTITY_BLOCKED =
  /^(useremail|doCc|cc_list|doBcc|bcc_list|reply_to|smtpServer|valid|sig_file|sig_file-rel|openpgp.*|is_gnupg_key_id|last_entered_external_gnupg_key_id|signing_cert.*|encryption_cert.*|encryptionpolicy|sign_mail|e2etechpref|attachPgpKey|autoEncryptDrafts|protectSubject|sendAutocryptHeaders|autocrypt.*)$/i;

const GENERIC_BLOCKED = /token|password|secret|timestamp/i;

// Values that point at something on this device only.
function isLocalValue(value) {
  return (
    typeof value == "string" &&
    (/^(\/|[A-Za-z]:\\|file:)/.test(value) || /^(account|server|id|smtp)\d+$/.test(value))
  );
}

function isSyncedPref(name) {
  return (
    PREF_BRANCHES.some((branch) => name.startsWith(branch)) &&
    !PREF_BLOCKED.some((pattern) => pattern.test(name))
  );
}

function isSyncedAttribute(blocked, attribute) {
  return !blocked.test(attribute) && !GENERIC_BLOCKED.test(attribute);
}

function readPref(name) {
  if (!Services.prefs.prefHasUserValue(name)) {
    return null;
  }
  switch (Services.prefs.getPrefType(name)) {
    case Ci.nsIPrefBranch.PREF_BOOL:
      return Services.prefs.getBoolPref(name);
    case Ci.nsIPrefBranch.PREF_INT:
      return Services.prefs.getIntPref(name);
    case Ci.nsIPrefBranch.PREF_STRING:
      return Services.prefs.getStringPref(name);
    default:
      return null;
  }
}

function writePref(name, value) {
  if (value === null) {
    Services.prefs.clearUserPref(name);
    return;
  }
  switch (Services.prefs.getPrefType(name)) {
    case Ci.nsIPrefBranch.PREF_BOOL:
      Services.prefs.setBoolPref(name, Boolean(value));
      return;
    case Ci.nsIPrefBranch.PREF_INT:
      Services.prefs.setIntPref(name, Number(value));
      return;
    case Ci.nsIPrefBranch.PREF_STRING:
      Services.prefs.setStringPref(name, String(value));
      return;
  }
  if (typeof value == "boolean") {
    Services.prefs.setBoolPref(name, value);
  } else if (typeof value == "number") {
    Services.prefs.setIntPref(name, value);
  } else {
    Services.prefs.setStringPref(name, String(value));
  }
}

function serverId(server) {
  return `${server.type}:${server.username}@${server.hostname}`.toLowerCase();
}

function allServers() {
  return MailServices.accounts.accounts
    .map((account) => account.incomingServer)
    .filter((server) => server && server.type != "im");
}

function serverById(sid) {
  return allServers().find((server) => serverId(server) == sid) || null;
}

// The first identity with this address (several identities can share one).
function identityByEmail(email) {
  return (
    MailServices.accounts.allIdentities.find(
      (identity) => (identity.email || "").toLowerCase() == email
    ) || null
  );
}

// "server:graph:a@b:check_time" -> { id: "graph:a@b", attribute: "check_time" }
function splitKey(key, prefix) {
  const rest = key.slice(prefix.length);
  const at = rest.lastIndexOf(":");
  return at < 0 ? null : { id: rest.slice(0, at), attribute: rest.slice(at + 1) };
}

function filterFile(server) {
  const file = server.rootFolder.filePath.clone();
  file.append("msgFilterRules.dat");
  return file;
}

function keySupported(key) {
  if (key.startsWith("pref:")) {
    return isSyncedPref(key.slice("pref:".length));
  }
  if (key.startsWith("filters:")) {
    return Boolean(serverById(key.slice("filters:".length)));
  }
  if (key.startsWith("server:")) {
    const parts = splitKey(key, "server:");
    return Boolean(
      parts && serverById(parts.id) && isSyncedAttribute(SERVER_BLOCKED, parts.attribute)
    );
  }
  if (key.startsWith("identity:")) {
    const parts = splitKey(key, "identity:");
    return Boolean(
      parts && identityByEmail(parts.id) && isSyncedAttribute(IDENTITY_BLOCKED, parts.attribute)
    );
  }
  return false;
}

function readBranch(prefix, blocked, keyPrefix, values) {
  const branch = Services.prefs.getBranch(prefix);
  for (const attribute of branch.getChildList("")) {
    if (!isSyncedAttribute(blocked, attribute)) {
      continue;
    }
    const value = readPref(prefix + attribute);
    if (value !== null && !isLocalValue(value)) {
      values[keyPrefix + attribute] = value;
    }
  }
}

async function readSettings() {
  const values = {};
  for (const branch of PREF_BRANCHES) {
    for (const name of Services.prefs.getChildList(branch)) {
      if (!isSyncedPref(name)) {
        continue;
      }
      const value = readPref(name);
      if (value !== null && !isLocalValue(value)) {
        values[`pref:${name}`] = value;
      }
    }
  }

  const servers = [];
  const accounts = {};
  for (const account of MailServices.accounts.accounts) {
    const server = account.incomingServer;
    if (!server || server.type == "im") {
      continue;
    }
    const sid = serverId(server);
    servers.push(sid);
    accounts[account.key] = sid;
    readBranch(`mail.server.${server.key}.`, SERVER_BLOCKED, `server:${sid}:`, values);
    try {
      const file = filterFile(server);
      if (file.exists()) {
        values[`filters:${sid}`] = await IOUtils.readUTF8(file.path);
      }
    } catch (error) {
      console.error("[outlook-mode] read filters", sid, error);
    }
  }

  const identities = [];
  for (const identity of MailServices.accounts.allIdentities) {
    const email = (identity.email || "").toLowerCase();
    if (!email || identities.includes(email)) {
      continue;
    }
    identities.push(email);
    readBranch(`mail.identity.${identity.key}.`, IDENTITY_BLOCKED, `identity:${email}:`, values);
  }
  return { values, servers, identities, accounts };
}

// Filters from another device are read with Thunderbird's own parser
// first. Filters that forward or reply automatically, or run another
// add-on's custom action, are not taken: they could send mail out without
// the user doing anything. Such filters are set up on each device.
const FILTER_SYNC_MAX_BYTES = 1024 * 1024;

async function checkSyncedFilters(server, text) {
  if (text.length > FILTER_SYNC_MAX_BYTES) {
    throw new Error("filter file too big");
  }
  const temp = Services.dirsvc.get("TmpD", Ci.nsIFile);
  temp.append("outlook-mode-filters.dat");
  temp.createUnique(Ci.nsIFile.NORMAL_FILE_TYPE, 0o600);
  try {
    await IOUtils.writeUTF8(temp.path, text);
    const list = MailServices.filters.OpenFilterList(temp, server.rootFolder, null);
    const refused = [Ci.nsMsgFilterAction.Forward, Ci.nsMsgFilterAction.Reply, Ci.nsMsgFilterAction.Custom];
    for (let i = 0; i < list.filterCount; i++) {
      const filter = list.getFilterAt(i);
      if (filter.sortedActionList.some((action) => refused.includes(action.type))) {
        throw new Error(`filter "${filter.filterName}" forwards, replies or runs a custom action; not synced`);
      }
    }
  } finally {
    await IOUtils.remove(temp.path, { ignoreAbsent: true });
  }
}

async function applyOne(key, value) {
  if (!keySupported(key)) {
    throw new Error("not synced on this device");
  }
  if (isLocalValue(value)) {
    throw new Error("device-specific value");
  }
  if (key.startsWith("pref:")) {
    writePref(key.slice("pref:".length), value);
  } else if (key.startsWith("server:")) {
    const { id, attribute } = splitKey(key, "server:");
    writePref(`mail.server.${serverById(id).key}.${attribute}`, value);
  } else if (key.startsWith("identity:")) {
    const { id, attribute } = splitKey(key, "identity:");
    writePref(`mail.identity.${identityByEmail(id).key}.${attribute}`, value);
  } else if (key.startsWith("filters:")) {
    if (value === null) {
      return; // filters are never deleted by a sync
    }
    const server = serverById(key.slice("filters:".length));
    await checkSyncedFilters(server, String(value));
    const file = filterFile(server);
    await IOUtils.writeUTF8(file.path, String(value));
    server.setFilterList(MailServices.filters.OpenFilterList(file, server.rootFolder, null));
  }
}

/* ------------------------------------------------------------------ */
/* Collected contacts                                                  */
/* ------------------------------------------------------------------ */

// A UUID (version 5 layout) from SHA-1 of the address.
function contactUid(email) {
  const hash = Cc["@mozilla.org/security/hash;1"].createInstance(Ci.nsICryptoHash);
  hash.init(Ci.nsICryptoHash.SHA1);
  const bytes = [...unescape(encodeURIComponent(`outlook-mode:${email}`))].map((c) => c.charCodeAt(0));
  hash.update(bytes, bytes.length);
  const hex = [...hash.finish(false)].map((c) => c.charCodeAt(0).toString(16).padStart(2, "0")).join("");
  const variant = ((parseInt(hex[16], 16) & 3) | 8).toString(16);
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-5${hex.slice(13, 16)}-${variant}${hex.slice(17, 20)}-${hex.slice(20, 32)}`;
}

// Merges Collected Addresses contacts that share an address. Returns the
// number of contacts removed.
function mergeDuplicates(book) {
  const byEmail = new Map();
  for (const card of book.childCards) {
    const email = String(card.primaryEmail || "").toLowerCase();
    if (!card.isMailList && email) {
      byEmail.set(email, [...(byEmail.get(email) || []), card]);
    }
  }
  const extra = [];
  for (const cards of byEmail.values()) {
    if (cards.length < 2) {
      continue;
    }
    cards.sort((a, b) => (a.UID < b.UID ? -1 : 1));
    const [keep, ...rest] = cards;
    const named = rest.find((card) => card.displayName);
    if (!keep.displayName && named) {
      keep.displayName = named.displayName;
      book.modifyCard(keep);
    }
    extra.push(...rest);
  }
  if (extra.length) {
    book.deleteCards(extra);
  }
  return extra.length;
}

function idle() {
  return new Promise((resolve) => Services.tm.dispatchToMainThread(resolve));
}

// Adds the people in the mail since `since` (seconds; 0 = all mail) to
// Collected Addresses, unless an address book already has them.
async function collectContacts(since) {
  const book = MailServices.ab.getDirectory(COLLECTED_BOOK_URI);
  if (!book) {
    throw new Error("Collected Addresses is missing");
  }
  const mine = new Set(
    MailServices.accounts.allIdentities.map((identity) => (identity.email || "").toLowerCase())
  );
  const people = new Map(); // email -> Map(name -> count)
  const note = (raw) => {
    if (!raw) {
      return;
    }
    let parsed = [];
    try {
      parsed = MailServices.headerParser.parseEncodedHeaderW(raw);
    } catch (error) {
      return;
    }
    for (const { name, email } of parsed) {
      const address = String(email || "").trim().toLowerCase();
      if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(address) || mine.has(address) || COLLECT_SKIPPED_ADDRESS.test(address)) {
        continue;
      }
      const names = people.get(address) || new Map();
      const clean = String(name || "").replace(/^["'\s]+|["'\s]+$/g, "");
      if (clean && clean.toLowerCase() != address) {
        names.set(clean, (names.get(clean) || 0) + 1);
      }
      people.set(address, names);
    }
  };

  let scanned = 0;
  let folders = 0;
  for (const server of MailServices.accounts.allServers) {
    if (["rss", "nntp"].includes(server.type)) {
      continue; // feeds and newsgroups are not people you write to
    }
    for (const folder of server.rootFolder.descendants) {
      if (folder.flags & COLLECT_SKIPPED_FOLDERS) {
        continue;
      }
      const sentFolder = folder.isSpecialFolder(Ci.nsMsgFolderFlags.SentMail, true);
      let db;
      try {
        db = folder.msgDatabase;
      } catch (error) {
        continue; // no local summary for this folder
      }
      folders++;
      for (const hdr of db.enumerateMessages()) {
        if (hdr.dateInSeconds < since) {
          continue;
        }
        let fromMe = sentFolder;
        if (!fromMe) {
          const author = MailServices.headerParser.parseEncodedHeaderW(hdr.author)[0];
          fromMe = Boolean(author && mine.has(String(author.email || "").toLowerCase()));
        }
        if (fromMe) {
          note(hdr.recipients);
          note(hdr.ccList);
          note(hdr.bccList);
        } else if (hdr.flags & Ci.nsMsgMessageFlags.Replied) {
          note(hdr.author);
        }
        if (++scanned % 2000 == 0) {
          await idle();
        }
      }
      try {
        folder.closeDBIfFolderNotOpen(false);
      } catch (error) {
        // still in use
      }
    }
  }

  const merged = mergeDuplicates(book);
  let added = 0;
  for (const [email, names] of people) {
    if (added >= COLLECT_MAX_ADDS) {
      break;
    }
    if (MailServices.ab.cardForEmailAddress(email)) {
      continue;
    }
    const name = [...names].sort((a, b) => b[1] - a[1])[0]?.[0] || "";
    const card = new AddrBookCard();
    card.UID = contactUid(email);
    card.primaryEmail = email;
    if (name) {
      card.displayName = name;
    }
    try {
      book.addCard(card);
      added++;
    } catch (error) {
      console.error("[outlook-mode] collect contact", email, error);
    }
    if (added % 200 == 0) {
      await idle();
    }
  }
  return { scanned, folders, people: people.size, added, merged, book: book.dirName };
}

/* ------------------------------------------------------------------ */
/* Recommended Thunderbird settings                                    */
/* ------------------------------------------------------------------ */

// What the add-on's features expect: HTML compose (@mention links), the
// reply above the quote (as Outlook does), and the Cards View (the pin
// button). Returns what was changed.
function applyRecommended() {
  const changed = [];
  for (const identity of MailServices.accounts.allIdentities) {
    if (!identity.email) {
      continue;
    }
    if (!identity.composeHtml) {
      identity.composeHtml = true;
      changed.push({ what: "html", identity: identity.email });
    }
    if (identity.replyOnTop != 1) {
      identity.replyOnTop = 1;
      changed.push({ what: "replyOnTop", identity: identity.email });
    }
  }
  if (Services.prefs.getIntPref("mail.threadpane.listview", 0) != 0) {
    Services.prefs.setIntPref("mail.threadpane.listview", 0);
    changed.push({ what: "cardsView", identity: "" });
  }
  if (changed.length) {
    Services.prefs.savePrefFile(null);
  }
  return { changed };
}

this.outlookSync = class extends ExtensionCommon.ExtensionAPI {
  getAPI(context) {
    return {
      outlookSync: {
        async readSettings() {
          return readSettings();
        },
        async supportedKeys(keys) {
          return keys.filter(keySupported);
        },
        async applySettings(changes) {
          const failed = {};
          for (const [key, value] of Object.entries(changes)) {
            try {
              await applyOne(key, value);
            } catch (error) {
              failed[key] = String(error && error.message ? error.message : error);
            }
          }
          try {
            Services.prefs.savePrefFile(null);
          } catch (error) {
            console.error("[outlook-mode] save prefs", error);
          }
          return { failed };
        },
        async applyRecommended() {
          return applyRecommended();
        },
        async collectContacts(options) {
          return collectContacts(Math.max(0, Number(options?.since) || 0));
        },
      },
    };
  }
};
