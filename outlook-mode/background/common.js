"use strict";

// Shared settings, helpers and the message-handler registry.
// All background files share one global scope (MV2 background scripts).

/* ------------------------------------------------------------------ */
/* Settings                                                            */
/* ------------------------------------------------------------------ */

const DEFAULT_SETTINGS = {
  replyHeader: true, // Outlook "From/Sent/To/Cc/Subject" block on reply/forward
  removeQuoteBar: true, // no vertical quote bar on replies
  readingPaneNewestFirst: true, // Conversations view: newest message on top
  readingPaneExpandAll: true, // Conversations view: every message open, each with its attachments
  readingPaneFoldQuotes: true, // Conversations view: earlier mail quoted in a message folded
  readingPaneAttachmentsTop: true, // Conversations view: a message's attachments above its text
  readingPaneLatestOnly: true, // Conversations view: hide collapsed earlier messages
  readingPaneOutlookHeader: true, // Conversations view: Outlook To / Cc lines in message headers
  readingPaneLight: true, // Conversations view: white, also in dark theme
  inlineReply: true, // Conversations view: reply / forward inline in the reading pane
  inlineNewMessage: true, // new messages in a dock at the bottom of the main window
  pinnedOnTop: true, // keep pinned mail at the top of the Inbox
  pinSync: true, // Microsoft Graph accounts: sync pins with the server (Outlook's pin)
  pinButton: true, // cards view: a pin button in place of the star
  refreshButton: true, // message list header: get new mail and sync everything
  unreadHighlight: true, // cards view: unread mail with Outlook's blue bar, bold blue subject and time
  markConversationRead: true, // opening a mail marks its whole conversation read, as in Outlook
  oneCardPerConversation: true, // cards view: threads stay closed, one card per conversation
  pinnedSection: true, // cards view: pinned conversations in their own "Pinned" section at the top
  dateSections: true, // cards view: the rest grouped as Today, Yesterday, This week (last 7 days), Older
  newestInCard: true, // cards view: a conversation's card shows its newest mail's sender and date
  outlookSearch: true, // search box: suggestions, results in the mail tab (no new tab)
  collectContacts: true, // add the people you write to / reply to into Collected Addresses
  contactCompanyFormat: true, // address book: new contacts as "Name (Company-Department)", email first
  refreshKey: true, // F5 = the ⟳ refresh (all accounts, folders, pins, settings, contacts)
  undoSendSeconds: 10, // 0 = Undo Send off
  mentionTagging: true, // tag incoming mail that @mentions me
  settingsSync: false, // sync settings between devices through the Microsoft 365 mailbox
  syncCategories: { prefs: true, accounts: true, filters: true, contacts: true, addon: true },
};

const DEFAULT_STEPS = [
  {
    id: "done",
    name: "Done: mark read + archive",
    markRead: true,
    pin: false,
    compose: "none",
    forwardTo: "",
    after: "archive",
    moveTo: "",
  },
  {
    id: "pin-replyall",
    name: "Pin + reply all",
    markRead: true,
    pin: true,
    compose: "replyAll",
    forwardTo: "",
    after: "none",
    moveTo: "",
  },
];

// Undo Send holds a message at most this long (the longest choice in
// Preferences), whatever the stored value says.
const UNDO_SEND_MAX_SECONDS = 60;

function undoSendDelay(value) {
  const seconds = Math.floor(Number(value));
  return Number.isFinite(seconds) ? Math.min(Math.max(seconds, 0), UNDO_SEND_MAX_SECONDS) : 0;
}

async function getSettings() {
  const { settings } = await messenger.storage.local.get("settings");
  const merged = { ...DEFAULT_SETTINGS, ...(settings || {}) };
  merged.undoSendSeconds = undoSendDelay(merged.undoSendSeconds);
  return merged;
}

const STEP_COMPOSE = ["none", "reply", "replyAll", "forward"];
const STEP_AFTER = ["none", "move", "archive", "delete"];
const EMAIL_PATTERN = /^[^\s@<>(),;:"\\]+@[^\s@<>(),;:"\\]+\.[^\s@<>(),;:"\\]+$/;

// Quick Steps as stored (they can also arrive through settings sync), with
// every field checked: unknown actions become "none", and only well-formed
// addresses are kept for a forward. The destination folder is checked when
// the step runs (runStep).
// A folder id as the WebExtension API gives it ("account1://Inbox/Sub";
// Thunderbird names every account "account<n>"). It
// names a folder, not a file: Thunderbird resolves it to one of its own
// folders, and runStep() also checks it is a folder of the user's accounts.
// Ids with control characters or a ".." part are dropped all the same.
function folderIdText(value) {
  const id = typeof value === "string" ? value.slice(0, 2000) : "";
  return /^account[\w-]*:\/\//.test(id) && !/[\u0000-\u001f\u007f]/.test(id) && !/(^|\/)\.\.(\/|$)/.test(id) ? id : "";
}

function normalizeStep(step) {
  const source = step && typeof step === "object" ? step : {};
  // Plain text, one line (it is shown as a menu label and in Preferences).
  const text = (value, max) =>
    typeof value === "string" ? value.replace(/[\u0000-\u001f\u007f]+/g, " ").slice(0, max) : "";
  const forwardTo = parseAddressList(text(source.forwardTo, 2000))
    .filter((address) => EMAIL_PATTERN.test(emailOf(address)))
    .join(", ");
  return {
    ...source,
    id: text(source.id, 100),
    name: text(source.name, 200),
    markRead: Boolean(source.markRead),
    pin: Boolean(source.pin),
    compose: STEP_COMPOSE.includes(source.compose) ? source.compose : "none",
    forwardTo,
    after: STEP_AFTER.includes(source.after) ? source.after : "none",
    moveTo: folderIdText(source.moveTo),
  };
}

async function getSteps() {
  const { quickSteps } = await messenger.storage.local.get("quickSteps");
  return (Array.isArray(quickSteps) ? quickSteps : DEFAULT_STEPS).map(normalizeStep);
}

/* ------------------------------------------------------------------ */
/* Message handlers (filled in by each feature file, used by main.js)  */
/* ------------------------------------------------------------------ */

// type -> (message, sender) => value | Promise
const handlers = {};

// Wraps every handler except the six the compose script uses: only the
// add-on's own pages (moz-extension:) may call it. main.js checks every
// sender as well; this keeps the check next to each handler.
function fromOwnPage(handler) {
  return (message, sender) => {
    const ownPage = messenger.runtime.getURL("/");
    if (!sender || sender.id !== messenger.runtime.id || !String(sender.url || "").startsWith(ownPage)) {
      throw userError("Not allowed.");
    }
    return handler(message, sender);
  };
}

// Like fromOwnPage, for a handler that only some pages use: the caller must
// be one of those pages (paths in the package, e.g. "options/options.html").
// Another page of the add-on (one that shows mail content, for example)
// cannot call it.
function fromPages(paths, handler) {
  const allowed = paths.map((path) => messenger.runtime.getURL(path));
  return fromOwnPage((message, sender) => {
    const page = String(sender.url || "").replace(/[?#].*$/, "");
    if (!allowed.includes(page)) {
      throw userError("Not allowed.");
    }
    return handler(message, sender);
  });
}

// A Message-ID as Thunderbird stores it (without the angle brackets).
const MESSAGE_ID_PATTERN = /^[^\s<>\u0000-\u001f\u007f]{1,995}$/;
// A plain "left@right" Message-ID: letters, digits and . _ - + = / & $ # ! ^
// ` { | } ~ only (no quotes, brackets, parentheses, spaces, * ? %). Used
// where the id goes into a search (Graph $filter, thread lookup).
const PLAIN_MESSAGE_ID = /^[\w.!#$&+\/=^`{|}~-]{1,500}@[\w.-]{1,400}$/;

/* ------------------------------------------------------------------ */
/* Diagnostics                                                         */
/* ------------------------------------------------------------------ */

const recentErrors = [];

// An error caused by user input (shown to the user, not logged as a fault).
function userError(text) {
  const error = new Error(text);
  error.userFacing = true;
  return error;
}

function logError(where, error) {
  const text = String(error && error.message ? error.message : error);
  console.error(`[outlook-mode] ${where}:`, error);
  recentErrors.push({ at: new Date().toISOString(), where, text });
  if (recentErrors.length > 20) {
    recentErrors.shift();
  }
}

/* ------------------------------------------------------------------ */
/* Text and address helpers                                            */
/* ------------------------------------------------------------------ */

// Lowercase and strip diacritics so "tien" matches "Tiên".
function fold(text) {
  return (text || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .toLowerCase();
}

function emailOf(address) {
  if (typeof address !== "string") {
    return "";
  }
  const match = address.match(/<([^>]+)>/);
  return (match ? match[1] : address).trim().toLowerCase();
}

function nameOf(address) {
  if (typeof address !== "string") {
    return "";
  }
  const match = address.match(/^\s*"?([^"<]*?)"?\s*</);
  return match && match[1] ? match[1].trim() : emailOf(address);
}

function parseAddressList(text) {
  return (text || "")
    .split(/[;,\n]/)
    .map((s) => s.trim())
    .filter(Boolean);
}

// Outlook style: "Saturday, November 7, 2026 3:05 PM"
function formatOutlookDate(value) {
  const date = new Date(value);
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

/* ------------------------------------------------------------------ */
/* Messages, tabs, accounts                                            */
/* ------------------------------------------------------------------ */

async function collectList(list, cap = Infinity) {
  if (!list) {
    return [];
  }
  if (Array.isArray(list)) {
    return list;
  }
  const out = [...(list.messages || [])];
  let id = list.id;
  while (id && out.length < cap) {
    const next = await messenger.messages.continueList(id);
    out.push(...(next.messages || []));
    id = next.id;
  }
  if (id) {
    messenger.messages.abortList(id).catch(() => {});
  }
  return out;
}

async function activeTab() {
  const [tab] = await messenger.tabs.query({ active: true, currentWindow: true });
  return tab;
}

function isMailTab(tab) {
  return Boolean(tab && (tab.type === "mail" || tab.type === "messageDisplay"));
}

async function selectedMessages(tab) {
  if (!isMailTab(tab)) {
    return [];
  }
  try {
    if (tab.type === "mail") {
      return await collectList(await messenger.mailTabs.getSelectedMessages(tab.id));
    }
    return await collectList(await messenger.messageDisplay.getDisplayedMessages(tab.id));
  } catch (error) {
    logError("read selected messages", error);
    return [];
  }
}

async function myAddresses() {
  const accounts = await messenger.accounts.list(false);
  const set = new Set();
  for (const account of accounts) {
    for (const identity of account.identities || []) {
      if (identity.email) {
        set.add(identity.email.toLowerCase());
      }
    }
  }
  return set;
}

function isInbox(folder) {
  return Boolean(
    folder && (folder.type === "inbox" || (folder.specialUse || []).includes("inbox"))
  );
}

/* ------------------------------------------------------------------ */
/* Tags                                                                */
/* ------------------------------------------------------------------ */

const tagPromises = new Map();

// Create the tag once if it does not exist yet.
function ensureTag(key, name, color) {
  if (!tagPromises.has(key)) {
    const promise = (async () => {
      const list = await messenger.messages.tags.list();
      const existing = list.find((t) => t.key === key);
      if (!existing) {
        await messenger.messages.tags.create(key, name, color);
      } else if (!existing.tag) {
        // Repair a tag that ended up without a name.
        await messenger.messages.tags.update(key, { tag: name });
      }
    })().catch((error) => {
      tagPromises.delete(key);
      throw error;
    });
    tagPromises.set(key, promise);
  }
  return tagPromises.get(key);
}

async function addTag(message, key) {
  const tags = new Set(message.tags || []);
  if (!tags.has(key)) {
    tags.add(key);
    await messenger.messages.update(message.id, { tags: [...tags] });
  }
}

/* ------------------------------------------------------------------ */
/* Notifications                                                       */
/* ------------------------------------------------------------------ */

function notify(title, message) {
  return messenger.notifications
    .create({ type: "basic", title, message })
    .catch((error) => logError("notification", error));
}
