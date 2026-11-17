"use strict";

// Mailbox cleanup (cleanup/cleanup.html), like Outlook's Mailbox Cleanup and
// more: one scan of an account's mail, then the page filters it at once
// (older than…, large mail, notifications, unread for long, top senders,
// any folder / date / sender / subject / size) and acts on what it shows.
// Everything goes through Thunderbird's own message API, so Thunderbird
// syncs the result to the server as for any delete.

const CLEANUP_PAGE = "cleanup/cleanup.html";
const CLEANUP_CHUNK = 25; // also how often the page hears of the progress
// Folders whose mail is not cleaned with the rest (Trash and Junk have
// their own card).
const CLEANUP_KEEP_OUT = ["drafts", "outbox", "templates"];
// Exchange / Outlook system folders (and their subfolders) that are not
// worth cleaning: left out of the page and of the scan.
const CLEANUP_SYSTEM_FOLDERS = [
  "Sync Issues",
  "RSS Subscriptions",
  "RSS Feeds",
  "Conversation History",
  "Conversation Action Settings",
  "Quick Step Settings",
  "Social Activity Notifications",
  "Yammer Root",
];

function isSystemFolder(folder) {
  const top = (folder.path || "").split("/").filter(Boolean)[0] || "";
  return CLEANUP_SYSTEM_FOLDERS.some((name) => name.toLowerCase() === top.toLowerCase());
}

let cleanupTabId = null;

async function openCleanup() {
  if (cleanupTabId != null) {
    try {
      await messenger.tabs.update(cleanupTabId, { active: true });
      return;
    } catch (error) {
      cleanupTabId = null; // closed
    }
  }
  cleanupTabId = (await messenger.tabs.create({ url: messenger.runtime.getURL(CLEANUP_PAGE) })).id;
}

// A cleanup tab restored with the session opens before this add-on runs,
// and stays an empty page: load it again.
(async () => {
  const url = messenger.runtime.getURL(CLEANUP_PAGE);
  for (const tab of await messenger.tabs.query({})) {
    if (tab.url && tab.url.startsWith(url)) {
      cleanupTabId = tab.id;
      await messenger.tabs.update(tab.id, { url }).catch((error) => logError("cleanup tab reload", error));
    }
  }
})().catch((error) => logError("cleanup tab", error));

function cleanupProgress(text) {
  messenger.runtime.sendMessage({ type: "cleanup-progress", text }).catch(() => {});
}

async function listAll(list, onPage = () => {}) {
  const out = [...list.messages];
  let page = list;
  while (page.id) {
    onPage(out.length);
    page = await messenger.messages.continueList(page.id);
    out.push(...page.messages);
  }
  return out;
}

function folderKind(folder) {
  const special = folder.specialUse || [];
  return special[0] || folder.type || "";
}

async function cleanupAccounts() {
  const accounts = await messenger.accounts.list(false);
  return accounts
    .filter((account) => !["rss", "nntp", "im"].includes(account.type)) // mail accounts only
    .map((account) => ({ id: account.id, name: account.name, type: account.type }));
}

// Server quota (IMAP) or, for Microsoft Graph accounts, the folder sizes.
async function mailboxUsage(accountId, folders) {
  const inbox = folders.find((f) => (f.specialUse || []).includes("inbox"));
  if (inbox) {
    try {
      const info = await messenger.folders.getFolderInfo(inbox.id);
      const storage = (info.quota || []).find((q) => q.type === "STORAGE");
      if (storage && storage.limit) {
        return { used: storage.used * 1024, limit: storage.limit * 1024, source: "server" };
      }
    } catch (error) {
      // no quota information
    }
  }
  if ((await graphAccountIds()).includes(accountId)) {
    try {
      const top = await graphFetch(accountId, "/me/mailFolders?$top=100&$select=displayName,sizeInBytes", {}, { wait: false });
      const sizes = (top.value || []).map((f) => f.sizeInBytes).filter((n) => typeof n === "number");
      if (sizes.length) {
        return { used: sizes.reduce((a, b) => a + b, 0), limit: null, source: "server" };
      }
    } catch (error) {
      // sizeInBytes not offered, or offline
    }
  }
  return null;
}

// Every message of the account (but drafts, outbox and templates), as
// compact rows: [id, folder index, date ms, size, author, subject, flags].
const CLEANUP_READ = 1;
const CLEANUP_FLAGGED = 2;
const CLEANUP_PINNED = 4;
const CLEANUP_ATTACHMENT = 8;

// The ids of the mail the last scan listed: cleanupApply acts only on those
// (the page can only choose among what the scan showed).
let cleanupScanned = new Set();

async function cleanupScan({ accountId }) {
  if (!(await cleanupAccounts()).some((account) => account.id === accountId)) {
    throw userError("Unknown account.");
  }
  cleanupScanned = new Set();
  const folders = (await messenger.folders.query({ accountId, isVirtual: false })).filter(
    (folder) => !folder.isRoot && !CLEANUP_KEEP_OUT.includes(folderKind(folder)) && !isSystemFolder(folder)
  );
  const rows = [];
  for (const [index, folder] of folders.entries()) {
    const progress = (inFolder) =>
      cleanupProgress(`Reading ${folder.name} (${index + 1}/${folders.length}) — ${(rows.length + inFolder).toLocaleString()} mails so far…`);
    progress(0);
    let messages = [];
    try {
      messages = await listAll(await messenger.messages.list(folder.id), progress);
    } catch (error) {
      continue; // a folder that cannot be read (not synced)
    }
    let withAttachments = new Set();
    try {
      withAttachments = new Set(await messenger.outlookPin.withAttachments(messages.map((m) => m.id)));
    } catch (error) {
      logError("cleanup attachments", error);
    }
    for (const m of messages) {
      const flags =
        (m.read ? CLEANUP_READ : 0) |
        (m.flagged ? CLEANUP_FLAGGED : 0) |
        ((m.tags || []).includes(PIN_TAG_KEY) ? CLEANUP_PINNED : 0) |
        (withAttachments.has(m.id) ? CLEANUP_ATTACHMENT : 0);
      rows.push([m.id, index, new Date(m.date).getTime(), m.size || 0, m.author || "", m.subject || "", flags]);
    }
  }
  cleanupProgress("");
  cleanupScanned = new Set(rows.map((row) => row[0]));
  return {
    folders: folders.map((f) => ({ id: f.id, name: f.name, path: f.path, kind: folderKind(f) })),
    rows,
    usage: await mailboxUsage(accountId, folders),
    scannedAt: Date.now(),
  };
}

// The batches of cleanupApply. Thunderbird does not always report a batch
// done (Microsoft 365 slows down many changes in a row, and the report may
// then never come), so each batch is also watched mail by mail: it is over
// as soon as all its mail is done, when Stop is pressed, or when nothing
// has moved for a while. What is not done is tried again once at the end.
const CLEANUP_POLL = 1500;
const CLEANUP_STALL = 30e3;
let cleanupStopped = false;

function cleanupOperation(action, ids) {
  return action === "archive"
    ? messenger.messages.archive(ids)
    : action === "delete"
      ? messenger.messages.delete(ids, { deletePermanently: true })
      : messenger.messages.delete(ids, { isUserAction: true });
}

// Whether a message (before: as it was, now: as it is, null once gone)
// has had the action done.
function cleanupIsDone(action, before, now) {
  if (!now) {
    return true;
  }
  if (action === "delete") {
    return false;
  }
  if (action === "trash") {
    return (now.folder?.specialUse || []).includes("trash");
  }
  return now.folder?.id !== before.folder?.id;
}

// Runs one batch (entries: [id, message before]) and watches it. Returns
// { done: ids, left: entries, error, stalled }.
async function cleanupBatch(action, entries, onTick) {
  let finished = false;
  let error = null;
  cleanupOperation(
    action,
    entries.map(([id]) => id)
  ).then(
    () => (finished = true),
    (e) => {
      error = e;
      finished = true;
    }
  );
  const left = new Map(entries);
  const done = [];
  let lastMove = Date.now();
  for (;;) {
    const step = Math.min(100, CLEANUP_POLL);
    for (let waited = 0; waited < CLEANUP_POLL && !finished && !cleanupStopped; waited += step) {
      await new Promise((resolve) => setTimeout(resolve, step));
    }
    if (finished && !error) {
      done.push(...left.keys());
      return { done, left: [] };
    }
    for (const [id, before] of left) {
      const now = await messenger.messages.get(id).catch(() => null);
      if (cleanupIsDone(action, before, now)) {
        left.delete(id);
        done.push(id);
        lastMove = Date.now();
      }
    }
    const waiting = Date.now() - lastMove;
    onTick(done.length, waiting);
    if (!left.size || error || cleanupStopped) {
      return { done, left: [...left], error };
    }
    if (waiting > CLEANUP_STALL) {
      return { done, left: [...left], stalled: true };
    }
  }
}

// action: "trash" (to Deleted Items, as the Delete key does), "archive", or
// "delete" (permanently).
const CLEANUP_ACTIONS = ["archive", "trash", "delete"];

async function cleanupApply({ ids, action }) {
  if (!CLEANUP_ACTIONS.includes(action) || !Array.isArray(ids) || !ids.every(Number.isInteger)) {
    throw userError("Invalid cleanup request.");
  }
  if (!ids.every((id) => cleanupScanned.has(id))) {
    throw userError("Some of that mail was not in the last scan. Scan again.");
  }
  cleanupStopped = false;
  const doneIds = [];
  const failed = [];
  let skipped = 0;
  let late = [];
  // Batches in a row where nothing moved for CLEANUP_STALL: Thunderbird
  // gets no answer from the server any more, so going on would only wait.
  let stalls = 0;
  let stalled = false;
  let retrying = false;
  const report = (inBatch = 0, note = "") =>
    messenger.runtime
      .sendMessage({
        type: "cleanup-apply-progress",
        handled: retrying ? ids.length : Math.min(doneIds.length + inBatch + skipped + late.length, ids.length),
        total: ids.length,
        note,
      })
      .catch(() => {});
  const lateNote = () => (late.length ? `${mails(late.length)} not confirmed by the server yet — tried again at the end.` : "");

  const run = async (entries) => {
    const result = await cleanupBatch(action, entries, (count, waiting) =>
      report(count, waiting > 5e3 ? `Waiting for the server (${Math.round(waiting / 1000)} s)…` : lateNote())
    );
    doneIds.push(...result.done);
    if (result.error) {
      failed.push(String(result.error.message || result.error));
      logError(`cleanup ${action}`, result.error);
    }
    if (result.stalled) {
      stalls = result.done.length ? 0 : stalls + 1;
      if (stalls >= 2) {
        stalled = true;
        cleanupStopped = true;
      }
    } else {
      stalls = 0;
    }
    return result.left;
  };

  report();
  let next = 0;
  for (; next < ids.length && !cleanupStopped; next += CLEANUP_CHUNK) {
    // Pinned mail is never cleaned up, also when it was pinned after the
    // scan (on another device, through pin sync): read the tags again.
    const entries = [];
    for (const id of ids.slice(next, next + CLEANUP_CHUNK)) {
      const message = await messenger.messages.get(id).catch(() => null);
      if (!message) {
        doneIds.push(id); // already gone
      } else if ((message.tags || []).includes(PIN_TAG_KEY)) {
        skipped++;
      } else {
        entries.push([id, message]);
      }
    }
    if (entries.length) {
      late.push(...(await run(entries)));
    }
    report(0, lateNote());
  }
  const untouched = Math.max(0, ids.length - next);
  if (late.length && !cleanupStopped) {
    const retry = late;
    late = [];
    retrying = true;
    for (let i = 0; i < retry.length; i += CLEANUP_CHUNK) {
      if (cleanupStopped) {
        late.push(...retry.slice(i));
        break;
      }
      report(0, `Trying again the ${mails(retry.length - i)} the server did not confirm…`);
      late.push(...(await run(retry.slice(i, i + CLEANUP_CHUNK))));
    }
  }
  return {
    ok: !failed.length && !late.length && !untouched,
    doneIds,
    skipped,
    notDone: late.length + untouched,
    stopped: cleanupStopped && !stalled,
    stalled,
    failed,
  };
}

function mails(count) {
  return `${count} mail${count === 1 ? "" : "s"}`;
}

handlers["cleanup-open"] = fromOwnPage(() => openCleanup());
handlers["cleanup-accounts"] = fromOwnPage(() => cleanupAccounts());
handlers["cleanup-scan"] = fromOwnPage((message) => cleanupScan(message));
handlers["cleanup-apply"] = fromOwnPage((message) => cleanupApply(message));
handlers["cleanup-stop"] = fromOwnPage(() => {
  cleanupStopped = true;
});
