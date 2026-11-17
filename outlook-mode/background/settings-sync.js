"use strict";

// Settings sync between devices. Thunderbird keeps its settings, filters and
// local address books on each device only and has no sync of its own. This
// file keeps them in step through the Microsoft 365 mailbox: the state is a
// gzip JSON attachment on a message in a hidden mail folder
// ("Outlook Mode Sync"), read and written with Microsoft Graph.
//
// What is synced (each part can be turned off in Settings):
// - prefs: global preferences (layout, reading, composing, fonts, tags,
//   calendar view...); the rules are in experiments/outlookSync;
// - accounts: account and identity settings (signature, reply position,
//   Sent/Drafts folders, check interval...), matched by user@host and email;
// - filters: the message filters of each account;
// - contacts: local address books (Personal, Collected and your own books);
// - addon: this add-on's settings and Quick Steps.
// Not synced: passwords and sign-ins, encryption keys, file paths, accounts
// themselves (add each account once per device), scheduled mail, pins (they
// have their own sync, pin-sync.js).
//
// Every value is an entry { v: value, t: time changed, d: device }, with
// v = null for "not set / deleted". On each sync the local values are compared
// with the state of the last sync ("base") to find local changes; for each key
// the newer of local and server wins. A device that joins for the first time
// takes the server's values. The values a sync replaces are saved first, so
// Preferences → Undo last sync can put them back.

const SYNC_FOLDER_NAME = "Outlook Mode Sync";
const SYNC_SUBJECT = "Outlook Mode settings sync (do not delete)";
const SYNC_FILE_NAME = "outlook-mode-sync.json.gz";
const SYNC_FORMAT = 1;
const SETTINGS_SYNC_ALARM = "settings-sync";
const SETTINGS_SYNC_MINUTES = 10;
const SETTINGS_SYNC_DEBOUNCE_MS = 30 * 1000;
const SYNC_BASE_KEY = "settingsSyncBase";
const SYNC_DEVICE_KEY = "settingsSyncDevice";
const SYNC_BACKUP_KEY = "settingsSyncBackup";
const MAX_SYNC_BASE64 = 3 * 1024 * 1024; // Graph limit for an attachment sent inline
const TOMBSTONE_MS = 180 * 864e5;
// A synced entry dated more than this in the future is ignored (clock
// differences between devices are far smaller), so a made-up time cannot
// make an entry win forever.
const SYNC_MAX_FUTURE_MS = 864e5;

function plausibleEntry(entry, now) {
  return entry && Number.isFinite(entry.t) && entry.t >= 0 && entry.t <= now + SYNC_MAX_FUTURE_MS ? entry : undefined;
}
// Add-on settings that stay per device.
const LOCAL_ONLY_SETTINGS = ["settingsSync", "syncCategories", "outlookSearch"];

const settingsSyncStatus = {
  lastSync: "",
  lastError: "",
  account: "",
  keys: 0,
  lastApplied: 0,
  lastFailed: {},
  bytes: 0,
};

function categoryOf(key) {
  if (key.startsWith("pref:")) {
    return "prefs";
  }
  if (key.startsWith("server:") || key.startsWith("identity:")) {
    return "accounts";
  }
  if (key.startsWith("filters:")) {
    return "filters";
  }
  if (key.startsWith("contact:")) {
    return "contacts";
  }
  if (key.startsWith("addon:")) {
    return "addon";
  }
  return "";
}

function sameValue(a, b) {
  return JSON.stringify(a === undefined ? null : a) === JSON.stringify(b === undefined ? null : b);
}

async function deviceId() {
  const stored = await messenger.storage.local.get(SYNC_DEVICE_KEY);
  if (stored[SYNC_DEVICE_KEY]) {
    return stored[SYNC_DEVICE_KEY];
  }
  const id = `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
  await messenger.storage.local.set({ [SYNC_DEVICE_KEY]: id });
  return id;
}

/* ------------------------------------------------------------------ */
/* Local values                                                        */
/* ------------------------------------------------------------------ */

// A stable form of a vCard: unfolded, without the photo (too big) and the
// fields Thunderbird rewrites on every save, other lines sorted.
function cleanVCard(vCard) {
  const lines = String(vCard || "")
    .replace(/\r?\n[ \t]/g, "")
    .split(/\r?\n/)
    .filter((line) => line && !/^(BEGIN|END|VERSION|PHOTO|PRODID|REV|UID)[;:]/i.test(line));
  const version = (String(vCard).match(/^VERSION:.*$/im) || ["VERSION:4.0"])[0];
  return ["BEGIN:VCARD", version, ...lines.sort(), "END:VCARD"].join("\r\n");
}

// Folder ids ("account1://Inbox") contain the account key, which differs
// between devices: Quick Steps store the account id (type:user@host) instead.
function exportSteps(steps, accounts) {
  return steps.map((step) => {
    const { moveToSynced, ...rest } = step;
    const match = /^([^:]+):(.*)$/.exec(step.moveTo || "");
    let moveTo = step.moveTo || "";
    if (match && accounts[match[1]]) {
      moveTo = { account: accounts[match[1]], path: match[2] };
    } else if (!moveTo && moveToSynced) {
      moveTo = moveToSynced; // the folder's account is not on this device
    }
    return { ...rest, moveTo };
  });
}

function importSteps(steps, accounts) {
  const keyOf = Object.fromEntries(Object.entries(accounts).map(([key, sid]) => [sid, key]));
  return steps.map((step) => {
    if (!step.moveTo || typeof step.moveTo === "string") {
      return step;
    }
    const key = keyOf[step.moveTo.account];
    // A folder path of that account ("//Inbox/Done"): no ".." part, no
    // control characters (folderIdText); anything else leaves the step
    // without a folder, to be picked again in Preferences.
    const path = typeof step.moveTo.path === "string" ? step.moveTo.path : "";
    const moveTo = key && path.startsWith("//") ? folderIdText(`${key}:${path}`) : "";
    return moveTo ? { ...step, moveTo } : { ...step, moveTo: "", moveToSynced: step.moveTo };
  });
}

async function readLocal(categories, remoteKeys) {
  const native = await messenger.outlookSync.readSettings();
  const values = {};
  const books = new Map(); // address book name -> id
  const contactBooks = new Map(); // contact id -> address book id (local books only)
  for (const [key, value] of Object.entries(native.values)) {
    if (categories[categoryOf(key)]) {
      values[key] = value;
    }
  }
  if (categories.contacts) {
    for (const book of await messenger.addressBooks.list(true)) {
      if (book.remote || book.readOnly) {
        continue; // CardDAV / LDAP books sync on their own
      }
      books.set(book.name, book.id);
      for (const contact of book.contacts || []) {
        const vCard = contact.vCard || (contact.properties && contact.properties.vCard);
        contactBooks.set(contact.id, book.id);
        if (vCard) {
          values[`contact:${book.name}:${contact.id}`] = cleanVCard(vCard);
        }
      }
    }
  }
  if (categories.addon) {
    const settings = await getSettings();
    for (const [name, value] of Object.entries(settings)) {
      if (!LOCAL_ONLY_SETTINGS.includes(name)) {
        values[`addon:setting:${name}`] = value;
      }
    }
    values["addon:quickSteps"] = exportSteps(await getSteps(), native.accounts);
  }
  // Keys of Thunderbird settings this device can take; the rest (accounts
  // that are not here, settings of a newer add-on version) are left alone.
  const nativeKeys = remoteKeys.filter((key) => ["prefs", "accounts", "filters"].includes(categoryOf(key)));
  const supported = new Set([
    ...Object.keys(native.values),
    ...(await messenger.outlookSync.supportedKeys(nativeKeys)),
  ]);
  const applicable = (key) => {
    const category = categoryOf(key);
    if (!categories[category]) {
      return false;
    }
    return ["contacts", "addon"].includes(category) || supported.has(key);
  };
  return { values, books, contactBooks, accounts: native.accounts, applicable };
}

/* ------------------------------------------------------------------ */
/* Applying values                                                     */
/* ------------------------------------------------------------------ */

// storage.onChanged fires a little after a write: changes this soon after
// a sync wrote the add-on's settings come from that sync.
let addonSettingsAppliedAt = 0;
const OWN_WRITE_WINDOW_MS = 2000;

// Only settings this add-on has, with a value of the same kind as the
// default (true/false, a number, text), are taken from another device;
// per-device settings never are. getSettings() also bounds Undo Send.
function syncedSettingProblem(name, value) {
  if (!Object.hasOwn(DEFAULT_SETTINGS, name) || LOCAL_ONLY_SETTINGS.includes(name)) {
    return "unknown add-on setting";
  }
  const expected = typeof DEFAULT_SETTINGS[name];
  if (value !== null && (typeof value !== expected || expected === "object")) {
    return "invalid value";
  }
  if (typeof value === "number" && !Number.isFinite(value)) {
    return "invalid value";
  }
  return "";
}

async function applyAddon(changes, local) {
  const failed = {};
  if (!Object.keys(changes).length) {
    return failed;
  }
  const stored = await messenger.storage.local.get(["settings"]);
  const settings = { ...(stored.settings || {}) };
  const update = {};
  for (const [key, value] of Object.entries(changes)) {
    if (key.startsWith("addon:setting:")) {
      const name = key.slice("addon:setting:".length);
      const problem = syncedSettingProblem(name, value);
      if (problem) {
        failed[key] = problem;
        continue;
      }
      if (value === null) {
        delete settings[name];
      } else {
        settings[name] = value;
      }
      update.settings = settings;
    } else if (key === "addon:quickSteps" && Array.isArray(value)) {
      update.quickSteps = importSteps(value, local.accounts);
    } else {
      failed[key] = "unknown add-on setting";
    }
  }
  addonSettingsAppliedAt = Date.now();
  await messenger.storage.local.set(update);
  return failed;
}

// A contact from another device: one vCard of reasonable size, made only of
// property lines, with no markup (< >) or script URL in any value. It is
// imported in the clean form (no photo or UID line).
const MAX_SYNC_VCARD = 64 * 1024;
const VCARD_LINE = /^[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)?(;[^:\r\n<>]*)?:[^<>]*$/;

function importableVCard(value) {
  if (typeof value !== "string" || value.length > MAX_SYNC_VCARD) {
    return null;
  }
  const lines = value.replace(/\r?\n[ \t]/g, "").split(/\r?\n/).filter(Boolean);
  if (
    lines[0] !== "BEGIN:VCARD" ||
    lines[lines.length - 1] !== "END:VCARD" ||
    lines.filter((line) => /^(BEGIN|END):/i.test(line)).length !== 2 ||
    !lines.every((line) => VCARD_LINE.test(line)) ||
    /(javascript|vbscript):/i.test(value)
  ) {
    return null;
  }
  return cleanVCard(value);
}

// At most this many address books are created by one sync (a new device
// gets the user's few local books, not an unbounded number).
const MAX_NEW_BOOKS_PER_SYNC = 5;

async function applyContacts(changes, local) {
  const failed = {};
  let newBooks = 0;
  for (const [key, value] of Object.entries(changes)) {
    const rest = key.slice("contact:".length);
    const at = rest.lastIndexOf(":");
    const bookName = rest.slice(0, at);
    const id = rest.slice(at + 1);
    if (!/^[\w-]{1,100}$/.test(id) || !bookName || bookName.length > 200 || /[\u0000-\u001f\u007f]/.test(bookName)) {
      failed[key] = "invalid contact";
      continue;
    }
    // Only contacts of this device's own local (writable) address books are
    // changed or deleted, and only in the book the change names; a new
    // contact gets an id that no local contact has.
    const bookId = local.books.get(bookName);
    const ownContact = Boolean(bookId) && local.contactBooks.get(id) === bookId;
    if ((key in local.values && !ownContact) || (!(key in local.values) && local.contactBooks.has(id))) {
      failed[key] = "contact not in that address book";
      continue;
    }
    const vCard = value === null ? null : importableVCard(value);
    if (value !== null && !vCard) {
      failed[key] = "invalid vCard";
      continue;
    }
    try {
      if (value === null) {
        if (key in local.values) {
          await messenger.contacts.delete(id);
        }
      } else if (key in local.values) {
        await messenger.contacts.update(id, { vCard });
      } else {
        let targetBook = bookId;
        if (!targetBook) {
          if (newBooks >= MAX_NEW_BOOKS_PER_SYNC) {
            failed[key] = "too many new address books";
            continue;
          }
          targetBook = await messenger.addressBooks.create({ name: bookName });
          newBooks++;
          local.books.set(bookName, targetBook);
        }
        await messenger.contacts.create(targetBook, id, { vCard });
        local.contactBooks.set(id, targetBook);
      }
    } catch (error) {
      failed[key] = String(error.message || error);
    }
  }
  return failed;
}

// A Thunderbird setting from the sync: a key this device syncs (the
// experiment's own list: allowed preference branches, no account,
// security or secret settings) with a plain value of reasonable size.
const MAX_SYNC_NATIVE_VALUE = 512 * 1024;

function nativeValueProblem(value) {
  if (value === null || typeof value === "boolean" || Number.isSafeInteger(value)) {
    return "";
  }
  if (typeof value === "string" && value.length <= MAX_SYNC_NATIVE_VALUE) {
    return "";
  }
  return "invalid value";
}

// Returns { key: reason } for the values that could not be applied.
async function applyLocal(changes, local) {
  const native = {};
  const contacts = {};
  const addon = {};
  for (const [key, value] of Object.entries(changes)) {
    const category = categoryOf(key);
    if (category === "contacts") {
      contacts[key] = value;
    } else if (category === "addon") {
      addon[key] = value;
    } else {
      native[key] = value;
    }
  }
  const failed = {};
  const supported = new Set(await messenger.outlookSync.supportedKeys(Object.keys(native)));
  for (const [key, value] of Object.entries(native)) {
    const problem = !categoryOf(key) || !supported.has(key) ? "not synced on this device" : nativeValueProblem(value);
    if (problem) {
      failed[key] = problem;
      delete native[key];
    }
  }
  if (Object.keys(native).length) {
    Object.assign(failed, (await messenger.outlookSync.applySettings(native)).failed);
  }
  Object.assign(failed, await applyContacts(contacts, local));
  Object.assign(failed, await applyAddon(addon, local));
  return failed;
}

/* ------------------------------------------------------------------ */
/* Server storage                                                      */
/* ------------------------------------------------------------------ */

async function gzipBase64(text) {
  const stream = new Blob([text]).stream().pipeThrough(new CompressionStream("gzip"));
  const bytes = new Uint8Array(await new Response(stream).arrayBuffer());
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return btoa(binary);
}

// The sync file unpacked is at most this big; a real one is far smaller.
const MAX_SYNC_JSON_BYTES = 32 * 1024 * 1024;

// Unpacks a sync file, stopping as soon as it grows past the limit (so a
// small file that unpacks to gigabytes cannot exhaust memory).
async function gunzipBase64(base64) {
  if (typeof base64 !== "string" || base64.length > MAX_SYNC_BASE64 * 2) {
    throw new Error("The sync file is too big.");
  }
  const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
  const reader = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("gzip")).getReader();
  const chunks = [];
  let total = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) {
      break;
    }
    total += value.length;
    if (total > MAX_SYNC_JSON_BYTES) {
      await reader.cancel();
      throw new Error("The sync file is too big once unpacked.");
    }
    chunks.push(value);
  }
  return new TextDecoder().decode(await new Blob(chunks).arrayBuffer());
}

const syncFolderIds = new Map(); // accountId -> Graph folder id

async function syncFolderId(accountId) {
  if (syncFolderIds.has(accountId)) {
    return syncFolderIds.get(accountId);
  }
  const params = new URLSearchParams({
    includeHiddenFolders: "true",
    $filter: `displayName eq ${odataString(SYNC_FOLDER_NAME)}`,
    $select: "id",
  });
  const found = await graphFetch(accountId, `/me/mailFolders?${params}`);
  let id = found.value && found.value.length ? found.value[0].id : null;
  if (!id) {
    const created = await graphFetch(accountId, "/me/mailFolders", {
      method: "POST",
      body: JSON.stringify({ displayName: SYNC_FOLDER_NAME, isHidden: true }),
    });
    id = created.id;
  }
  syncFolderIds.set(accountId, id);
  return id;
}

function newer(a, b) {
  if (!a) {
    return b;
  }
  if (!b) {
    return a;
  }
  return b.t > a.t ? b : a;
}

// All states in the folder, merged. There is normally one message; there are
// more when two devices wrote at the same time.
async function downloadState(accountId) {
  const folderId = encodeURIComponent(await syncFolderId(accountId));
  const list = await graphFetch(accountId, `/me/mailFolders/${folderId}/messages?$select=id&$top=50`);
  const entries = {};
  const messageIds = [];
  for (const item of list.value || []) {
    messageIds.push(item.id);
    const attachments = await graphFetch(accountId, `/me/messages/${encodeURIComponent(item.id)}/attachments`);
    for (const attachment of attachments.value || []) {
      if (attachment.name !== SYNC_FILE_NAME || !attachment.contentBytes) {
        continue;
      }
      const state = JSON.parse(await gunzipBase64(attachment.contentBytes));
      if (state.format > SYNC_FORMAT) {
        throw userError("Another device uses a newer Outlook Mode. Update the add-on on this device.");
      }
      for (const [key, entry] of Object.entries(state.entries || {})) {
        entries[key] = newer(entries[key], entry);
      }
    }
  }
  return { entries, messageIds };
}

async function uploadState(accountId, entries, oldMessageIds, device) {
  const folderId = encodeURIComponent(await syncFolderId(accountId));
  const json = JSON.stringify({ format: SYNC_FORMAT, updated: new Date().toISOString(), device, entries });
  const contentBytes = await gzipBase64(json);
  if (contentBytes.length > MAX_SYNC_BASE64) {
    throw userError(
      `The settings are too large to sync (${Math.round(contentBytes.length / 1024)} KB). Turn off contacts sync.`
    );
  }
  settingsSyncStatus.bytes = contentBytes.length;
  await graphFetch(accountId, `/me/mailFolders/${folderId}/messages`, {
    method: "POST",
    body: JSON.stringify({
      subject: SYNC_SUBJECT,
      body: {
        contentType: "text",
        content: "Settings of the Outlook Mode add-on for Thunderbird, shared between your devices.",
      },
      attachments: [
        {
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: SYNC_FILE_NAME,
          contentType: "application/gzip",
          contentBytes,
        },
      ],
    }),
  });
  for (const id of oldMessageIds) {
    await graphFetch(accountId, `/me/messages/${encodeURIComponent(id)}/permanentDelete`, {
      method: "POST",
    }).catch((error) => {
      if (!String(error.message).startsWith("Graph 404")) {
        logError("settings sync: remove old state", error); // merged again next time
      }
    });
  }
}

/* ------------------------------------------------------------------ */
/* Sync                                                                */
/* ------------------------------------------------------------------ */

let settingsSyncChain = Promise.resolve();
function withSettingsSyncLock(task) {
  const run = settingsSyncChain.then(task, task);
  settingsSyncChain = run.catch(() => {});
  return run;
}

async function runSettingsSync() {
  const settings = await getSettings();
  const [accountId] = await graphAccountIds();
  if (!accountId) {
    throw userError("Settings sync needs a Microsoft 365 (Graph) account.");
  }
  settingsSyncStatus.account = accountId;
  const device = await deviceId();
  const remote = await downloadState(accountId);
  const local = await readLocal(settings.syncCategories, Object.keys(remote.entries));
  const base = (await messenger.storage.local.get(SYNC_BASE_KEY))[SYNC_BASE_KEY] || {};
  const now = Date.now();
  const merged = {};
  const toApply = {};
  const keys = new Set([...Object.keys(local.values), ...Object.keys(remote.entries), ...Object.keys(base)]);
  for (const key of keys) {
    const theirs = plausibleEntry(remote.entries[key], now);
    if (!local.applicable(key)) {
      const keep = newer(theirs, base[key]);
      if (keep) {
        merged[key] = keep;
      }
      continue;
    }
    const current = key in local.values ? local.values[key] : null;
    const last = base[key];
    let mine = null;
    if (last && sameValue(current, last.v)) {
      mine = last;
    } else if (last || current !== null) {
      // Changed here since the last sync. A value from before this device
      // joined (no base yet) loses against the server's.
      mine = { v: current, t: last || !theirs ? now : 0, d: device };
    }
    const winner = !mine ? theirs : !theirs ? mine : mine.t > theirs.t ? mine : theirs;
    if (!winner || (winner.v === null && now - winner.t > TOMBSTONE_MS)) {
      continue;
    }
    merged[key] = winner;
    if (!sameValue(winner.v, current)) {
      toApply[key] = winner.v;
    }
  }

  let failed = {};
  if (Object.keys(toApply).length) {
    const backup = {};
    for (const key of Object.keys(toApply)) {
      backup[key] = key in local.values ? local.values[key] : null;
    }
    await messenger.storage.local.set({
      [SYNC_BACKUP_KEY]: { at: new Date().toISOString(), values: backup },
    });
    failed = await applyLocal(toApply, local);
  }
  const applied = Object.keys(toApply).length - Object.keys(failed).length;

  const changedOnServer =
    remote.messageIds.length !== 1 ||
    Object.keys(merged).length !== Object.keys(remote.entries).length ||
    Object.keys(merged).some((key) => !sameValue(merged[key], remote.entries[key]));
  if (changedOnServer) {
    await uploadState(accountId, merged, remote.messageIds, device);
  }

  // What failed to apply is retried next time: keep its old base.
  const newBase = { ...merged };
  for (const key of Object.keys(failed)) {
    if (base[key]) {
      newBase[key] = base[key];
    } else {
      delete newBase[key];
    }
  }
  await messenger.storage.local.set({ [SYNC_BASE_KEY]: newBase });

  settingsSyncStatus.lastSync = new Date().toISOString();
  settingsSyncStatus.lastError = "";
  settingsSyncStatus.keys = Object.keys(merged).length;
  settingsSyncStatus.lastApplied = applied;
  settingsSyncStatus.lastFailed = failed;
  if (applied) {
    notify(
      "Outlook Mode",
      `Applied ${applied} setting(s) from your other devices. Some take effect after restarting Thunderbird.`
    );
  }
  return { ok: true, applied, failed, uploaded: changedOnServer, keys: settingsSyncStatus.keys };
}

async function syncSettings() {
  if (!(await getSettings()).settingsSync) {
    return { ok: false, error: "Settings sync is off. Turn it on and click Save." };
  }
  // While Microsoft's server is failing, no new sync starts; the server
  // coming back runs one.
  if (await anyServerFailing()) {
    return { ok: false, waiting: true, error: "Waiting for Microsoft's server to come back." };
  }
  return withSettingsSyncLock(async () => {
    try {
      return await runSettingsSync();
    } catch (error) {
      settingsSyncStatus.lastError = String(error && error.message ? error.message : error);
      throw error;
    }
  });
}

// Put back the values the last sync replaced. They count as changes made on
// this device, so the next sync sends them to the other devices.
async function undoLastSettingsSync() {
  const backup = (await messenger.storage.local.get(SYNC_BACKUP_KEY))[SYNC_BACKUP_KEY];
  if (!backup) {
    throw userError("There is nothing to undo.");
  }
  const failed = await withSettingsSyncLock(async () => {
    const settings = await getSettings();
    const local = await readLocal(settings.syncCategories, Object.keys(backup.values));
    const result = await applyLocal(backup.values, local);
    await messenger.storage.local.remove(SYNC_BACKUP_KEY);
    return result;
  });
  await syncSettings();
  return { ok: true, restored: Object.keys(backup.values).length - Object.keys(failed).length, failed };
}

/* ------------------------------------------------------------------ */
/* Triggers                                                            */
/* ------------------------------------------------------------------ */

function logSyncError(where) {
  return (error) => {
    if (!error || !error.userFacing) {
      logError(where, error);
    }
  };
}

serverBackHooks.push(() => syncSettings().catch(logSyncError("settings sync after outage")));

messenger.alarms.create(SETTINGS_SYNC_ALARM, { periodInMinutes: SETTINGS_SYNC_MINUTES });
messenger.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === SETTINGS_SYNC_ALARM) {
    syncSettings().catch(logSyncError("settings sync alarm"));
  }
});

// Soon after this add-on's own settings change (Preferences page, Quick Steps).
let settingsSyncTimer = null;
messenger.storage.onChanged.addListener((changes) => {
  if (
    Date.now() - addonSettingsAppliedAt < OWN_WRITE_WINDOW_MS ||
    !(changes.settings || changes.quickSteps)
  ) {
    return;
  }
  clearTimeout(settingsSyncTimer);
  settingsSyncTimer = setTimeout(
    () => syncSettings().catch(logSyncError("settings sync after change")),
    SETTINGS_SYNC_DEBOUNCE_MS
  );
});

setTimeout(() => syncSettings().catch(logSyncError("startup settings sync")), 30000);

handlers["settings-sync-now"] = fromOwnPage(() => syncSettings());
handlers["settings-sync-undo"] = fromOwnPage(() => undoLastSettingsSync());
handlers["settings-sync-status"] = fromOwnPage(async () => ({
  ...settingsSyncStatus,
  backup: ((await messenger.storage.local.get(SYNC_BACKUP_KEY))[SYNC_BACKUP_KEY] || {}).at || "",
}));
