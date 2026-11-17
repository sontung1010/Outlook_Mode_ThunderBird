"use strict";

// Pin sync for Microsoft Graph accounts. Thunderbird keeps tags in its local
// database only, so on a Graph account the 📌 tag never reaches the server and
// other devices do not see it. This file mirrors pins to Outlook's own pin
// (the one Outlook on the web and mobile show) and reads them back:
// - pinning or unpinning writes the pin to the server right away; if that
//   fails (offline), the change waits in a queue and is retried;
// - at startup, every few minutes and when an Inbox is opened, the pins on the
//   server are read and the local 📌 tags are made to match.
//
// Outlook marks a pinned item with two MAPI properties set to a date in the
// year 4500: 0x0F01 and PR_RENEWTIME (0x0F02). Outlook sorts by 0x0F01, so
// unpinning sets it back to when the mail was received, and removes 0x0F02.
// Graph cannot remove an extended property, so 0x0F02 gets that date too.
//
// Reading the pins: Graph is asked for the messages whose 0x0F02 is in the
// far future. If it refuses that query, or answers nothing before a pin was
// ever seen, the newest mail of the Inbox is read instead (once an hour), plus
// every message changed since the last sync, in any folder: a pin or unpin
// changes lastModifiedDateTime, so pins made in Outlook on old mail come
// through too. Local pins outside what was read are left alone.

const PIN_PROP_FLAG = "SystemTime 0x0F02";
const PIN_PROP_TIME = "SystemTime 0x0F01";
const PIN_DATE = "4500-09-01T00:00:00Z";
const PIN_CUTOFF = "4000-01-01T00:00:00Z"; // a later 0x0F02 counts as pinned
const PIN_SYNC_ALARM = "pin-sync";
const PIN_SYNC_MINUTES = 5;
const PIN_SYNC_MIN_GAP_MS = 60 * 1000; // between syncs triggered by opening a folder
const MAX_SERVER_PINS = 2000;
const PIN_SCAN_LIMIT = 1000; // Inbox messages read when the query cannot be used
const PIN_INBOX_SCAN_MS = 60 * 60 * 1000;
const PIN_CHANGES_LIMIT = 5000;
// accountId -> { headerMessageId: pinned } of changes not yet on the server.
const PIN_SYNC_PENDING_KEY = "pinSyncPending";
// Accounts whose local pins were uploaded once, before the first download.
const PIN_SYNC_SEEDED_KEY = "pinSyncSeeded";
// Accounts where the server query has returned a pin at least once. Until
// then, local pins are never removed: an empty answer could also mean that
// the query does not work on this server.
const PIN_SYNC_VERIFIED_KEY = "pinSyncVerified";

// accountId -> { lastSync, serverPins, pending, lastError } (Diagnostics).
const pinSyncStatus = {};

function pinStatus(accountId) {
  pinSyncStatus[accountId] = pinSyncStatus[accountId] || {
    lastSync: "",
    serverPins: null,
    pending: 0,
    lastError: "",
  };
  return pinSyncStatus[accountId];
}

async function pinSyncEnabled() {
  return (await getSettings()).pinSync;
}

/* ------------------------------------------------------------------ */
/* Graph requests                                                      */
/* ------------------------------------------------------------------ */

// The server copies of a message (usually one; more if it is in several
// folders, e.g. mail sent to yourself).
// The Message-ID (PLAIN_MESSAGE_ID, common.js) is sent as an OData string
// literal (quotes doubled by odataString, then URL-encoded), so it cannot
// change the query; ids of another form (quotes, parentheses, spaces...) are
// not looked up at all.

async function graphItemsFor(accountId, headerMessageId, withPinProps = false) {
  if (typeof headerMessageId !== "string" || !PLAIN_MESSAGE_ID.test(headerMessageId)) {
    return [];
  }
  const params = new URLSearchParams({
    $filter: `internetMessageId eq ${odataString(`<${headerMessageId}>`)}`,
    $select: "id,subject,receivedDateTime",
  });
  if (withPinProps) {
    params.set(
      "$expand",
      `singleValueExtendedProperties($filter=id eq '${PIN_PROP_FLAG}' or id eq '${PIN_PROP_TIME}')`
    );
  }
  const result = await graphFetch(accountId, `/me/messages?${params}`);
  return result.value || [];
}

async function pushPin(accountId, headerMessageId, pinned) {
  const items = await graphItemsFor(accountId, headerMessageId);
  for (const item of items) {
    // Unpinned: back to its place among the mail of its day.
    const value = pinned ? PIN_DATE : item.receivedDateTime || new Date().toISOString();
    await graphFetch(accountId, `/me/messages/${encodeURIComponent(item.id)}`, {
      method: "PATCH",
      body: JSON.stringify({
        singleValueExtendedProperties: [
          { id: PIN_PROP_TIME, value },
          { id: PIN_PROP_FLAG, value },
        ],
      }),
    });
  }
  // No server copy (e.g. not uploaded yet): nothing to do.
}

// Graph names the properties its own way ("SystemTime 0xf02").
function propValue(item, id) {
  const key = (name) => name.toLowerCase().replace(/0x0*/, "0x");
  const prop = (item.singleValueExtendedProperties || []).find((p) => key(p.id) === key(id));
  return prop ? prop.value : null;
}

function isPinValue(value) {
  return Boolean(value) && Date.parse(value) > Date.parse(PIN_CUTOFF);
}

// All pinned messages, with a Graph query on the property.
async function queryPinnedIds(accountId) {
  const filter =
    `singleValueExtendedProperties/any(ep: ep/id eq '${PIN_PROP_FLAG}' and ` +
    `cast(ep/value, Edm.DateTimeOffset) gt ${PIN_CUTOFF})`;
  const params = new URLSearchParams({ $filter: filter, $select: "internetMessageId", $top: "100" });
  const ids = new Set();
  let next = `/me/messages?${params}`;
  while (next && ids.size < MAX_SERVER_PINS) {
    const page = await graphFetch(accountId, next);
    for (const item of page.value || []) {
      const [id] = messageIdsIn(item.internetMessageId);
      if (id) {
        ids.add(id);
      }
    }
    next = page["@odata.nextLink"];
  }
  return ids;
}

// The pinned messages among the newest of the Inbox, reading the property
// of each. `scope` is every message that was read.
async function scanPinnedIds(accountId) {
  const params = new URLSearchParams({
    $select: "internetMessageId",
    $expand: `singleValueExtendedProperties($filter=id eq '${PIN_PROP_FLAG}')`,
    $orderby: "receivedDateTime desc",
    $top: "100",
  });
  const ids = new Set();
  const scope = new Set();
  let next = `/me/mailFolders/inbox/messages?${params}`;
  while (next && scope.size < PIN_SCAN_LIMIT) {
    const page = await graphFetch(accountId, next);
    for (const item of page.value || []) {
      const [id] = messageIdsIn(item.internetMessageId);
      if (!id) {
        continue;
      }
      scope.add(id);
      if ((item.singleValueExtendedProperties || []).some((prop) => isPinValue(prop.value))) {
        ids.add(id);
      }
    }
    next = page["@odata.nextLink"];
  }
  return { ids, scope };
}

function isoSeconds(ms) {
  return new Date(ms).toISOString().replace(/\.\d+Z$/, "Z");
}

// Messages changed since `since` (ms), in every folder, with the pin
// properties. Graph keeps lastModifiedDateTime for any change of a message.
async function modifiedSince(accountId, since, select, limit) {
  const params = new URLSearchParams({
    $filter: `lastModifiedDateTime ge ${isoSeconds(since)}`,
    $select: select,
    $expand: `singleValueExtendedProperties($filter=id eq '${PIN_PROP_FLAG}' or id eq '${PIN_PROP_TIME}')`,
    $top: "200",
  });
  const items = [];
  let next = `/me/messages?${params}`;
  while (next && items.length < limit) {
    const page = await graphFetch(accountId, next);
    items.push(...(page.value || []));
    next = page["@odata.nextLink"];
  }
  return items;
}

// The pins on the server: { ids, scope, method }. scope is null when ids
// covers the whole mailbox.
async function serverPins(accountId, verified) {
  const status = pinStatus(accountId);
  if (status.method !== "scan") {
    try {
      const ids = await queryPinnedIds(accountId);
      // An empty answer before any pin was ever seen may mean the query does
      // not match on this server: check the Inbox itself.
      if (ids.size || verified) {
        status.method = "query";
        return { ids, scope: null, method: "query" };
      }
    } catch (error) {
      if (!/^Graph 4\d\d/.test(error.message || "") || /^Graph 401/.test(error.message || "")) {
        throw error; // offline, signed out or a server error: not the query's fault
      }
    }
  }
  // The newest Inbox mail now and then, and every change since last time.
  const startedAt = Date.now();
  const fullScan = !status.inboxScannedAt || startedAt - status.inboxScannedAt > PIN_INBOX_SCAN_MS;
  const scan = fullScan ? await scanPinnedIds(accountId) : { ids: new Set(), scope: new Set() };
  const since = status.changesReadAt ? status.changesReadAt - 10 * 60 * 1000 : startedAt - PIN_REPAIR_DAYS * 864e5;
  for (const item of await modifiedSince(accountId, since, "internetMessageId", PIN_CHANGES_LIMIT)) {
    const [id] = messageIdsIn(item.internetMessageId);
    if (!id) {
      continue;
    }
    scan.scope.add(id);
    if (isPinValue(propValue(item, PIN_PROP_FLAG))) {
      scan.ids.add(id);
    } else {
      scan.ids.delete(id);
    }
  }
  if (fullScan) {
    status.inboxScannedAt = startedAt;
  }
  status.changesReadAt = startedAt;
  if (scan.ids.size || status.method === "scan") {
    status.method = "scan";
  }
  return { ...scan, method: "scan" };
}

/* ------------------------------------------------------------------ */
/* Repair: mail an earlier unpin moved to the top                      */
/* ------------------------------------------------------------------ */

// Before v3.7.0, unpinning set both properties to the time of the unpin,
// not back to when the mail was received. Outlook sorts by them, so the
// mail jumped to the top of the list (Outlook mobile also shows a clock on
// it, as on mail back from Snooze). Such mail has both properties equal to
// the second, not in the year 4500, later than its receipt, and recent.
// They are set back to the received date, once per account.
// Renamed when the search got better, so that it runs once more.
const PIN_REPAIR_KEY = "pinUnpinDatesRepaired2";
const PIN_REPAIR_DAYS = 14;

function unpinDamaged(item) {
  const renew = propValue(item, PIN_PROP_FLAG);
  const time = propValue(item, PIN_PROP_TIME);
  if (!renew || !time || isPinValue(renew)) {
    return false;
  }
  const renewAt = Date.parse(renew);
  return (
    Math.abs(Date.parse(time) - renewAt) < 2000 &&
    renewAt - Date.parse(item.receivedDateTime) > 60 * 1000 &&
    renewAt > Date.now() - PIN_REPAIR_DAYS * 864e5
  );
}

// The repair writes lastModifiedDateTime too: every message changed in the
// last days, in any folder and of any age, is checked.
async function damagedUnpins(accountId) {
  const changed = await modifiedSince(
    accountId,
    Date.now() - PIN_REPAIR_DAYS * 864e5,
    "id,receivedDateTime",
    10 * PIN_CHANGES_LIMIT
  );
  return changed.filter(unpinDamaged);
}

async function repairUnpinDates({ force = false } = {}) {
  const repaired = (await messenger.storage.local.get(PIN_REPAIR_KEY))[PIN_REPAIR_KEY] || [];
  let fixed = 0;
  const errors = [];
  for (const accountId of await graphAccountIds()) {
    if (!force && repaired.includes(accountId)) {
      continue;
    }
    try {
      for (const item of await damagedUnpins(accountId)) {
        await graphFetch(accountId, `/me/messages/${encodeURIComponent(item.id)}`, {
          method: "PATCH",
          body: JSON.stringify({
            singleValueExtendedProperties: [
              { id: PIN_PROP_TIME, value: item.receivedDateTime },
              { id: PIN_PROP_FLAG, value: item.receivedDateTime },
            ],
          }),
        });
        fixed++;
      }
      if (!repaired.includes(accountId)) {
        repaired.push(accountId);
      }
    } catch (error) {
      errors.push(`${accountId}: ${error.message || error}`);
      logError(`repair unpin dates ${accountId}`, error);
    }
  }
  await messenger.storage.local.set({ [PIN_REPAIR_KEY]: repaired });
  if (fixed) {
    notify(
      "Outlook Mode",
      `${fixed} mail(s) that an earlier unpin had moved to the top in Outlook are back in their place.`
    );
  }
  return { ok: !errors.length, fixed, errors };
}

/* ------------------------------------------------------------------ */
/* Queue and sync                                                      */
/* ------------------------------------------------------------------ */

// One sync step at a time, so a download never undoes an upload in flight.
let pinSyncChain = Promise.resolve();
function withPinSyncLock(task) {
  const run = pinSyncChain.then(task, task);
  pinSyncChain = run.catch(() => {});
  return run;
}

async function getPending() {
  const stored = await messenger.storage.local.get(PIN_SYNC_PENDING_KEY);
  return stored[PIN_SYNC_PENDING_KEY] || {};
}

async function addPending(accountId, changes) {
  const pending = await getPending();
  pending[accountId] = { ...(pending[accountId] || {}), ...changes };
  await messenger.storage.local.set({ [PIN_SYNC_PENDING_KEY]: pending });
}

// Upload the queued changes of one account. Changes that fail stay queued.
// At most this many pin changes are uploaded per run, one request at a time;
// the rest stay queued for the next run (every PIN_SYNC_MINUTES).
const PIN_FLUSH_MAX = 100;

async function flushPending(accountId) {
  const queued = (await getPending())[accountId] || {};
  const done = [];
  let failure = null;
  for (const [headerMessageId, pinned] of Object.entries(queued).slice(0, PIN_FLUSH_MAX)) {
    try {
      await pushPin(accountId, headerMessageId, pinned);
      done.push(headerMessageId);
    } catch (error) {
      failure = error;
      break; // most likely offline or signed out; retry on the next sync
    }
  }
  // Re-read: new changes may have been queued meanwhile.
  const pending = await getPending();
  const left = { ...(pending[accountId] || {}) };
  for (const id of done) {
    if (left[id] === queued[id]) {
      delete left[id];
    }
  }
  pending[accountId] = left;
  await messenger.storage.local.set({ [PIN_SYNC_PENDING_KEY]: pending });
  pinStatus(accountId).pending = Object.keys(left).length;
  if (failure) {
    throw failure;
  }
  return left;
}

async function localPinned(accountId) {
  return collectList(
    await messenger.messages.query({
      accountId,
      tags: { mode: "any", tags: { [PIN_TAG_KEY]: true } },
    })
  );
}

async function setLocalTag(message, pinned) {
  const tags = new Set(message.tags || []);
  if (tags.has(PIN_TAG_KEY) === pinned) {
    return false;
  }
  if (pinned) {
    tags.add(PIN_TAG_KEY);
  } else {
    tags.delete(PIN_TAG_KEY);
  }
  await messenger.messages.update(message.id, { tags: [...tags] });
  return true;
}

// Upload, then make the local tags match the server.
async function syncAccount(accountId) {
  const status = pinStatus(accountId);
  try {
    await ensureTag(PIN_TAG_KEY, PIN_TAG_NAME, PIN_TAG_COLOR);
    const seeded = (await messenger.storage.local.get(PIN_SYNC_SEEDED_KEY))[PIN_SYNC_SEEDED_KEY] || [];
    if (!seeded.includes(accountId)) {
      // First sync on this device: keep the pins made before sync existed.
      const changes = {};
      for (const message of await localPinned(accountId)) {
        changes[message.headerMessageId] = true;
      }
      await addPending(accountId, changes);
    }
    await flushPending(accountId);
    if (!seeded.includes(accountId)) {
      await messenger.storage.local.set({ [PIN_SYNC_SEEDED_KEY]: [...seeded, accountId] });
    }

    const verified =
      (await messenger.storage.local.get(PIN_SYNC_VERIFIED_KEY))[PIN_SYNC_VERIFIED_KEY] || [];
    const pins = await serverPins(accountId, verified.includes(accountId));
    const onServer = pins.ids;
    const canRemove = verified.includes(accountId) || onServer.size > 0;
    if (onServer.size && !verified.includes(accountId)) {
      await messenger.storage.local.set({ [PIN_SYNC_VERIFIED_KEY]: [...verified, accountId] });
    }
    // Changes not on the server yet win, including ones made during this sync.
    const left = (await getPending())[accountId] || {};
    let changed = 0;
    for (const message of canRemove ? await localPinned(accountId) : []) {
      const id = message.headerMessageId;
      const known = !pins.scope || pins.scope.has(id); // the scan only read part of the mail
      if (known && !onServer.has(id) && !(id in left) && (await setLocalTag(message, false))) {
        changed++;
      }
    }
    for (const id of onServer) {
      if (id in left) {
        continue;
      }
      const copies = await collectList(
        await messenger.messages.query({ accountId, headerMessageId: id }),
        10
      );
      for (const message of copies) {
        if (await setLocalTag(message, true)) {
          changed++;
        }
      }
    }
    status.lastSync = new Date().toISOString();
    status.serverPins = onServer.size;
    status.method = pins.method;
    status.lastError = canRemove ? "" : "the server shows no pins yet; local pins are kept";
    return changed;
  } catch (error) {
    status.lastError = String(error && error.message ? error.message : error);
    throw error;
  }
}

let lastPinSync = 0;

async function syncAllPins({ force = false } = {}) {
  if (!(await pinSyncEnabled())) {
    return { ok: false, error: "Pin sync is off." };
  }
  if (!force && Date.now() - lastPinSync < PIN_SYNC_MIN_GAP_MS) {
    return { ok: true, skipped: true };
  }
  // While Microsoft's server is failing, no new sync starts (it would only
  // wait); the server coming back runs one (serverBackHooks below).
  if (await anyServerFailing()) {
    return { ok: false, waiting: true, error: "Waiting for Microsoft's server to come back." };
  }
  lastPinSync = Date.now();
  return withPinSyncLock(async () => {
    let changed = 0;
    const errors = [];
    for (const accountId of await graphAccountIds()) {
      try {
        changed += await syncAccount(accountId);
      } catch (error) {
        errors.push(`${accountId}: ${error.message || error}`);
        logError(`pin sync ${accountId}`, error);
      }
    }
    if (changed) {
      for (const mailTab of await messenger.mailTabs.query({})) {
        await resortAfterPin({ id: mailTab.id, type: "mail" }).catch(() => {});
      }
    }
    return { ok: !errors.length, changed, errors };
  });
}

// Called by setPinned() after the local tags changed. Does not wait for the
// server, so pinning stays instant.
function queuePinSync(messages, pinned) {
  (async () => {
    if (!(await pinSyncEnabled())) {
      return;
    }
    const graph = new Set(await graphAccountIds());
    const byAccount = new Map();
    for (const given of messages) {
      // Thunderbird's own copy: the message must exist in one of the user's
      // accounts; its account and Message-ID come from there.
      const message = given && Number.isInteger(given.id) ? await messenger.messages.get(given.id).catch(() => null) : null;
      const accountId = message && message.folder && message.folder.accountId;
      // Uploaded only when the local tag already says so: the server follows
      // what Thunderbird shows, nothing else.
      if (accountId && graph.has(accountId) && message.headerMessageId && isPinned(message) === pinned) {
        byAccount.set(accountId, { ...(byAccount.get(accountId) || {}), [message.headerMessageId]: pinned });
      }
    }
    for (const [accountId, changes] of byAccount) {
      await addPending(accountId, changes);
      if (serverFailing(accountId)) {
        continue; // kept in the queue, sent when the server is back
      }
      await withPinSyncLock(() => flushPending(accountId)).catch((error) => {
        pinStatus(accountId).lastError = String(error.message || error);
        logError(`pin upload ${accountId}`, error);
      });
    }
  })().catch((error) => logError("pin sync queue", error));
}

/* ------------------------------------------------------------------ */
/* Triggers                                                            */
/* ------------------------------------------------------------------ */

// The server is back: pins changed meanwhile (here and in Outlook) and the
// queued uploads go now.
serverBackHooks.push(() => syncAllPins({ force: true }));

messenger.alarms.create(PIN_SYNC_ALARM, { periodInMinutes: PIN_SYNC_MINUTES });
messenger.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === PIN_SYNC_ALARM) {
    syncAllPins({ force: true }).catch((error) => logError("pin sync alarm", error));
  }
});

messenger.mailTabs.onDisplayedFolderChanged.addListener((tab, folder) => {
  if (isInbox(folder)) {
    syncAllPins().catch((error) => logError("pin sync on folder open", error));
  }
});

// Give Thunderbird a moment to sign in and load folders after startup. The
// repair runs even with pin sync off: it undoes this add-on's own mistake.
setTimeout(() => {
  withPinSyncLock(() => repairUnpinDates())
    .catch((error) => logError("repair unpin dates", error))
    .then(() => syncAllPins({ force: true }))
    .catch((error) => logError("startup pin sync", error));
}, 15000);

/* ------------------------------------------------------------------ */
/* Diagnostics                                                         */
/* ------------------------------------------------------------------ */

// The pin properties on the server for the messages selected in the mail
// tab, to compare with what Outlook on the web shows.
async function checkSelectedPins() {
  const [mailTab] = await messenger.mailTabs.query({ currentWindow: true });
  if (!mailTab) {
    throw userError("Open a mail folder and select a message first.");
  }
  const messages = (await collectList(await messenger.mailTabs.getSelectedMessages(mailTab.id))).slice(0, 10);
  if (!messages.length) {
    throw userError("Select a message in the mail tab first.");
  }
  const graph = new Set(await graphAccountIds());
  const out = [];
  for (const message of messages) {
    const row = { subject: message.subject, localPinned: isPinned(message), server: [] };
    const accountId = message.folder && message.folder.accountId;
    if (!graph.has(accountId)) {
      row.error = "not a Microsoft Graph account";
    } else {
      try {
        // The pin properties, and what else Outlook may use, to compare a
        // mail pinned in Outlook with one pinned here.
        const ids = [PIN_PROP_TIME, PIN_PROP_FLAG, "Integer 0x1090", "SystemTime 0x0E06", "SystemTime 0x1091"];
        const params = new URLSearchParams({
          $filter: `internetMessageId eq ${odataString(`<${message.headerMessageId}>`)}`,
          $select: "id,receivedDateTime,flag,parentFolderId,inferenceClassification",
          $expand: `singleValueExtendedProperties($filter=${ids.map((id) => `id eq '${id}'`).join(" or ")})`,
        });
        for (const item of (await graphFetch(accountId, `/me/messages?${params}`)).value || []) {
          const props = {
            received: item.receivedDateTime,
            flag: item.flag && item.flag.flagStatus,
            focused: item.inferenceClassification,
          };
          for (const prop of item.singleValueExtendedProperties || []) {
            props[prop.id] = prop.value;
          }
          row.server.push(props);
        }
      } catch (error) {
        row.error = String(error.message || error);
      }
    }
    out.push(row);
  }
  return out;
}

handlers["pin-sync-check"] = fromOwnPage(() => checkSelectedPins());
handlers["pin-sync-now"] = fromOwnPage(() => syncAllPins({ force: true }));
