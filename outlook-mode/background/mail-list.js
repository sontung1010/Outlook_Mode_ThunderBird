"use strict";

// Message list additions (experiment outlookPin) and collected contacts
// (experiment outlookSync):
// - the cards view's star is a pin button (pin.js does the pinning);
// - a refresh button in the list header gets new mail, then syncs settings,
//   pins and contacts;
// - the people you write to or reply to go into Collected Addresses, which
//   settings sync (category "contacts") keeps the same on all devices.

const COLLECT_UNTIL_KEY = "contactsCollectedUntil"; // seconds; mail before this was scanned
const COLLECT_ALARM = "collect-contacts";
const COLLECT_MINUTES = 6 * 60;
const MARK_THREAD_DELAY = 100; // ms: right away, as Outlook; only lets a selection settle
const collectStatus = { lastRun: "", lastAdded: 0, lastMerged: 0, lastScanned: 0, people: 0, lastError: "" };

async function configureMailList() {
  const settings = await getSettings();
  await messenger.outlookPin.configureMailList({
    pinButton: Boolean(settings.pinButton),
    refreshButton: Boolean(settings.refreshButton),
    refreshKey: Boolean(settings.refreshKey),
    unreadHighlight: Boolean(settings.unreadHighlight),
    oneCardPerConversation: Boolean(settings.oneCardPerConversation),
    pinnedSection: Boolean(settings.pinnedSection),
    dateSections: Boolean(settings.dateSections),
    newestInCard: Boolean(settings.newestInCard),
  });
}

/* ------------------------------------------------------------------ */
/* Pin button                                                          */
/* ------------------------------------------------------------------ */

messenger.outlookPin.onPinButton.addListener(async (messageId) => {
  try {
    const message = await messenger.messages.get(messageId);
    await togglePin([message]);
    const [tab] = await messenger.tabs.query({ active: true, currentWindow: true });
    await resortAfterPin(tab).catch(() => {});
  } catch (error) {
    logError("pin button", error);
  }
});

/* ------------------------------------------------------------------ */
/* Collected contacts                                                  */
/* ------------------------------------------------------------------ */

let collectChain = Promise.resolve();

// full: scan all mail again, not only the mail since the last scan.
// A full scan already queued or running is reused, so asking again (or many
// times) does not stack more scans.
let fullCollect = null;

function collectContacts({ full = false } = {}) {
  if (full && fullCollect) {
    return fullCollect;
  }
  const run = collectChain.then(async () => {
    if (!(await getSettings()).collectContacts) {
      return { ok: false, error: "Collecting contacts is off." };
    }
    const stored = await messenger.storage.local.get(COLLECT_UNTIL_KEY);
    const since = full ? 0 : stored[COLLECT_UNTIL_KEY] || 0;
    // A day of overlap, for mail that arrives late (offline sync).
    const until = Math.floor(Date.now() / 1000) - 24 * 3600;
    try {
      const result = await messenger.outlookSync.collectContacts({ since });
      await messenger.storage.local.set({ [COLLECT_UNTIL_KEY]: until });
      Object.assign(collectStatus, {
        lastRun: new Date().toISOString(),
        lastAdded: result.added,
        lastMerged: result.merged,
        lastScanned: result.scanned,
        people: result.people,
        lastError: "",
      });
      return { ok: true, ...result };
    } catch (error) {
      collectStatus.lastError = String(error && error.message ? error.message : error);
      throw error;
    }
  });
  collectChain = run.catch(() => {});
  if (full) {
    fullCollect = run;
    run.finally(() => (fullCollect = null)).catch(() => {});
  }
  return run;
}

messenger.alarms.create(COLLECT_ALARM, { periodInMinutes: COLLECT_MINUTES });
messenger.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === COLLECT_ALARM) {
    collectContacts().catch((error) => logError("collect contacts", error));
  }
});
// After startup, once folders are loaded (the first run scans all mail).
setTimeout(() => collectContacts().catch((error) => logError("startup collect contacts", error)), 60000);

/* ------------------------------------------------------------------ */
/* Refresh button                                                      */
/* ------------------------------------------------------------------ */

let refreshing = false;

// One part of the refresh button's tooltip.
function refreshSummary(okText, offText, result) {
  if (result.status === "rejected") {
    const reason = result.reason && result.reason.message ? result.reason.message : result.reason;
    return { text: `${offText.split(" ")[0]}: ${reason}`, error: true };
  }
  const value = result.value || {};
  if (value.ok === false && /is off/.test(value.error || "")) {
    return { text: offText, error: false };
  }
  if (value.ok === false && value.waiting) {
    return { text: `${offText.split(" ")[0]}: waits for Microsoft's server`, error: true };
  }
  if (value.ok === false) {
    return { text: `${offText.split(" ")[0]}: ${value.error || (value.errors || []).join("; ") || "failed"}`, error: true };
  }
  return { text: okText, error: false };
}

// Every experiment scans the open windows again when configured.
async function reapplyUi() {
  const results = await Promise.allSettled([configureMailList(), configureReadingPane(), configureSearch()]);
  for (const result of results) {
    if (result.status === "rejected") {
      logError("reapply", result.reason);
    }
  }
}

// Thunderbird can still be opening its window (and its tabs) when the add-on
// starts; the additions are put in again a little later as a safety net.
for (const delay of [5000, 20000]) {
  setTimeout(() => reapplyUi().catch((error) => logError("reapply", error)), delay);
}

// Lists sorted "Pinned first" (or the Inbox with pinned mail on top) are
// sorted again, so conversations with new mail move up.
async function resortMailTabs() {
  for (const tab of await messenger.mailTabs.query({})) {
    await resortAfterPin({ id: tab.id, type: "mail" }).catch((error) => logError("refresh sort", error));
  }
}

// Folder list sync of each Microsoft Graph account; on success the
// account's new mail is fetched too. Returns parts of the tooltip.
async function restartFolderLists() {
  const parts = [];
  if (!messenger.outlookServer) {
    return parts;
  }
  for (const accountId of await graphAccountIds()) {
    const wasFailing = serverFailing(accountId);
    const result = await checkServer(accountId, { force: true });
    if (result.ok) {
      if (!wasFailing) {
        // (when it was failing, markBack() already did this)
        await messenger.outlookServer.getNewMail(accountId).catch((error) => logError("new mail", error));
      }
      parts.push({ text: wasFailing ? "Microsoft's server is back, folder list synced" : "folder list synced", error: false });
    } else if (result.network && !serverFailing(accountId)) {
      parts.push({ text: `folder list: ${result.reason}`, error: true });
    } else {
      parts.push({ text: `Microsoft's server is failing (${result.reason || "no answer"}); Thunderbird keeps trying`, error: true });
    }
  }
  return parts;
}

async function refreshAll() {
  if (refreshing) {
    return;
  }
  refreshing = true;
  // Also puts the add-on's additions back into every window, in case one
  // was missed (for example a window that was still opening).
  await reapplyUi();
  await messenger.outlookPin.setRefreshState({ busy: true, title: "Refreshing…" });
  try {
    await messenger.outlookPin.getNewMessages();
    // Thunderbird's "get new mail" does not run the folder list sync of a
    // Microsoft Graph account; run it here, which also starts a failing one
    // again (and tells whether Microsoft's server is back).
    const folderLists = await restartFolderLists();
    await resortMailTabs();
    const [settings, pins, contacts, sentCopies] = await Promise.allSettled([
      syncSettings(),
      syncAllPins({ force: true }),
      collectContacts(),
      checkSentCopies(),
    ]);
    const parts = [
      { text: "new mail checked", error: false },
      ...folderLists,
      refreshSummary("settings synced", "settings sync off", settings),
      refreshSummary(
        pins.value && pins.value.changed ? `pins synced (${pins.value.changed} changed)` : "pins synced",
        "pin sync off",
        pins
      ),
      refreshSummary(
        contacts.value && (contacts.value.added || contacts.value.merged)
          ? `contacts: ${contacts.value.added} added, ${contacts.value.merged} duplicate(s) merged`
          : "contacts up to date",
        "contacts collection off",
        contacts
      ),
      refreshSummary(
        sentCopies.value && sentCopies.value.moved
          ? `Sent Items: ${sentCopies.value.moved} second cop${sentCopies.value.moved === 1 ? "y" : "ies"} removed`
          : "Sent Items checked",
        "Sent Items check off",
        sentCopies
      ),
    ];
    const time = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    await messenger.outlookPin.setRefreshState({
      busy: false,
      error: parts.some((p) => p.error),
      title: `Refreshed at ${time}: ${parts.map((p) => p.text).join(", ")}.\nClick to refresh again.`,
    });
  } catch (error) {
    logError("refresh", error);
    await messenger.outlookPin.setRefreshState({ busy: false, error: true, title: `Refresh failed: ${error.message || error}` });
  } finally {
    refreshing = false;
  }
}

/* ------------------------------------------------------------------ */
/* Opening a mail reads its whole conversation, as in Outlook          */
/* ------------------------------------------------------------------ */

// One conversation selected (a click on its card, or the arrow keys): once
// it is shown, all its mail is marked read, as Outlook does for an opened
// conversation. A closed thread selects all its mail; the experiment
// leaves alone a selection of several conversations.
let markThreadTimer = null;
const markReadStatus = { last: "" }; // for Diagnostics

messenger.mailTabs.onSelectedMessagesChanged.addListener(async (tab, list) => {
  clearTimeout(markThreadTimer);
  const messages = (list && list.messages) || [];
  if (!messages.length || list.id) {
    return; // nothing, or more than a page of mail: not one conversation
  }
  const settings = await getSettings();
  if (!settings.markConversationRead) {
    return;
  }
  const ids = messages.map((message) => message.id);
  // As soon as the selection has settled: one click reads it all.
  markThreadTimer = setTimeout(async () => {
    try {
      const result = await messenger.outlookPin.markThreadRead(tab.id, ids);
      Object.assign(markReadStatus, { last: new Date().toLocaleString() }, result);
    } catch (error) {
      logError("mark conversation read", error);
    }
  }, MARK_THREAD_DELAY);
});

messenger.outlookPin.onRefresh.addListener(() => {
  refreshAll();
});

messenger.storage.onChanged.addListener((changes) => {
  if (!changes.settings) {
    return;
  }
  configureMailList().catch((error) => logError("mail list", error));
  const before = (changes.settings.oldValue || {}).collectContacts;
  const after = (changes.settings.newValue || {}).collectContacts;
  if (after && before === false) {
    collectContacts().catch((error) => logError("collect contacts", error));
  }
});

configureMailList().catch((error) => logError("mail list", error));

handlers["collect-contacts-now"] = fromOwnPage(() => collectContacts({ full: true }));
handlers["collect-contacts-status"] = fromOwnPage(async () => ({ ...collectStatus }));
handlers["refresh-now"] = fromOwnPage(() => refreshAll());
handlers["mark-read-status"] = fromOwnPage(async () => ({
  ...markReadStatus,
  threadCountsRepaired: await messenger.outlookPin.threadCountsRepaired().catch(() => null),
}));
