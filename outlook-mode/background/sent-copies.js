"use strict";

// One copy of sent mail in Sent Items, on Microsoft Graph accounts.
//
// Thunderbird sends Graph mail with POST /me/messages and /send, and Exchange
// keeps the sent message in Sent Items itself. Thunderbird also filed its own
// copy there (Fcc), uploaded as a new message: every mail sent from
// Thunderbird was in Sent Items twice, and Exchange stores the uploaded copy
// as an unsent draft, which Outlook (web, Windows, mobile) shows as "[Draft]"
// in the conversation, although the mail was sent.
// - outlookPane.setSingleSentCopy: from now on, sends (reading pane, compose
//   dock, compose window) file no copy of their own when Exchange keeps one;
// - the second copies already there (made before, or by a device without
//   this add-on) are moved to Deleted Items: in Sent Items, messages with the
//   same Message-ID as another one there. The one kept is the oldest that is
//   not a draft (Exchange's own copy). Checked at startup, every 30 minutes,
//   on ⟳ Refresh and when Microsoft's server is back.

const SENT_COPIES_ALARM = "sent-copies";
const SENT_COPIES_MINUTES = 30;
// accountId -> ISO time of the last complete check; the next one reads only
// what was created since (minus a margin: the two copies of a mail are
// created seconds apart).
const SENT_COPIES_CHECKED_KEY = "sentCopiesChecked";
const SENT_COPIES_MARGIN_MS = 60 * 60 * 1000;
const SENT_COPIES_SCAN_LIMIT = 3000; // Sent Items messages read in one check
// The two copies of a sent mail are created a few seconds apart; copies
// further apart are left alone (for example a mail moved back to Sent Items).
const SENT_COPIES_MAX_GAP_MS = 10 * 60 * 1000;

// accountId -> { lastCheck, scanned, moved, lastError } (Diagnostics).
const sentCopiesStatus = {};

function sentCopyStatus(accountId) {
  sentCopiesStatus[accountId] = sentCopiesStatus[accountId] || { lastCheck: "", scanned: 0, moved: 0, lastError: "" };
  return sentCopiesStatus[accountId];
}

async function configureSentCopies() {
  const settings = await getSettings();
  try {
    await messenger.outlookPane.setSingleSentCopy(settings.singleSentCopy !== false);
  } catch (error) {
    logError("sent copies", error);
  }
}

// The Sent Items messages created since `since` (ISO, or "" for all, newest
// first, up to SENT_COPIES_SCAN_LIMIT).
async function sentItems(accountId, since) {
  const params = new URLSearchParams({
    $select: "id,internetMessageId,isDraft,createdDateTime",
    $orderby: "createdDateTime desc",
    $top: "100",
  });
  if (since) {
    params.set("$filter", `createdDateTime ge ${since}`);
  }
  let url = `/me/mailFolders/sentitems/messages?${params}`;
  const items = [];
  while (url && items.length < SENT_COPIES_SCAN_LIMIT) {
    const page = await graphFetch(accountId, url);
    items.push(...((page && page.value) || []));
    url = page && page["@odata.nextLink"];
  }
  return items;
}

// The second copies among Sent Items messages: same Message-ID as another
// one, created within SENT_COPIES_MAX_GAP_MS of the one kept (the oldest that
// is not a draft, else the oldest).
function secondSentCopies(items) {
  const groups = new Map();
  for (const item of items) {
    const key = typeof item.internetMessageId === "string" ? item.internetMessageId.trim() : "";
    if (key && item.id) {
      groups.set(key, [...(groups.get(key) || []), item]);
    }
  }
  const time = (item) => Date.parse(item.createdDateTime) || 0;
  const extra = [];
  for (const group of groups.values()) {
    if (group.length < 2) {
      continue;
    }
    group.sort((a, b) => Number(Boolean(a.isDraft)) - Number(Boolean(b.isDraft)) || time(a) - time(b));
    const [kept, ...rest] = group;
    extra.push(...rest.filter((item) => Math.abs(time(item) - time(kept)) <= SENT_COPIES_MAX_GAP_MS));
  }
  return extra;
}

async function checkSentCopiesOf(accountId, checked) {
  const status = sentCopyStatus(accountId);
  const startedAt = new Date().toISOString();
  const last = Date.parse(checked[accountId] || "");
  const since = Number.isFinite(last) ? new Date(last - SENT_COPIES_MARGIN_MS).toISOString().replace(/\.\d+Z$/, "Z") : "";
  try {
    const items = await sentItems(accountId, since);
    const extra = secondSentCopies(items);
    let moved = 0;
    for (const item of extra) {
      await graphFetch(accountId, `/me/messages/${encodeURIComponent(item.id)}/move`, {
        method: "POST",
        body: JSON.stringify({ destinationId: "deleteditems" }),
      });
      moved++;
    }
    status.lastCheck = startedAt;
    status.scanned = items.length;
    status.moved += moved;
    status.lastError = "";
    checked[accountId] = startedAt;
    return moved;
  } catch (error) {
    status.lastError = String(error && error.message ? error.message : error);
    throw error;
  }
}

let sentCopiesRunning = null;

// Returns { ok, moved, errors } (or { ok: false, error } when off).
function checkSentCopies() {
  if (sentCopiesRunning) {
    return sentCopiesRunning;
  }
  sentCopiesRunning = (async () => {
    if ((await getSettings()).singleSentCopy === false) {
      return { ok: false, error: "One copy of sent mail is off." };
    }
    if (await anyServerFailing()) {
      return { ok: false, waiting: true, error: "Waiting for Microsoft's server to come back." };
    }
    const checked = (await messenger.storage.local.get(SENT_COPIES_CHECKED_KEY))[SENT_COPIES_CHECKED_KEY] || {};
    let moved = 0;
    const errors = [];
    for (const accountId of await graphAccountIds()) {
      try {
        moved += await checkSentCopiesOf(accountId, checked);
      } catch (error) {
        errors.push(`${accountId}: ${error.message || error}`);
        logError(`sent copies ${accountId}`, error);
      }
    }
    await messenger.storage.local.set({ [SENT_COPIES_CHECKED_KEY]: checked });
    if (moved) {
      notify(
        "Sent Items: second copies removed",
        `${moved} mail(s) were twice in Sent Items (the second copy showed as a Draft in Outlook). ` +
          "The second copies were moved to Deleted Items."
      );
    }
    return { ok: !errors.length, moved, errors };
  })().finally(() => {
    sentCopiesRunning = null;
  });
  return sentCopiesRunning;
}

/* ------------------------------------------------------------------ */
/* Triggers                                                            */
/* ------------------------------------------------------------------ */

messenger.storage.onChanged.addListener((changes) => {
  if (changes.settings) {
    configureSentCopies();
  }
});

serverBackHooks.push(() => checkSentCopies());

messenger.alarms.create(SENT_COPIES_ALARM, { periodInMinutes: SENT_COPIES_MINUTES });
messenger.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === SENT_COPIES_ALARM) {
    checkSentCopies().catch((error) => logError("sent copies alarm", error));
  }
});

configureSentCopies();
// After Thunderbird has signed in and loaded its folders.
setTimeout(() => {
  checkSentCopies().catch((error) => logError("startup sent copies", error));
}, 20000);

