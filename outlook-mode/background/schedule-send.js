"use strict";

// Schedule send: the message is saved as a draft and the compose window is
// closed. At the chosen time the draft is opened in a compose window, sent,
// and then deleted from Drafts. Thunderbird has to be running at that time;
// mail that became due while it was closed is sent at the next start.

const SCHEDULE_KEY = "scheduled";
const ALARM_PREFIX = "olsched:";
const TICK_ALARM = "olsched-tick";
let scheduleBusy = Promise.resolve();

async function getScheduled() {
  const stored = await messenger.storage.local.get(SCHEDULE_KEY);
  return Array.isArray(stored[SCHEDULE_KEY]) ? stored[SCHEDULE_KEY] : [];
}

async function saveScheduled(list) {
  await messenger.storage.local.set({ [SCHEDULE_KEY]: list });
}

// Serialize every change to the scheduled list.
function withScheduleLock(fn) {
  const run = scheduleBusy.then(fn, fn);
  scheduleBusy = run.catch(() => {});
  return run;
}

function newId() {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

const SCHEDULE_MAX_MS = 366 * 864e5;

async function scheduleFromTab(tabId, sendAt) {
  const when = new Date(sendAt);
  if (Number.isNaN(when.getTime())) {
    throw userError("Invalid date.");
  }
  if (when.getTime() < Date.now() + 30 * 1000) {
    throw userError("Choose a time in the future.");
  }
  if (when.getTime() > Date.now() + SCHEDULE_MAX_MS) {
    throw userError("Choose a time within a year.");
  }
  const details = await messenger.compose.getComposeDetails(tabId);
  if (![...(details.to || []), ...(details.cc || []), ...(details.bcc || [])].length) {
    throw userError("Add at least one recipient first.");
  }
  await tellCompose(tabId, { type: "ui-cleanup" });
  const saved = await messenger.compose.saveMessage(tabId, { mode: "draft" });
  const draft = saved && saved.messages && saved.messages[0];
  if (!draft || !draft.headerMessageId) {
    throw new Error("Could not save the message as a draft.");
  }
  const item = {
    id: newId(),
    headerMessageId: draft.headerMessageId,
    sendAt: when.toISOString(),
    subject: details.subject || "(no subject)",
    to: [...(details.to || []), ...(details.cc || [])].join(", "),
    status: "pending",
    error: "",
  };
  await withScheduleLock(async () => {
    const list = await getScheduled();
    list.push(item);
    await saveScheduled(list);
  });
  messenger.alarms.create(ALARM_PREFIX + item.id, { when: when.getTime() });
  // The draft is saved, so the window closes without asking.
  await messenger.tabs.remove(tabId).catch((error) => logError("schedule: close window", error));
  return { ok: true, item };
}

async function findDraft(headerMessageId) {
  const found = await collectList(await messenger.messages.query({ headerMessageId }), 20);
  return (
    found.find((m) => m.folder && (m.folder.specialUse || []).includes("drafts")) || found[0]
  );
}

async function waitForComposeReady(tabId) {
  for (let i = 0; i < 50; i++) {
    try {
      const state = await messenger.compose.getComposeState(tabId);
      if (state && state.canSendNow) {
        return;
      }
    } catch (error) {
      // window still loading
    }
    await new Promise((resolve) => setTimeout(resolve, 200));
  }
}

async function sendScheduledItem(item) {
  const draft = await findDraft(item.headerMessageId);
  if (!draft) {
    throw new Error("The draft was not found. It may have been edited, sent or deleted.");
  }
  const tab = await messenger.compose.beginNew(draft.id);
  await waitForComposeReady(tab.id);
  await sendNowBypassingUndo(tab.id, "sendNow");
  await messenger.messages.delete([draft.id], true).catch((error) =>
    logError("schedule: delete draft", error)
  );
}

function runDueScheduled() {
  return withScheduleLock(async () => {
    const list = await getScheduled();
    const now = Date.now();
    let changed = false;
    for (const item of list) {
      if (item.status !== "pending" || new Date(item.sendAt).getTime() > now) {
        continue;
      }
      try {
        await sendScheduledItem(item);
        item.status = "sent";
        notify("Scheduled message sent", item.subject);
      } catch (error) {
        item.status = "failed";
        item.error = String(error && error.message ? error.message : error);
        logError("schedule: send", error);
        notify("Scheduled message NOT sent", `${item.subject}: ${item.error}`);
      }
      changed = true;
    }
    if (changed) {
      await saveScheduled(list.filter((item) => item.status !== "sent"));
    }
  });
}

async function isScheduled(id) {
  return typeof id === "string" && (await getScheduled()).some((item) => item.id === id);
}

async function cancelScheduled(id) {
  if (!(await isScheduled(id))) {
    return { ok: false };
  }
  await withScheduleLock(async () => {
    const list = await getScheduled();
    await saveScheduled(list.filter((item) => item.id !== id));
  });
  await messenger.alarms.clear(ALARM_PREFIX + id);
  return { ok: true };
}

async function sendScheduledNow(id) {
  if (!(await isScheduled(id))) {
    return { ok: false };
  }
  await withScheduleLock(async () => {
    const list = await getScheduled();
    const item = list.find((i) => i.id === id);
    if (item) {
      item.sendAt = new Date().toISOString();
      item.status = "pending";
      item.error = "";
      await saveScheduled(list);
    }
  });
  await runDueScheduled();
  return { ok: true };
}

async function restoreScheduleAlarms() {
  const list = await getScheduled();
  for (const item of list) {
    if (item.status === "pending") {
      messenger.alarms.create(ALARM_PREFIX + item.id, {
        when: Math.max(Date.now() + 1000, new Date(item.sendAt).getTime()),
      });
    }
  }
  // Safety net in case an alarm is missed (for example after suspend).
  messenger.alarms.create(TICK_ALARM, { periodInMinutes: 1 });
}

messenger.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === TICK_ALARM || alarm.name.startsWith(ALARM_PREFIX)) {
    runDueScheduled().catch((error) => logError("schedule: alarm", error));
  }
});

// Each from the pages that use it: the Send later popup schedules, the
// Quick Steps popup and Preferences list, and only Preferences (where the
// user sees the list) cancels or sends one now.
handlers["schedule-send"] = fromPages(["popup/schedule.html"], (message) => scheduleFromTab(message.tabId, message.sendAt));
handlers["schedule-list"] = fromPages(["options/options.html", "popup/quicksteps.html"], () => getScheduled());
handlers["schedule-cancel"] = fromPages(["options/options.html"], (message) => cancelScheduled(message.id));
handlers["schedule-send-now"] = fromPages(["options/options.html"], (message) => sendScheduledNow(message.id));
