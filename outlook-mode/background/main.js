"use strict";

// Context menus, message routing and startup. Loaded last.

/* ------------------------------------------------------------------ */
/* Menus                                                               */
/* ------------------------------------------------------------------ */

// Opens the Settings tab, scrolled to a section (an id in options.html). A
// tab that is already open scrolls on the message; a new one asks for the
// section when it loads.
let pendingSettingsSection = "";
async function openSettings(section = "") {
  pendingSettingsSection = section;
  await messenger.runtime.openOptionsPage();
  messenger.runtime.sendMessage({ type: "show-settings-section", section }).catch(() => {});
}

handlers["open-settings"] = fromOwnPage((message) => openSettings(message.section || ""));
handlers["take-settings-section"] = fromOwnPage(() => {
  const section = pendingSettingsSection;
  pendingSettingsSection = "";
  return section;
});

async function rebuildMenus() {
  await messenger.menus.removeAll();
  messenger.menus.create({ id: "ol-settings", title: "Outlook Mode preferences…", contexts: ["tools_menu"] });
  messenger.menus.create({ id: "ol-cleanup", title: "🧹 Clean up mailbox…", contexts: ["tools_menu"] });
  const contexts = ["message_list"];
  messenger.menus.create({ id: "ol-pin", title: "📌 Pin / Unpin  (Alt+Shift+P)", contexts });
  messenger.menus.create({
    id: "ol-move",
    title: "📁 Move to folder…  (Ctrl+Shift+V)",
    contexts,
  });
  messenger.menus.create({
    id: "ol-forward-attachment",
    title: "📎 Forward as attachment  (Ctrl+Alt+F)",
    contexts,
  });
  const steps = await getSteps();
  messenger.menus.create({ id: "ol-qs", title: "⚡ Quick Steps", contexts });
  steps.forEach((step, index) => {
    const key = index < 5 ? `  (Ctrl+Shift+${index + 1})` : "";
    messenger.menus.create({
      id: `ol-qs-${index}`,
      parentId: "ol-qs",
      // Plain text label; markup characters are removed all the same.
      title: `${String(step.name || "").replace(/[<>&"'`]/g, "").trim() || `Step ${index + 1}`}${key}`,
      contexts,
    });
  });
  messenger.menus.create({ id: "ol-qs-sep", parentId: "ol-qs", type: "separator", contexts });
  messenger.menus.create({
    id: "ol-qs-manage",
    parentId: "ol-qs",
    title: "Manage Quick Steps…",
    contexts,
  });
}

messenger.menus.onClicked.addListener(async (info, tab) => {
  try {
    const id = String(info.menuItemId);
    if (id === "ol-qs-manage") {
      await openSettings("quick-steps");
      return;
    }
    if (id === "ol-settings") {
      await openSettings();
      return;
    }
    if (id === "ol-cleanup") {
      await openCleanup();
      return;
    }
    const messages = await collectList(info.selectedMessages);
    if (id === "ol-pin") {
      await togglePin(messages);
      await resortIfInbox(tab).catch(() => {});
    } else if (id === "ol-move") {
      await openMovePicker(tab);
    } else if (id === "ol-forward-attachment") {
      await forwardAsAttachment(messages);
    } else if (id.startsWith("ol-qs-")) {
      await runStepOnMessages(Number(id.slice("ol-qs-".length)), messages, tab);
    }
  } catch (error) {
    logError("menu action", error);
  }
});

messenger.storage.onChanged.addListener((changes) => {
  if (changes.quickSteps) {
    rebuildMenus().catch((error) => logError("menus", error));
  }
});

/* ------------------------------------------------------------------ */
/* Diagnostics                                                         */
/* ------------------------------------------------------------------ */

handlers.diagnostics = fromOwnPage(async (message, sender) => {
  await composeScriptReady;
  let contacts = contactStats;
  try {
    await getContacts();
    contacts = contactStats;
  } catch (error) {
    // contactStats.error is set
  }
  const browserInfo = await messenger.runtime.getBrowserInfo();
  return {
    version: messenger.runtime.getManifest().version,
    thunderbird: `${browserInfo.name} ${browserInfo.version}`,
    boot: bootStatus,
    composeWindowsSeen: composeScriptPings,
    contacts,
    pinSync: pinSyncStatus,
    settingsSync: settingsSyncStatus,
    readingPane: await handlers["reading-pane-status"](message, sender),
    markRead: await handlers["mark-read-status"](message, sender),
    server: {
      ...serverHealthStatus(),
      watch: messenger.outlookServer
        ? await messenger.outlookServer.status().catch((error) => ({ lastError: String(error.message || error) }))
        : null,
    },
    errors: recentErrors,
  };
});

// The compose script says hello when it starts in a compose window.
let composeScriptPings = 0;
handlers["compose-hello"] = () => {
  composeScriptPings++;
  return { ok: true };
};

handlers["reload-contacts"] = fromOwnPage(async () => {
  invalidateContacts();
  await getContacts();
  return contactStats;
});

/* ------------------------------------------------------------------ */
/* Message routing                                                     */
/* ------------------------------------------------------------------ */

// Messages come only from this add-on (runtime.onMessage never receives
// other add-ons' messages, and no onMessageExternal listener exists). Its
// own pages (moz-extension:) may call every handler; the compose script,
// which runs inside compose windows next to mail content, only these.
const COMPOSE_SCRIPT_TYPES = new Set([
  "compose-hello",
  "compose-info",
  "mention-search",
  "mention-add",
  "undo-cancel",
  "undo-send-now",
]);
const OWN_PAGE_PREFIX = messenger.runtime.getURL("/");
// Each page of the add-on may send only the requests it uses.
const PAGE_TYPES = new Map(
  Object.entries({
    "options/options.html": [
      "apply-recommended", "cleanup-open", "collect-contacts-now", "collect-contacts-status",
      "conversations-state", "diagnostics", "get-signature", "get-steps", "import-signature",
      "mentions-scan", "pin-sync-check", "pin-sync-now", "reading-pane-status", "reload-contacts",
      "schedule-cancel", "schedule-list", "schedule-send-now", "settings-sync-now",
      "search-set", "search-status", "settings-sync-status", "settings-sync-undo", "signature-from-mail", "take-settings-section",
    ],
    "popup/quicksteps.html": [
      "cleanup-open", "get-steps", "mentions-show", "open-settings", "run-step", "schedule-list", "toggle-pin",
    ],
    "popup/history.html": ["history", "open-message"],
    "popup/move.html": ["move-data", "move-to"],
    "popup/schedule.html": ["schedule-send"],
    "cleanup/cleanup.html": ["cleanup-accounts", "cleanup-apply", "cleanup-scan", "cleanup-stop"],
  }).map(([path, types]) => [messenger.runtime.getURL(path), new Set(types)])
);

function senderAllowed(type, sender) {
  if (!sender || sender.id !== messenger.runtime.id) {
    return false;
  }
  if (typeof sender.url === "string" && sender.url.startsWith(OWN_PAGE_PREFIX)) {
    const types = PAGE_TYPES.get(sender.url.replace(/[?#].*$/, ""));
    return Boolean(types && types.has(type));
  }
  return COMPOSE_SCRIPT_TYPES.has(type);
}

messenger.runtime.onMessage.addListener((message, sender) => {
  if (!message || typeof message.type !== "string" || !Object.hasOwn(handlers, message.type)) {
    return undefined;
  }
  if (!senderAllowed(message.type, sender)) {
    logError("message", new Error(`refused "${message.type}" from ${(sender && sender.url) || "unknown sender"}`));
    return undefined;
  }
  try {
    return Promise.resolve(handlers[message.type](message, sender)).catch((error) => {
      if (!error || !error.userFacing) {
        logError(message.type, error);
      }
      throw error;
    });
  } catch (error) {
    logError(message.type, error);
    return Promise.reject(error);
  }
});

/* ------------------------------------------------------------------ */
/* Startup                                                             */
/* ------------------------------------------------------------------ */

rebuildMenus().catch((error) => logError("menus", error));
ensureTag(PIN_TAG_KEY, PIN_TAG_NAME, PIN_TAG_COLOR).catch((error) => logError("pin tag", error));
ensureTag(MENTION_TAG_KEY, MENTION_TAG_NAME, MENTION_TAG_COLOR).catch((error) =>
  logError("mention tag", error)
);
restoreScheduleAlarms()
  .then(() => runDueScheduled())
  .catch((error) => logError("schedule restore", error));
