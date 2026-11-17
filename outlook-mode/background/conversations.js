"use strict";

// Thunderbird Conversations (gconversation@xulforum.org) shows a whole thread
// in the message pane; the reading pane features need it. This add-on never
// downloads, installs or enables it, and makes no network request for it:
// installing software is left to the user, through their usual channel for add-ons.
// When a reading pane feature is on and Conversations is missing or turned
// off, the user is told once (per state) what to do; the features then simply
// do nothing. To go without it, turn the reading pane features off in
// Preferences.

const READING_PANE_KEYS = ["inlineReply", "readingPaneNewestFirst", "readingPaneExpandAll", "readingPaneFoldQuotes", "readingPaneLatestOnly"];
const CONVERSATIONS_NOTICE_KEY = "conversationsNoticeFor";

function readingPaneWanted(settings) {
  return READING_PANE_KEYS.some((key) => settings[key]);
}

// What the user should do, or "" when nothing is needed.
function conversationsAdvice(state) {
  if (state.installed && state.enabled) {
    return "";
  }
  if (state.installed) {
    return "Thunderbird Conversations is installed but turned off. Turn it on in Add-ons and Themes, or turn off the reading pane features in Outlook Mode's Preferences.";
  }
  return "The reading pane features need the Thunderbird Conversations add-on. Install it through your usual channel for add-ons, or turn off the reading pane features in Outlook Mode's Preferences.";
}

async function checkConversations() {
  const settings = await getSettings();
  if (!readingPaneWanted(settings)) {
    return;
  }
  const state = await messenger.outlookPane.conversationsState();
  const advice = conversationsAdvice(state);
  if (!advice) {
    return;
  }
  // Once per situation (missing / turned off, per Thunderbird version), not
  // at every startup.
  const situation = `${state.installed ? "off" : "missing"}@${state.appVersion}`;
  const stored = (await messenger.storage.local.get(CONVERSATIONS_NOTICE_KEY))[
    CONVERSATIONS_NOTICE_KEY
  ];
  if (stored !== situation) {
    await messenger.storage.local.set({ [CONVERSATIONS_NOTICE_KEY]: situation });
    notify("Outlook Mode", advice);
  }
}

// After startup, when the main window is up; and when a reading pane
// feature is turned on.
setTimeout(() => checkConversations().catch((error) => logError("check Conversations", error)), 15000);
messenger.storage.onChanged.addListener((changes) => {
  const before = (changes.settings && changes.settings.oldValue) || {};
  const after = (changes.settings && changes.settings.newValue) || {};
  if (READING_PANE_KEYS.some((key) => after[key] && !before[key])) {
    checkConversations().catch((error) => logError("check Conversations", error));
  }
});

handlers["conversations-state"] = fromOwnPage(async () => {
  const state = await messenger.outlookPane.conversationsState();
  return { ...state, advice: conversationsAdvice(state) };
});
