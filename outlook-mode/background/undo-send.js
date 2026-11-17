"use strict";

// Undo Send: when the user presses Send, the send is cancelled and a countdown
// banner ("Sending in 10s... Undo / Send now") is shown in the compose window.
// When the countdown ends the message is sent with compose.sendMessage(),
// which keeps everything the compose window holds (threading, signature,
// attachments, encryption settings).

const pendingSends = new Map(); // tabId -> { timer, endsAt }
const bypassSend = new Set(); // tabIds whose next onBeforeSend must pass through

function clearPending(tabId) {
  const pending = pendingSends.get(tabId);
  if (pending) {
    clearTimeout(pending.timer);
    pendingSends.delete(tabId);
  }
  return Boolean(pending);
}

async function tellCompose(tabId, message) {
  try {
    return await messenger.tabs.sendMessage(tabId, message);
  } catch (error) {
    return undefined; // compose script not loaded in this window
  }
}

// Send a compose tab for real, skipping the Undo Send countdown.
async function sendNowBypassingUndo(tabId, mode = "default") {
  clearPending(tabId);
  await tellCompose(tabId, { type: "undo-hide" });
  bypassSend.add(tabId);
  try {
    return await messenger.compose.sendMessage(tabId, { mode });
  } finally {
    bypassSend.delete(tabId);
  }
}

async function startCountdown(tabId, seconds) {
  // Only cancel the send when the banner is really visible, so a message can
  // never get stuck in a window without a way to undo or send it.
  const shown = await tellCompose(tabId, { type: "undo-show", seconds });
  if (!shown || !shown.ok) {
    return false;
  }
  const timer = setTimeout(() => {
    sendNowBypassingUndo(tabId).catch((error) => logError("undo send: delayed send", error));
  }, seconds * 1000);
  pendingSends.set(tabId, { timer, endsAt: Date.now() + seconds * 1000 });
  return true;
}

messenger.compose.onBeforeSend.addListener(async (tab) => {
  // Never let the mention popup or our banner end up in a sent message. The
  // reply also tells when Ctrl+Shift+Enter was last pressed in the editor.
  const cleanup = await tellCompose(tab.id, { type: "ui-cleanup" });

  if (bypassSend.has(tab.id)) {
    return {};
  }
  // Pressing Send again during the countdown sends right away.
  if (clearPending(tab.id)) {
    return {};
  }
  // Ctrl+Shift+Enter (Send Later) puts the message in the Outbox, where it can
  // still be edited, so no countdown is needed.
  if (cleanup && Date.now() - (cleanup.sendLaterKeyAt || 0) < 3000) {
    return {};
  }
  try {
    const { undoSendSeconds } = await getSettings();
    const seconds = Number(undoSendSeconds) || 0;
    if (seconds > 0 && (await startCountdown(tab.id, seconds))) {
      return { cancel: true };
    }
  } catch (error) {
    logError("undo send", error);
  }
  return {};
});

messenger.tabs.onRemoved.addListener((tabId) => {
  clearPending(tabId);
});

handlers["undo-cancel"] = (message, sender) => {
  if (sender.tab) {
    clearPending(sender.tab.id);
  }
  return { ok: true };
};

// "Send now" on the banner finishes a send the user already started: only a
// compose tab with an Undo Send countdown running, and only the tab the
// banner is in (sender.tab is set by Thunderbird, not by the page).
handlers["undo-send-now"] = (message, sender) => {
  const tabId = sender.tab && sender.tab.id;
  if (sender.tab?.type !== "messageCompose" || !pendingSends.has(tabId)) {
    return { ok: false };
  }
  sendNowBypassingUndo(tabId).catch((error) => logError("undo send: send now", error));
  return { ok: true };
};
