"use strict";

// Quick Steps: mark read -> pin -> reply/forward -> move/archive/delete.

async function runStep(step, messages) {
  if (!step || !messages.length) {
    return { ran: false };
  }
  if (step.after === "move" && !(await isOwnFolder(step.moveTo))) {
    throw userError(`The folder of Quick Step "${step.name}" no longer exists. Pick it again in Preferences.`);
  }
  const ids = messages.map((m) => m.id);
  if (step.markRead) {
    await Promise.all(ids.map((id) => messenger.messages.update(id, { read: true })));
  }
  if (step.pin) {
    await setPinned(messages, true);
  }
  const first = messages[0];
  if (step.compose === "reply" || step.compose === "replyAll") {
    await messenger.compose.beginReply(
      first.id,
      step.compose === "replyAll" ? "replyToAll" : "replyToSender"
    );
  } else if (step.compose === "forward") {
    // Opens the compose window for the user to check and send; never sends.
    // Only people in the user's address books are filled in as recipients.
    const known = new Set((await getContacts().catch(() => [])).map((c) => c.email.toLowerCase()));
    const to = parseAddressList(step.forwardTo).filter((address) => known.has(emailOf(address).toLowerCase()));
    await messenger.compose.beginForward(first.id, "forwardInline", to.length ? { to } : {});
  }
  if (step.after === "move" && step.moveTo) {
    // Only to a folder of the account the mail is in (never to another account).
    const target = await messenger.folders.get(step.moveTo).catch(() => null);
    if (!target || messages.some((m) => !m.folder || m.folder.accountId !== target.accountId)) {
      throw userError(`Quick Step "${step.name}" moves mail only within the same account. Pick a folder of this account in Preferences.`);
    }
    await messenger.messages.move(ids, step.moveTo);
  } else if (step.after === "archive") {
    await messenger.messages.archive(ids);
  } else if (step.after === "delete") {
    await messenger.messages.delete(ids, false);
  }
  return { ran: true, count: ids.length };
}

// Runs only on the user's own action in the mail tab they are looking at
// (menu, shortcut or popup on that tab), on the mail selected there.
async function runStepOnMessages(index, messages, tab) {
  if (!Number.isInteger(index) || !isMailTab(tab) || !tab.active) {
    return { ran: false };
  }
  const steps = await getSteps();
  if (index < 0 || index >= steps.length) {
    return { ran: false };
  }
  const result = await runStep(steps[index], messages);
  if (steps[index] && steps[index].pin) {
    await resortIfInbox(tab).catch(() => {});
  }
  return result;
}

async function runStepOnTab(index, tab) {
  return runStepOnMessages(index, await selectedMessages(tab), tab);
}

handlers["get-steps"] = fromOwnPage(() => getSteps());

// Only from the Quick Steps popup, on the mail tab the user is looking at.
handlers["run-step"] = fromPages(["popup/quicksteps.html"], async (message) => {
  const tab = Number.isInteger(message.tabId) ? await messenger.tabs.get(message.tabId).catch(() => null) : null;
  if (!tab || !tab.active || tab.type !== "mail") {
    throw userError("Open the message list first.");
  }
  return runStepOnTab(message.index, tab);
});
