"use strict";

// Recent mail with the person in the displayed message.

const HISTORY_DAYS = 365;
const HISTORY_LIMIT = 30;

async function history(tabId) {
  // Only the tab the user is looking at (the popup asks for the active tab
  // of its window); any other tab id is refused.
  const tab = Number.isInteger(tabId) ? await messenger.tabs.get(tabId).catch(() => null) : null;
  if (!tab || !tab.active) {
    return { error: "No message is open." };
  }
  const [shown] = await messenger.messageDisplay.getDisplayedMessages(tabId).then(collectList);
  if (!shown) {
    return { error: "No message is open." };
  }
  const mine = await myAddresses();
  let person = shown.author;
  if (mine.has(emailOf(person)) && (shown.recipients || []).length) {
    person = shown.recipients[0];
  }
  const email = emailOf(person);
  // messages.query matches this text against the local headers (no query
  // language); only a plain address is looked up all the same.
  if (!EMAIL_PATTERN.test(email)) {
    return { error: "This message has no usable address." };
  }
  // Searched only in the account of the displayed mail.
  const accountId = shown.folder && shown.folder.accountId;
  if (!accountId) {
    return { error: "No message is open." };
  }
  const fromDate = new Date(Date.now() - HISTORY_DAYS * 864e5);
  const [sent, received] = await Promise.all([
    messenger.messages.query({ accountId, author: email, fromDate }).then((l) => collectList(l, 500)),
    messenger.messages.query({ accountId, recipients: email, fromDate }).then((l) => collectList(l, 500)),
  ]);
  const byId = new Map();
  for (const m of [...sent, ...received]) {
    byId.set(m.id, m);
  }
  const messages = [...byId.values()]
    .sort((a, b) => new Date(b.date) - new Date(a.date))
    .slice(0, HISTORY_LIMIT)
    .map((m) => ({
      id: m.id,
      date: new Date(m.date).toISOString(),
      subject: m.subject || "(no subject)",
      fromMe: mine.has(emailOf(m.author)),
      folder: (m.folder && (m.folder.name || m.folder.path)) || "",
      current: m.id === shown.id,
    }));
  return { person: { name: nameOf(person), email }, days: HISTORY_DAYS, messages };
}

// Only from the History popup, for the tab the user is looking at (found
// here; the tab id in the request is not used).
handlers.history = fromPages(["popup/history.html"], async () => {
  const [tab] = await messenger.tabs.query({ active: true, lastFocusedWindow: true });
  return history(tab ? tab.id : null);
});

handlers["open-message"] = fromOwnPage((message) =>
  messenger.messageDisplay.open({ messageId: message.messageId, location: "tab" }));
