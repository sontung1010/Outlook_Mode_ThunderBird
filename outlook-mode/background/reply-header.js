"use strict";

// Data for the Outlook-style "From / Sent / To / Cc / Subject" block that the
// compose script puts on replies and forwards.

async function composeInfo(tabId) {
  const settings = await getSettings();
  const details = await messenger.compose.getComposeDetails(tabId);
  if (!["reply", "forward"].includes(details.type) || !details.relatedMessageId) {
    return { type: details.type, settings };
  }
  const original = await messenger.messages.get(details.relatedMessageId);
  return {
    type: details.type,
    settings,
    header: {
      from: original.author || "",
      sent: formatOutlookDate(original.date),
      to: (original.recipients || []).join("; "),
      cc: (original.ccList || []).join("; "),
      subject: original.subject || "",
    },
  };
}

handlers["compose-info"] = (message, sender) =>
  sender.tab ? composeInfo(sender.tab.id) : undefined;
