"use strict";

// @Mentioned me: incoming mail whose HTML body contains an Outlook-style
// mention of one of my addresses (<a href="mailto:me@...">@Name</a>) gets the
// "@ Mentioned" tag; the Quick Filter can then show only those messages.

const MENTION_TAG_KEY = "olmention";
const MENTION_TAG_NAME = "@ Mentioned";
const MENTION_TAG_COLOR = "#0078D4";

// Bounded (an anchor's attributes and @Name text are short), so the scan
// stays linear whatever the HTML is (no catastrophic backtracking).
const ANCHOR_RE = /<a\b([^<>]{0,1000})>([\s\S]{0,300}?)<\/a\s{0,10}>/gi;
const HREF_RE = /\bhref\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/i;

function decodeEntities(text) {
  return text
    .replace(/&nbsp;/gi, " ")
    .replace(/&#64;|&commat;/gi, "@")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"');
}

// True when the HTML contains a mention link to one of the given addresses.
function htmlMentions(html, addresses) {
  ANCHOR_RE.lastIndex = 0;
  let match;
  while ((match = ANCHOR_RE.exec(html))) {
    const attrs = match[1];
    const href = HREF_RE.exec(attrs);
    const url = href ? decodeEntities(href[1] || href[2] || href[3] || "") : "";
    if (!/^mailto:/i.test(url)) {
      continue;
    }
    let email = url.slice(7).split("?")[0];
    try {
      email = decodeURIComponent(email);
    } catch (error) {
      // keep as is
    }
    if (!addresses.has(email.trim().toLowerCase())) {
      continue;
    }
    const text = decodeEntities(match[2].replace(/<[^>]*>/g, "")).trim();
    // Outlook marks mentions with id="OWAAM..."; both Outlook and this add-on
    // start the link text with "@".
    if (text.startsWith("@") || /\bid\s*=\s*["']?OWAAM/i.test(attrs)) {
      return true;
    }
  }
  return false;
}

async function messageHtml(messageId) {
  try {
    const parts = await messenger.messages.listInlineTextParts(messageId);
    return parts
      .filter((p) => /html/i.test(p.contentType))
      .map((p) => p.content)
      .join("\n");
  } catch (error) {
    // Older API: walk the full MIME tree.
    const full = await messenger.messages.getFull(messageId);
    const out = [];
    const walk = (part) => {
      if (/^text\/html/i.test(part.contentType || "") && part.body) {
        out.push(part.body);
      }
      (part.parts || []).forEach(walk);
    };
    walk(full);
    return out.join("\n");
  }
}

// Mail in these folders is mine or already handled; never tag it.
const SKIP_FOLDER_USES = ["archives", "drafts", "junk", "outbox", "sent", "templates", "trash"];

function skipFolder(folder) {
  return Boolean(folder && (folder.specialUse || []).some((use) => SKIP_FOLDER_USES.includes(use)));
}

async function readHtmlWithRetry(messageId) {
  try {
    return await messageHtml(messageId);
  } catch (error) {
    // A message that has just arrived may not be fully stored yet.
    await new Promise((resolve) => setTimeout(resolve, 2000));
    return messageHtml(messageId);
  }
}

async function checkMentions(messages, addresses) {
  let tagged = 0;
  for (const message of messages) {
    if ((message.tags || []).includes(MENTION_TAG_KEY) || skipFolder(message.folder)) {
      continue;
    }
    if (addresses.has(emailOf(message.author))) {
      continue; // my own mail
    }
    try {
      if (htmlMentions(await readHtmlWithRetry(message.id), addresses)) {
        await ensureTag(MENTION_TAG_KEY, MENTION_TAG_NAME, MENTION_TAG_COLOR);
        await addTag(message, MENTION_TAG_KEY);
        tagged++;
      }
    } catch (error) {
      logError(`mention check (message ${message.id})`, error);
    }
  }
  return tagged;
}

messenger.messages.onNewMailReceived.addListener(async (folder, messageList) => {
  try {
    const settings = await getSettings();
    if (!settings.mentionTagging || skipFolder(folder)) {
      return;
    }
    const messages = await collectList(messageList, 500);
    await checkMentions(messages, await myAddresses());
  } catch (error) {
    logError("new mail mention check", error);
  }
});

// Tag mail from the last `days` days that is addressed to me.
async function scanMentions(days) {
  const fromDate = new Date(Date.now() - days * 864e5);
  const messages = await collectList(await messenger.messages.query({ fromDate, toMe: true }), 3000);
  const tagged = await checkMentions(messages, await myAddresses());
  return { scanned: messages.length, tagged };
}

async function showMentions(tabId, show) {
  if (show) {
    await ensureTag(MENTION_TAG_KEY, MENTION_TAG_NAME, MENTION_TAG_COLOR);
    await messenger.mailTabs.setQuickFilter(tabId, {
      show: true,
      tags: { mode: "any", tags: { [MENTION_TAG_KEY]: true } },
    });
  } else {
    await messenger.mailTabs.setQuickFilter(tabId, { tags: false });
  }
  return { ok: true };
}

async function mentionTargetTab(tabId) {
  const tab = tabId ? await messenger.tabs.get(tabId) : await activeTab();
  if (tab && tab.type === "mail") {
    return tab;
  }
  const [mailTab] = await messenger.mailTabs.query({ active: true, currentWindow: true });
  return mailTab;
}

handlers["mentions-scan"] = fromOwnPage((message) => scanMentions(Number(message.days) || 30));
handlers["mentions-show"] = fromOwnPage(async (message) => {
  const tab = await mentionTargetTab(message.tabId);
  if (!tab) {
    return { ok: false, error: "Open a mail folder first." };
  }
  return showMentions(tab.id, message.show !== false);
});
