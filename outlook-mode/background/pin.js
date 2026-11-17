"use strict";

// Pin: the 📌 tag. Like Outlook, pinning works on whole conversations: the
// root of the thread is tagged too. The "Pinned first" sort (experiment
// outlookPin) puts conversations with a pinned message on top, newest first,
// and keeps the view threaded. The star flag is not used: Microsoft Graph
// accounts do not keep it, and in Outlook a flag means "follow up", not "pin".

const PIN_TAG_KEY = "olpinned";
const PIN_TAG_NAME = "📌 Pinned";
const PIN_TAG_COLOR = "#C0392B";
const MAX_REFERENCES_CHECKED = 15;
// Inboxes whose threads were turned off by v3.0.1 and have been turned back on.
const THREADS_RESTORED_KEY = "threadsRestoredFolders";

function isPinned(message) {
  return (message.tags || []).includes(PIN_TAG_KEY);
}

function messageIdsIn(values) {
  const text = (Array.isArray(values) ? values : [values || ""]).join(" ");
  return (text.match(/<[^<>\s]+>/g) || []).map((id) => id.slice(1, -1));
}

async function readHeaders(messageId) {
  if (messenger.messages.getHeaders) {
    return messenger.messages.getHeaders(messageId);
  }
  return (await messenger.messages.getFull(messageId)).headers || {};
}

// The thread root as Thunderbird shows it: the oldest message in the
// References chain that is in the same folder. null if the message is a root.
async function threadRootOf(message) {
  if (!message.folder || !message.folder.id) {
    return null;
  }
  let headers;
  try {
    headers = await readHeaders(message.id);
  } catch (error) {
    return null; // headers not available (e.g. offline); pin the message only
  }
  // The ids come from the mail's headers: only well-formed Message-IDs are
  // looked up (messages.query compares the whole id, in this folder only).
  const refs = [...messageIdsIn(headers.references), ...messageIdsIn(headers["in-reply-to"])]
    .filter((id, index, all) => PLAIN_MESSAGE_ID.test(id) && all.indexOf(id) === index && id !== message.headerMessageId)
    .slice(0, MAX_REFERENCES_CHECKED);
  for (const ref of refs) {
    const found = await collectList(
      await messenger.messages.query({ folderId: message.folder.id, headerMessageId: ref }),
      5
    );
    if (found.length) {
      return found[0];
    }
  }
  return null;
}

// The selected messages plus the roots of their threads, without duplicates.
async function withThreadRoots(messages) {
  const byId = new Map(messages.map((m) => [m.id, m]));
  for (const message of messages) {
    const root = await threadRootOf(message);
    if (root && !byId.has(root.id)) {
      byId.set(root.id, root);
    }
  }
  return [...byId.values()];
}

// The given messages and every message of their conversations (thread, in
// their folder). As in Outlook, a conversation is pinned and unpinned as a
// whole: then the reading pane, which shows its newest mail, shows the pin
// too, and pin sync from the server cannot leave a copy pinned.
async function conversationsOf(messages) {
  const byId = new Map(messages.map((m) => [m.id, m]));
  for (const message of messages) {
    let ids = [];
    try {
      ids = await messenger.outlookPin.conversationMessageIds(message.id);
    } catch (error) {
      logError("conversation of pinned mail", error);
    }
    for (const id of ids) {
      if (!byId.has(id)) {
        const other = await messenger.messages.get(id).catch(() => null);
        if (other) {
          byId.set(id, other);
        }
      }
    }
  }
  return [...byId.values()];
}

// The given messages and the pinned ones of their conversations.
async function pinnedInConversations(messages) {
  const given = new Set(messages.map((m) => m.id));
  return (await conversationsOf(messages)).filter((m) => given.has(m.id) || isPinned(m));
}

// Mail read again from Thunderbird (not the objects passed in), kept only
// when it is in one of the user's own accounts in this profile.
async function ownMessages(list) {
  const accounts = new Set((await messenger.accounts.list(false)).map((a) => a.id));
  const fresh = await Promise.all(list.map((m) => messenger.messages.get(m.id).catch(() => null)));
  return fresh.filter((m) => m && m.folder && accounts.has(m.folder.accountId));
}

async function setPinned(messages, pinned) {
  await ensureTag(PIN_TAG_KEY, PIN_TAG_NAME, PIN_TAG_COLOR);
  const all = await conversationsOf(await withThreadRoots(messages));
  // Only what changes (and what was asked for), so pin sync uploads no more.
  const given = new Set(messages.map((m) => m.id));
  const targets = await ownMessages(all.filter((m) => given.has(m.id) || isPinned(m) != pinned));
  await Promise.all(
    targets.map((m) => {
      const tags = new Set(m.tags || []);
      if (pinned) {
        tags.add(PIN_TAG_KEY);
      } else {
        tags.delete(PIN_TAG_KEY);
      }
      return messenger.messages.update(m.id, { tags: [...tags] });
    })
  );
  queuePinSync(targets, pinned); // Microsoft Graph accounts: also on the server
}

// New mail in a pinned conversation is pinned too, as the rest of it.
messenger.messages.onNewMailReceived.addListener(async (folder, list) => {
  try {
    const incoming = list.messages || [];
    const toPin = [];
    for (const message of incoming) {
      if (isPinned(message)) {
        continue;
      }
      const others = (await conversationsOf([message])).filter((m) => m.id != message.id);
      if (others.some(isPinned)) {
        toPin.push(message);
      }
    }
    if (toPin.length) {
      await ensureTag(PIN_TAG_KEY, PIN_TAG_NAME, PIN_TAG_COLOR);
      for (const message of toPin) {
        await messenger.messages.update(message.id, { tags: [...new Set([...(message.tags || []), PIN_TAG_KEY])] });
      }
      queuePinSync(toPin, true);
    }
  } catch (error) {
    logError("pin new mail of a pinned conversation", error);
  }
});

async function togglePin(messages) {
  if (!messages.length) {
    return { pinned: null };
  }
  // A conversation is pinned while any of its mail is. Mixed selection: pin
  // all. All pinned: unpin all.
  const pinnedNow = new Set((await pinnedInConversations(messages)).filter(isPinned).map((m) => m.id));
  let pin = false;
  for (const message of messages) {
    let ids = [message.id];
    try {
      ids = await messenger.outlookPin.conversationMessageIds(message.id);
    } catch (error) {
      // the message alone
    }
    if (!ids.some((id) => pinnedNow.has(id))) {
      pin = true;
      break;
    }
  }
  await setPinned(messages, pin);
  return { pinned: pin };
}

// v3.0.1 switched pinned Inboxes to a flat list. After an update from 3.0.1,
// turn threads back on once per Inbox; after that the user's own view choice
// is kept. Fresh installs and other updates never touch the view type.
const RESTORE_THREADS_PENDING_KEY = "restoreThreadsPending";

messenger.runtime.onInstalled.addListener(({ reason, previousVersion }) => {
  if (reason === "update" && previousVersion === "3.0.1") {
    messenger.storage.local
      .set({ [RESTORE_THREADS_PENDING_KEY]: true })
      .catch((error) => logError("restore threads flag", error));
  }
});

async function restoreThreadsOnce(tabId, folder) {
  if (!folder || !folder.id) {
    return;
  }
  const stored = await messenger.storage.local.get([
    THREADS_RESTORED_KEY,
    RESTORE_THREADS_PENDING_KEY,
  ]);
  if (!stored[RESTORE_THREADS_PENDING_KEY]) {
    return;
  }
  const done = stored[THREADS_RESTORED_KEY] || [];
  if (done.includes(folder.id)) {
    return;
  }
  await messenger.mailTabs.update(tabId, { viewType: "groupedByThread" });
  await messenger.storage.local.set({ [THREADS_RESTORED_KEY]: [...done, folder.id] });
}

// Register the "Pinned" column and the "Pinned first" Sort by entries.
const pinSortReady = messenger.outlookPin
  ? messenger.outlookPin
      .register({ pinKeyword: PIN_TAG_KEY, columnName: "📌 Pinned", menuLabel: "📌 Pinned first" })
      .then(() => true)
      .catch((error) => {
        logError("register pinned sort", error);
        return false;
      })
  : Promise.resolve(false);

// Pinned conversations on top, then the rest; both newest first.
async function applyPinnedSort(tabId, folder) {
  await restoreThreadsOnce(tabId, folder).catch((error) => logError("restore threads", error));
  if ((await pinSortReady) && (await messenger.outlookPin.sortByPinned(tabId))) {
    return;
  }
  // Fallback without the experiment: tagged mail first.
  await messenger.mailTabs.update(tabId, { sortType: "date", sortOrder: "descending" });
  await messenger.mailTabs.update(tabId, { sortType: "tags", sortOrder: "descending" });
}

async function isPinnedSortActive(tabId) {
  return (await pinSortReady) && messenger.outlookPin.isSortedByPinned(tabId);
}

// After pinning, re-sort if the list is sorted by "Pinned first" (chosen in
// the Sort by menu) or is an Inbox with "Keep pinned mail on top".
async function resortAfterPin(tab) {
  if (!tab || tab.type !== "mail") {
    return;
  }
  const settings = await getSettings();
  const mailTab = await messenger.mailTabs.get(tab.id);
  if (
    (settings.pinnedOnTop && isInbox(mailTab.displayedFolder)) ||
    (await isPinnedSortActive(tab.id))
  ) {
    await applyPinnedSort(tab.id, mailTab.displayedFolder);
  }
}

async function resortIfInbox(tab) {
  return resortAfterPin(tab);
}

messenger.mailTabs.onDisplayedFolderChanged.addListener(async (tab, folder) => {
  try {
    const settings = await getSettings();
    if (settings.pinnedOnTop && isInbox(folder)) {
      await applyPinnedSort(tab.id, folder);
    }
  } catch (error) {
    logError("pinned sort", error);
  }
});

// Inboxes that are already open when the add-on starts.
(async () => {
  const settings = await getSettings();
  if (!settings.pinnedOnTop) {
    return;
  }
  for (const mailTab of await messenger.mailTabs.query({})) {
    if (isInbox(mailTab.displayedFolder)) {
      await applyPinnedSort(mailTab.id, mailTab.displayedFolder);
    }
  }
})().catch((error) => logError("startup pinned sort", error));

async function pinSelected(tab) {
  const result = await togglePin(await selectedMessages(tab));
  await resortIfInbox(tab).catch(() => {});
  return result;
}

handlers["toggle-pin"] = fromOwnPage((message) =>
  messenger.tabs.get(message.tabId).then((tab) => pinSelected(tab)));
