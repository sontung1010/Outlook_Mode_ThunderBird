"use strict";

// Outlook-style keyboard shortcuts (commands API) and the folder picker used
// by "Move to folder" (Ctrl+Shift+V).
//
// Thunderbird's own shortcuts always win over add-on shortcuts, so only key
// combinations Thunderbird leaves free can be used (see options page).

const RECENT_FOLDERS_KEY = "recentFolders";
const MAX_RECENT_FOLDERS = 6;

async function markRead(messages, read) {
  await Promise.all(messages.map((m) => messenger.messages.update(m.id, { read })));
  return { count: messages.length };
}

async function forwardAsAttachment(messages) {
  if (!messages.length) {
    return { count: 0 };
  }
  const [first, ...rest] = messages;
  const tab = await messenger.compose.beginForward(first.id, "forwardAsAttachment");
  // beginForward takes one message; attach the others as .eml files.
  for (const message of rest) {
    const file = await messenger.messages.getRaw(message.id, { data_format: "File" });
    const name = `${(message.subject || "message").replace(/[\\/:*?"<>|]+/g, "_").slice(0, 80)}.eml`;
    await messenger.compose.addAttachment(tab.id, {
      file: new File([file], name, { type: "message/rfc822" }),
      name,
    });
  }
  return { count: messages.length };
}

/* ------------------------------------------------------------------ */
/* Move to folder                                                      */
/* ------------------------------------------------------------------ */

let moveSource = null; // { tabId, messageIds } captured when the picker opens

async function listAllFolders() {
  const accounts = await messenger.accounts.list(true);
  const out = [];
  const walk = (folders, prefix, accountName) => {
    for (const folder of folders || []) {
      if (!folder.id) {
        continue;
      }
      const path = `${prefix}${folder.name}`;
      out.push({ id: folder.id, name: folder.name, path, account: accountName });
      walk(folder.subFolders, `${path} / `, accountName);
    }
  };
  for (const account of accounts) {
    const top = (account.rootFolder && account.rootFolder.subFolders) || account.folders || [];
    walk(top, "", account.name);
  }
  return out;
}

// A folder (id) of one of the user's accounts, as the folder picker lists them.
async function isOwnFolder(folderId) {
  return typeof folderId === "string" && (await listAllFolders()).some((folder) => folder.id === folderId);
}

async function openMovePicker(tab) {
  const messages = await selectedMessages(tab);
  if (!messages.length) {
    return { ok: false };
  }
  moveSource = { tabId: tab.id, messageIds: messages.map((m) => m.id) };
  await messenger.windows.create({
    url: "/popup/move.html",
    type: "popup",
    width: 460,
    height: 540,
  });
  return { ok: true };
}

async function moveFolderData() {
  const { [RECENT_FOLDERS_KEY]: recent = [] } = await messenger.storage.local.get(
    RECENT_FOLDERS_KEY
  );
  return {
    count: moveSource ? moveSource.messageIds.length : 0,
    folders: await listAllFolders(),
    recent,
  };
}

async function moveSelectedTo(folderId) {
  if (!moveSource || !moveSource.messageIds.length) {
    throw userError("No messages selected.");
  }
  if (!(await isOwnFolder(folderId))) {
    throw userError("That folder does not exist any more.");
  }
  const { messageIds } = moveSource;
  moveSource = null;
  await messenger.messages.move(messageIds, folderId);
  const { [RECENT_FOLDERS_KEY]: recent = [] } = await messenger.storage.local.get(
    RECENT_FOLDERS_KEY
  );
  const updated = [folderId, ...recent.filter((id) => id !== folderId)].slice(
    0,
    MAX_RECENT_FOLDERS
  );
  await messenger.storage.local.set({ [RECENT_FOLDERS_KEY]: updated });
  return { ok: true, count: messageIds.length };
}

handlers["move-data"] = fromOwnPage(() => moveFolderData());
handlers["move-to"] = fromOwnPage((message) => moveSelectedTo(message.folderId));

/* ------------------------------------------------------------------ */
/* Commands                                                            */
/* ------------------------------------------------------------------ */

messenger.commands.onCommand.addListener(async (command, tab) => {
  try {
    const target = tab || (await activeTab());
    // Some combinations (e.g. Ctrl+D) also reach the compose window; mail
    // actions only make sense in a mail or message tab.
    if (!isMailTab(target)) {
      return;
    }
    if (command === "toggle-pin") {
      await pinSelected(target);
    } else if (command.startsWith("quickstep-")) {
      await runStepOnTab(Number(command.slice("quickstep-".length)) - 1, target);
    } else if (command === "ol-delete") {
      const ids = (await selectedMessages(target)).map((m) => m.id);
      if (ids.length) {
        await messenger.messages.delete(ids, false);
      }
    } else if (command === "ol-move") {
      await openMovePicker(target);
    } else if (command === "ol-forward-attachment") {
      await forwardAsAttachment(await selectedMessages(target));
    } else if (command === "ol-mark-read") {
      await markRead(await selectedMessages(target), true);
    } else if (command === "ol-mark-unread") {
      await markRead(await selectedMessages(target), false);
    } else if (command === "ol-show-mentions") {
      if (target.type === "mail") {
        await showMentions(target.id, true);
      }
    }
  } catch (error) {
    logError(`shortcut ${command}`, error);
  }
});

handlers["get-commands"] = fromOwnPage(() => messenger.commands.getAll());
