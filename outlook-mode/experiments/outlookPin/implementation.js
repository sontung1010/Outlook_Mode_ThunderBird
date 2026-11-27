"use strict";

/* global ExtensionAPI, ExtensionCommon, ChromeUtils, Services, Ci */

// Experiment: a sortable "Pinned" column for the message list, plus a
// "Pinned first" entry in View > Sort by and in the message list header's
// Sort by menu.
//
// The sort key puts conversations that contain a pinned message first, then
// everything else; both groups newest first (a thread counts with its newest
// message, like Outlook). It reads the pin tag directly and does not depend on
// the star flag, which some account types (Microsoft Graph) do not keep.
//
// It also hands the OAuth access token of a Microsoft Graph account to the
// background script, which uses it to sync pins with the server (pin-sync.js).
//
// In the message list it replaces the star of the cards view with a pin
// button, and adds a refresh button to the list header (get new mail, then
// the background syncs settings, pins and contacts).

var { ExtensionSupport } = ChromeUtils.importESModule(
  "resource:///modules/ExtensionSupport.sys.mjs"
);
var { ThreadPaneColumns } = ChromeUtils.importESModule(
  "chrome://messenger/content/ThreadPaneColumns.mjs"
);
var { MailServices } = ChromeUtils.importESModule("resource:///modules/MailServices.sys.mjs");
var { OAuth2Module } = ChromeUtils.importESModule("resource:///modules/OAuth2Module.sys.mjs");
var { ExtensionUtils } = ChromeUtils.importESModule(
  "resource://gre/modules/ExtensionUtils.sys.mjs"
);
var { ExtensionError } = ExtensionUtils;

const COLUMN_ID = "outlookModePinnedCol";
const MENU_ITEM_ID = "outlookModeSortByPinned";
const SORT_POPUPS = ["menu_viewSortPopup", "menu_threadPaneSortPopup"];
// Sort keys are unsigned 32-bit numbers, sorted ascending.
const PINNED_BASE = 0x7fffffff;
const OTHER_BASE = 0xffffffff;

let pinKeyword = "olpinned";
let menuLabel = "Pinned first";
let registered = false;
let listenerName = "";
let menuIdPrefix = ""; // id prefix of this add-on's context menu items
const windowsWithListener = new Set();
const OWN_SEPARATOR_CLASS = "outlook-mode-top-separator";

const LIST_STYLE_ID = "outlookModeListStyle";
const REFRESH_BUTTON_ID = "outlookModeRefresh";
const PIN_ATTRIBUTE = "data-outlook-mode-pin";
const UNREAD_ATTRIBUTE = "data-outlook-mode-unread";
const CONVERSATION_ATTRIBUTE = "data-outlook-mode-conversation";
const SECTIONS_ATTRIBUTE = "data-outlook-mode-sections";
const SECTION_HEADER = 30; // px above the pinned cards ("📌 Pinned")
const SECTION_GAP = 38; // px above each next section ("Today", "Older"…)
const listState = {
  extension: null,
  pinButton: false,
  refreshButton: false,
  refreshKey: false, // F5 runs the full refresh (as the ⟳ button)
  unreadHighlight: false,
  oneCardPerConversation: false,
  pinnedSection: false,
  dateSections: false,
  newestInCard: false,
  refresh: { busy: false, title: "", error: false },
  panes: new Map(), // about:3pane window -> cleanup function
  redraws: new Set(), // about:3pane windows with a redraw coming
  repairs: new Set(), // folders with a thread count check coming
  threadCountsRepaired: 0, // for Diagnostics
  sectionLabels: new Map(), // about:3pane window -> Map(section name -> heading element)
  sections: new Map(), // about:3pane window -> sections of its list (listSections)
  observing: false,
  waiting: new Set(), // about:3pane windows waiting for their load event
  listeners: { pin: new Set(), refresh: new Set() },
};

function hasPin(hdr) {
  const keywords = hdr.getStringProperty("keywords") || "";
  return keywords.split(" ").includes(pinKeyword);
}

// A conversation is pinned when any of its mail is.
function conversationPinned(hdr) {
  if (hasPin(hdr)) {
    return true;
  }
  try {
    const thread = hdr.folder.msgDatabase.getThreadContainingMsgHdr(hdr);
    for (let i = 0; thread && i < thread.numChildren; i++) {
      const child = thread.getChildHdrAt(i);
      if (child && hasPin(child)) {
        return true;
      }
    }
  } catch (error) {
    // no thread information
  }
  return false;
}

/* ------------------------------------------------------------------ */
/* Your replies count in the conversation, as in Outlook               */
/* ------------------------------------------------------------------ */

// A reply you send goes to Sent Items, so in the Inbox its conversation kept
// the place and the card of its last received mail. Outlook counts your
// replies (from any folder) in the conversation: it moves up and its card
// shows your reply. Here, for each account, the mail in its Sent folders is
// indexed by the Message-IDs it refers to (References: the mail it answers
// and that one's own chain); a conversation's newest reply of yours is the
// newest sent mail referring to one of its messages. The index is made again
// when a Sent folder changes (and its message count is looked at again at
// most every SENT_INDEX_CHECK_MS).
const SENT_INDEX_CHECK_MS = 2000;
const sentIndexes = new Map(); // server key -> { stamp, checkedAt, replies: Map(id -> { date, author }) }

function isSentFolder(folder) {
  try {
    return folder.isSpecialFolder(Ci.nsMsgFolderFlags.SentMail, true);
  } catch (error) {
    return false;
  }
}

function sentFoldersOf(server) {
  try {
    return [...server.rootFolder.getFoldersWithFlags(Ci.nsMsgFolderFlags.SentMail)];
  } catch (error) {
    return [];
  }
}

function sentIndex(server) {
  if (!server) {
    return null;
  }
  const now = Date.now();
  let entry = sentIndexes.get(server.key);
  if (entry && now - entry.checkedAt < SENT_INDEX_CHECK_MS) {
    return entry.replies;
  }
  const folders = sentFoldersOf(server);
  const stamp = folders.map((folder) => `${folder.URI}:${folder.getTotalMessages(false)}`).join("|");
  if (entry && entry.stamp == stamp) {
    entry.checkedAt = now;
    return entry.replies;
  }
  const replies = new Map();
  for (const folder of folders) {
    try {
      for (const hdr of folder.msgDatabase.enumerateMessages()) {
        for (let i = 0; i < hdr.numReferences; i++) {
          const id = hdr.getStringReference(i);
          const known = replies.get(id);
          if (id && (!known || hdr.dateInSeconds > known.date)) {
            replies.set(id, { date: hdr.dateInSeconds, author: hdr.mime2DecodedAuthor || "" });
          }
        }
      }
    } catch (error) {
      console.error("[outlook-mode] sent mail index", error);
    }
  }
  entry = { stamp, checkedAt: now, replies };
  sentIndexes.set(server.key, entry);
  return replies;
}

function forgetSentIndex(folder) {
  if (folder && folder.server && isSentFolder(folder)) {
    sentIndexes.delete(folder.server.key);
    return true;
  }
  return false;
}

// Your newest reply in the conversation of these messages (all in one
// folder, not a Sent folder): { date (s), author } or null.
function newestOwnReply(hdrs) {
  const first = hdrs[0];
  if (!first || !first.folder || isSentFolder(first.folder)) {
    return null;
  }
  const replies = sentIndex(first.folder.server);
  let newest = null;
  for (const hdr of hdrs) {
    const reply = replies && hdr.messageId ? replies.get(hdr.messageId) : null;
    if (reply && (!newest || reply.date > newest.date)) {
      newest = reply;
    }
  }
  return newest;
}

function threadHdrs(hdr) {
  const hdrs = [hdr];
  try {
    const thread = hdr.folder.msgDatabase.getThreadContainingMsgHdr(hdr);
    for (let i = 0; thread && i < thread.numChildren; i++) {
      const child = thread.getChildHdrAt(i);
      if (child && child.messageKey != hdr.messageKey) {
        hdrs.push(child);
      }
    }
  } catch (error) {
    // the message alone
  }
  return hdrs;
}

// The date a message is sorted by, in seconds: a thread is placed by its
// root row, which gets the date of the newest message so active
// conversations come first (your replies in Sent Items included).
// The newest date is read from the thread's mail itself: the thread's own
// newestMsgDate is not always updated (Microsoft 365 accounts keep some
// thread counts stale, see repairThreadCounts).
function sortDate(hdr) {
  let date = hdr.dateInSeconds;
  try {
    const thread = hdr.folder.msgDatabase.getThreadContainingMsgHdr(hdr);
    const root = thread?.getRootHdr();
    if (root && root.messageKey == hdr.messageKey) {
      date = Math.max(date, thread.newestMsgDate);
      const hdrs = [hdr];
      for (let i = 0; i < thread.numChildren; i++) {
        const child = thread.getChildHdrAt(i);
        if (child) {
          date = Math.max(date, child.dateInSeconds);
          hdrs.push(child);
        }
      }
      const reply = newestOwnReply(hdrs);
      if (reply) {
        date = Math.max(date, reply.date);
      }
    }
  } catch (error) {
    // No thread information (e.g. a synthetic view): use the message alone.
  }
  return Math.min(Math.max(date, 0), PINNED_BASE - 1);
}

function sortKey(hdr) {
  return (conversationPinned(hdr) ? PINNED_BASE : OTHER_BASE) - sortDate(hdr);
}

function isSorted(about3Pane) {
  const view = about3Pane?.gViewWrapper;
  return Boolean(
    view &&
      view.primarySortType == Ci.nsMsgViewSortType.byCustom &&
      about3Pane.gDBView?.curCustomColumn == COLUMN_ID
  );
}

// Thunderbird does not sort a list again when new mail joins a conversation
// already in it, so with "Pinned first" the conversation stayed where it was
// (by its older mail). Lists sorted that way are sorted again shortly after
// mail arrives in the folder they show (or in any folder, for a unified or
// saved-search view); several new mails in a row give one sort.
const RESORT_DELAY_MS = 2000;
const resortTimers = new Map(); // about:3pane window -> timer

function pinnedSortedPanes() {
  const panes = [];
  for (const win of Services.wm.getEnumerator("mail:3pane")) {
    for (const tab of win.gTabmail?.tabInfo || []) {
      const about3Pane = tab.mode?.name == "mail3PaneTab" ? tab.chromeBrowser?.contentWindow : null;
      if (about3Pane && !about3Pane.closed && isSorted(about3Pane)) {
        panes.push(about3Pane);
      }
    }
  }
  return panes;
}

function resortSoon(about3Pane) {
  if (resortTimers.has(about3Pane)) {
    return;
  }
  resortTimers.set(
    about3Pane,
    about3Pane.setTimeout(() => {
      resortTimers.delete(about3Pane);
      try {
        if (!about3Pane.closed && isSorted(about3Pane)) {
          sortPinned(about3Pane);
        }
      } catch (error) {
        console.error("[outlook-mode] resort", error);
      }
    }, RESORT_DELAY_MS)
  );
}

function resortShowing(changedFolder) {
  // Mail in a Sent folder moves the conversations it answers, in the other
  // folders of its account.
  const sent = forgetSentIndex(changedFolder);
  for (const about3Pane of pinnedSortedPanes()) {
    const folder = about3Pane.gFolder;
    if (
      about3Pane.gViewWrapper?.isVirtual ||
      (folder && changedFolder && folder.URI == changedFolder.URI) ||
      (sent && folder && folder.server && folder.server.key == changedFolder.server.key)
    ) {
      resortSoon(about3Pane);
    }
  }
}

const resortListener = {
  msgAdded(hdr) {
    resortShowing(hdr.folder);
  },
};
// Also when a folder's message count changes (in case an account type adds
// mail without the msgAdded notification).
const resortFolderListener = {
  onFolderIntPropertyChanged(folder, property, oldValue, newValue) {
    if (property == "TotalMessages" && newValue > oldValue) {
      resortShowing(folder);
    }
  },
};
let resortListening = false;

function sortPinned(about3Pane) {
  if (!about3Pane?.sortController || !about3Pane.gViewWrapper?.dbView) {
    return false;
  }
  if (!about3Pane.sortController.sortThreadPane(COLUMN_ID)) {
    return false;
  }
  about3Pane.threadPane?.restoreSortIndicator();
  return true;
}

function about3PaneForMenu(popup) {
  if (popup.id == "menu_threadPaneSortPopup") {
    return popup.ownerDocument.defaultView; // the popup lives in about:3pane itself
  }
  return popup.ownerDocument.defaultView.gTabmail?.currentAbout3Pane;
}

// Move this add-on's message list context menu entry ("Outlook Mode") from
// the bottom of the menu to the top, right above the first text entry
// ("Open") and below the icon row (Reply / Archive / Junk / Delete).
// The observer is set up in the capture phase, before Thunderbird inserts the
// extension items, and re-applies the position whenever the items are
// (re)inserted: when the menu is built, and again if an add-on calls
// menus.refresh() while the menu is open.
const menuObservers = new WeakMap();

function isVisible(el) {
  return !el.hidden && el.getAttribute("hidden") != "true";
}

function placeOwnItems(popup) {
  const own = [...popup.children].filter((el) => el.id.startsWith(menuIdPrefix));
  if (!own.length) {
    return;
  }
  const anchor = [...popup.children].find(
    (el) => !own.includes(el) && ["menuitem", "menu"].includes(el.localName) && isVisible(el)
  );
  if (!anchor) {
    return;
  }
  let separator = popup.querySelector(`:scope > menuseparator.${OWN_SEPARATOR_CLASS}`);
  const inPlace =
    separator &&
    anchor.previousElementSibling == separator &&
    separator.previousElementSibling == own[own.length - 1] &&
    own.every((el, i) => i == 0 || el.previousElementSibling == own[i - 1]);
  if (inPlace) {
    return;
  }
  if (!separator) {
    separator = popup.ownerDocument.createXULElement("menuseparator");
    separator.classList.add(OWN_SEPARATOR_CLASS);
  }
  anchor.before(...own, separator);
  // Thunderbird's separator before the extension group may now be dangling.
  for (const sep of popup.querySelectorAll(
    ":scope > menuseparator.webextension-group-separator"
  )) {
    const prev = sep.previousElementSibling;
    const next = sep.nextElementSibling;
    if (!next || next.localName == "menuseparator" || prev?.localName == "menuseparator") {
      sep.hidden = true;
    }
  }
}

function keepOwnItemsOnTop(popup) {
  if (!menuIdPrefix) {
    return;
  }
  placeOwnItems(popup);
  if (menuObservers.has(popup)) {
    return;
  }
  const win = popup.ownerDocument.defaultView;
  const observer = new win.MutationObserver(() => placeOwnItems(popup));
  observer.observe(popup, { childList: true });
  menuObservers.set(popup, observer);
  const cleanup = (hiddenEvent) => {
    if (hiddenEvent.target != popup) {
      return; // a submenu closed
    }
    popup.removeEventListener("popuphidden", cleanup);
    observer.disconnect();
    menuObservers.delete(popup);
    popup.querySelector(`:scope > menuseparator.${OWN_SEPARATOR_CLASS}`)?.remove();
  };
  popup.addEventListener("popuphidden", cleanup);
}

function onPopupShowing(event) {
  const popup = event.target;
  if (popup.id == "mailContext") {
    keepOwnItemsOnTop(popup);
    return;
  }
  if (!SORT_POPUPS.includes(popup.id)) {
    return;
  }
  const doc = popup.ownerDocument;
  let item = doc.getElementById(MENU_ITEM_ID);
  if (!item) {
    item = doc.createXULElement("menuitem");
    item.id = MENU_ITEM_ID;
    item.setAttribute("type", "radio");
    item.setAttribute("name", "sortby");
    item.setAttribute("label", menuLabel);
    item.addEventListener("command", (commandEvent) => {
      // Keep the popup's own cmd_sort handler from running.
      commandEvent.stopPropagation();
      sortPinned(about3PaneForMenu(popup));
    });
    popup.insertBefore(item, popup.firstElementChild);
  }
  // Thunderbird updates the other radio items in its own popupshowing
  // handler, which runs after this capturing listener.
  popup.ownerDocument.defaultView.setTimeout(() => {
    if (isSorted(about3PaneForMenu(popup))) {
      item.setAttribute("checked", "true");
    } else {
      item.removeAttribute("checked");
    }
  }, 0);
}

// F5 is Thunderbird's key_getNewMessages (Get Messages for the selected
// account). With refreshKey on, it runs the add-on's full refresh instead,
// as the ⟳ button does; Shift+F5 and F9 keep Thunderbird's own commands.
// Caught while the command event goes down to the <key>, so Thunderbird's
// own oncommand does not run as well.
function onKeyCommand(event) {
  if (event.target?.id != "key_getNewMessages" || !listState.refreshKey || !listState.listeners.refresh.size) {
    return;
  }
  event.stopPropagation();
  event.preventDefault();
  for (const fire of listState.listeners.refresh) {
    fire.async();
  }
}

function addWindowListener(win) {
  if (!windowsWithListener.has(win)) {
    // Capturing on the 3-pane window also sees popups inside about:3pane.
    win.addEventListener("popupshowing", onPopupShowing, true);
    win.addEventListener("command", onKeyCommand, true);
    windowsWithListener.add(win);
  }
}

function removeFromWindow(win) {
  win.removeEventListener("popupshowing", onPopupShowing, true);
  win.removeEventListener("command", onKeyCommand, true);
  windowsWithListener.delete(win);
  win.document.getElementById(MENU_ITEM_ID)?.remove();
  for (const tabInfo of win.gTabmail?.tabInfo || []) {
    tabInfo.chromeBrowser?.contentDocument?.getElementById(MENU_ITEM_ID)?.remove();
  }
}

function graphServer(accountId) {
  const server = MailServices.accounts.getAccount(accountId)?.incomingServer;
  return server?.type == "graph" ? server : null;
}

// The access token Thunderbird itself uses for the account. Never shows a
// sign-in window: if the account needs a new sign-in, Thunderbird asks for it
// on its own the next time it syncs mail.
async function graphToken(accountId) {
  const server = graphServer(accountId);
  if (!server) {
    throw new ExtensionError(`${accountId} is not a Microsoft Graph account`);
  }
  const oauth = new OAuth2Module();
  if (!oauth.initFromMail(server)) {
    throw new ExtensionError(`No OAuth2 settings for ${accountId}`);
  }
  const sasl = await new Promise((resolve, reject) =>
    oauth.connect(false, {
      onSuccess: resolve,
      onFailure: (code) => reject(new ExtensionError(`No access token (error ${code})`)),
    })
  );
  // connect() returns a SASL XOAUTH2 string: user=...\x01auth=Bearer <token>\x01\x01
  // The experiment scope has no atob(); a window has.
  const win = Services.wm.getMostRecentWindow("mail:3pane") || Services.appShell.hiddenDOMWindow;
  const match = win.atob(sasl).match(/auth=Bearer ([^\x01]+)/);
  if (!match) {
    throw new ExtensionError("Unexpected token format");
  }
  return match[1];
}

/* ------------------------------------------------------------------ */
/* Message list: pin button and refresh button                         */
/* ------------------------------------------------------------------ */

// The cards view's star button (.button-star, which raises "toggle-flag")
// shows Thunderbird's pin icon instead; it is red when the card's message
// has the pin tag. thread-card-tags gets the tag keys as its "tags"
// attribute and comes before the button in the card.
function listCss() {
  const tag = pinKeyword.replace(/[^\w-]/g, "");
  return `
html[${PIN_ATTRIBUTE}] #threadTree[rows="thread-card"] .button-star {
  background-image: var(--icon-pin) !important;
  fill: transparent !important;
  stroke: color-mix(in srgb, currentColor 35%, transparent) !important;
}
html[${PIN_ATTRIBUTE}] #threadTree[rows="thread-card"] tr:is(:hover, .selected) .button-star {
  stroke: color-mix(in srgb, currentColor 80%, transparent) !important;
}
html[${PIN_ATTRIBUTE}] #threadTree[rows="thread-card"] thread-card-tags[tags*="${tag}"] ~ .button-star,
html[${PIN_ATTRIBUTE}] #threadTree[rows="thread-card"] tr[data-outlook-mode-section="pinned"] .button-star {
  fill: #c0392b !important;
  stroke: #c0392b !important;
}
/* Unread mail as in Outlook: a blue bar on the left, a light blue card,
   bold sender, blue bold subject and time. Also a collapsed conversation
   with unread mail in it. The tint is a background image, so the card
   keeps its own colour for hover and selection underneath. */
html[${UNREAD_ATTRIBUTE}] {
  --outlook-mode-unread: light-dark(#0f6cbd, #479ef5);
}
html[${UNREAD_ATTRIBUTE}] #threadTree[rows="thread-card"] .card-layout:is([data-properties~="unread"], .collapsed[data-properties~="hasUnread"]) {
  & .card-container {
    box-shadow: inset 4px 0 0 var(--outlook-mode-unread);
    background-image: linear-gradient(
      color-mix(in srgb, var(--outlook-mode-unread) 10%, transparent),
      color-mix(in srgb, var(--outlook-mode-unread) 10%, transparent)
    );
  }
  & .sender {
    font-weight: bold;
  }
  & :is(.subject, .date) {
    color: var(--outlook-mode-unread) !important;
    font-weight: bold;
  }
}
/* ...and none of Thunderbird's own unread marks (the green dot, the
   sparkle of new mail): the highlight says it already. */
html[${UNREAD_ATTRIBUTE}] #threadTree[rows="thread-card"] .read-status {
  display: none !important;
}
/* Pinned conversations in their own section at the top, as in Outlook.
   The list draws rows of one fixed height, so the room for the headings is
   made by moving the rows down (the list's own scrolling still works); the
   headings sit in that room. */
html[${SECTIONS_ATTRIBUTE}] #threadTree[rows="thread-card"] {
  position: relative;
  & > table {
    margin-top: ${SECTION_HEADER}px;
  }
}
.outlook-mode-section-label {
  position: absolute;
  inset-inline: 0;
  z-index: 1;
  display: flex;
  align-items: flex-end;
  gap: 6px;
  box-sizing: border-box;
  padding: 0 14px 5px;
  font-size: 0.85em;
  font-weight: 600;
  color: color-mix(in srgb, currentColor 70%, transparent);
  pointer-events: none;
  &[hidden] {
    display: none;
  }
  & .count {
    font-weight: normal;
  }
}
.outlook-mode-section-label:not(.first) {
  border-top: 1px solid color-mix(in srgb, currentColor 18%, transparent);
  margin-inline: 8px;
  padding-inline: 6px;
}
/* One card per conversation: no button to open a thread into cards. */
html[${CONVERSATION_ATTRIBUTE}] #threadTree[rows="thread-card"] .thread-card-button.twisty {
  display: none !important;
}
#${REFRESH_BUTTON_ID} {
  background-image: var(--icon-sync);
  background-position: center;
  background-repeat: no-repeat;
  -moz-context-properties: fill, stroke;
  fill: currentColor;
  stroke: currentColor;
  min-width: 28px;
}
#${REFRESH_BUTTON_ID}[busy] {
  animation: outlookModeSpin 1s linear infinite;
}
#${REFRESH_BUTTON_ID}[error] {
  fill: #c0392b;
  stroke: #c0392b;
}
@keyframes outlookModeSpin {
  to { rotate: 360deg; }
}
`;
}

function isCardsView(about3Pane) {
  return about3Pane.threadTree?.getAttribute("rows") == "thread-card";
}

// One card per conversation (cards view, grouped by thread), as in
// Outlook: the reading pane shows the whole conversation, so threads stay
// closed. A thread opened anyway (arrow key, "expand all", going to a mail
// inside it) is closed again, with its conversation selected.
function keepConversationsClosed(about3Pane) {
  const view = about3Pane.gDBView;
  const flags = Ci.nsMsgViewFlagsType;
  if (
    !listState.oneCardPerConversation ||
    !isCardsView(about3Pane) ||
    !view ||
    !(view.viewFlags & flags.kThreadedDisplay) ||
    view.viewFlags & flags.kGroupBySort
  ) {
    return;
  }
  let open = false;
  for (let i = 0; i < view.rowCount && !open; i++) {
    open = view.getLevel(i) > 0;
  }
  if (!open) {
    return;
  }
  const tree = about3Pane.threadTree;
  const selected = tree.selectedIndex >= 0 ? view.getMsgHdrAt(tree.selectedIndex) : null;
  view.doCommand(Ci.nsMsgViewCommandType.collapseAll);
  if (about3Pane.gViewWrapper) {
    about3Pane.gViewWrapper._threadExpandAll = false;
  }
  if (selected) {
    const root = view.getThreadContainingMsgHdr(selected)?.getRootHdr();
    const index = root ? view.findIndexOfMsgHdr(root, false) : -1;
    if (index >= 0 && index != tree.selectedIndex) {
      tree.selectedIndex = index;
    }
  }
}

function applyRefreshState(button) {
  const { busy, title, error } = listState.refresh;
  button.toggleAttribute("busy", busy);
  button.toggleAttribute("error", Boolean(error));
  button.title = title || "Refresh: get new mail, sync settings, pins and contacts";
}

// How many rows at the top are pinned conversations (the list sorted
// pinned first, by the "Pinned" column).
function pinnedRowCount(about3Pane) {
  const view = about3Pane.gDBView;
  let count = 0;
  while (count < view.rowCount) {
    const hdr = view.getMsgHdrAt(count);
    if (!hdr || !conversationPinned(hdr)) {
      break;
    }
    count++;
  }
  return count;
}

// The date a row is sorted by (the same as the "Pinned" sort), in ms.
function rowDate(view, index) {
  const hdr = view.getMsgHdrAt(index);
  return hdr ? sortDate(hdr) * 1000 : 0;
}

// Starts of Today, Yesterday and This week, in ms. This week is the last
// 7 days, today included (not the calendar week).
function dayStarts(now = new Date()) {
  const day = (back) => new Date(now.getFullYear(), now.getMonth(), now.getDate() - back).getTime();
  return { today: day(0), yesterday: day(1), week: day(6) };
}

// The sections of the list, as Outlook: 📌 Pinned, then Today, Yesterday,
// This week (the last 7 days), Older (or "Other mail" when only pinned mail
// is set apart).
// Needs the list sorted pinned first (newest first after the pinned mail),
// in the cards view. Returns [{ name, text, start }], start = first row.
function listSections(about3Pane) {
  const view = about3Pane.gDBView;
  if (!(listState.pinnedSection || listState.dateSections) || !isCardsView(about3Pane) || !view || !isSorted(about3Pane)) {
    return [];
  }
  const rowCount = view.rowCount;
  const pinned = listState.pinnedSection ? pinnedRowCount(about3Pane) : 0;
  const sections = [];
  if (pinned) {
    sections.push({ name: "pinned", text: "📌 Pinned", count: pinned, start: 0 });
  }
  if (pinned >= rowCount) {
    return sections;
  }
  if (!listState.dateSections) {
    if (pinned) {
      sections.push({ name: "other", text: "Other mail", start: pinned });
    }
    return sections;
  }
  // Rows after the pinned ones must be newest first (not if the sort order
  // was turned around): then the first row older than a day start is found
  // by halving.
  if (rowDate(view, pinned) < rowDate(view, rowCount - 1)) {
    return sections;
  }
  const firstOlderThan = (time, from) => {
    let low = from;
    let high = rowCount;
    while (low < high) {
      const mid = (low + high) >> 1;
      if (rowDate(view, mid) >= time) {
        low = mid + 1;
      } else {
        high = mid;
      }
    }
    return low;
  };
  const starts = dayStarts();
  const yesterday = firstOlderThan(starts.today, pinned);
  const week = firstOlderThan(starts.yesterday, yesterday);
  const older = firstOlderThan(starts.week, week);
  for (const [name, text, start, end] of [
    ["today", "Today", pinned, yesterday],
    ["yesterday", "Yesterday", yesterday, week],
    ["week", "This week", week, older],
    ["older", "Older", older, rowCount],
  ]) {
    if (end > start) {
      sections.push({ name, text, start });
    }
  }
  return sections;
}

// Puts each row drawn in its section, and the headings in their place: the
// first heading in the room above the list, each next one in a gap that
// moves the rows below it down.
function layoutSections(about3Pane, recount = true) {
  const doc = about3Pane.document;
  const tree = about3Pane.threadTree;
  const labels = listState.sectionLabels.get(about3Pane);
  if (!tree || !labels) {
    return;
  }
  if (recount) {
    listState.sections.set(about3Pane, listSections(about3Pane));
  }
  const sections = listState.sections.get(about3Pane) || [];
  const on = sections.length > 0;
  doc.documentElement.toggleAttribute(SECTIONS_ATTRIBUTE, on);
  const rowHeight = tree._rowElementClass?.ROW_HEIGHT || 0;
  const used = new Set();
  sections.forEach((section, k) => {
    let label = labels.get(section.name);
    if (!label) {
      label = doc.createElement("div");
      label.className = `outlook-mode-section-label ${section.name}`;
      labels.set(section.name, label);
    }
    if (label.parentNode != tree) {
      tree.append(label);
    }
    label.textContent = section.text;
    if (section.count) {
      const count = doc.createElement("span");
      count.className = "count";
      count.textContent = String(section.count);
      label.append(count);
    }
    label.classList.toggle("first", k == 0);
    label.style.top = k == 0 ? "0" : `${SECTION_HEADER + section.start * rowHeight + (k - 1) * SECTION_GAP}px`;
    label.style.height = `${k == 0 ? SECTION_HEADER : SECTION_GAP}px`;
    label.hidden = false;
    used.add(section.name);
  });
  for (const [name, label] of labels) {
    if (!used.has(name)) {
      label.hidden = true;
    }
  }
  if (tree.table) {
    tree.table.style.marginBottom = on ? `${Math.max(0, sections.length - 1) * SECTION_GAP}px` : "";
  }
  for (const row of tree.table?.body?.children || []) {
    if (typeof row.index != "number") {
      continue;
    }
    let k = -1;
    while (k + 1 < sections.length && sections[k + 1].start <= row.index) {
      k++;
    }
    if (on && k >= 0) {
      row.setAttribute("data-outlook-mode-section", sections[k].name);
      row.style.transform = k > 0 ? `translateY(${k * SECTION_GAP}px)` : "";
    } else {
      row.removeAttribute("data-outlook-mode-section");
      row.style.transform = "";
    }
  }
}

// A closed conversation's card shows, as in Outlook, who sent its newest
// mail and when (your reply in Sent Items included), not its first mail.
// The subject stays the conversation's.
const NEWEST_DATE = new Intl.DateTimeFormat(undefined, { dateStyle: "short", timeStyle: "short" });
const NEWEST_TIME = new Intl.DateTimeFormat(undefined, { timeStyle: "short" });

function authorName(hdr) {
  const author = hdr.mime2DecodedAuthor || "";
  try {
    const [address] = MailServices.headerParser.parseDecodedHeader(author);
    if (address) {
      return address.name || address.email || author;
    }
  } catch (error) {
    // not an address: as it is
  }
  return author;
}

function showNewestInCard(about3Pane, card) {
  const view = about3Pane.gDBView;
  const index = card.index;
  if (!listState.newestInCard || !view || typeof index != "number" || index >= view.rowCount) {
    return;
  }
  const container = view.isContainer(index);
  if (container && view.isContainerOpen(index)) {
    return; // an open thread: each card is its own mail
  }
  const hdr = view.getMsgHdrAt(index);
  // A single mail is a conversation too when you answered it.
  const hdrs = container ? threadHdrs(hdr) : [hdr];
  let newest = { date: hdr.dateInSeconds, author: hdr.mime2DecodedAuthor };
  for (const child of hdrs) {
    if (child.dateInSeconds > newest.date) {
      newest = { date: child.dateInSeconds, author: child.mime2DecodedAuthor };
    }
  }
  const reply = newestOwnReply(hdrs);
  if (reply && reply.date > newest.date) {
    newest = reply;
  }
  if (newest.date == hdr.dateInSeconds && newest.author == hdr.mime2DecodedAuthor) {
    return;
  }
  const name = authorName({ mime2DecodedAuthor: newest.author });
  card.senderLine.textContent = name;
  card.senderLine.title = newest.author || name;
  const date = new Date(newest.date * 1000);
  const today = new Date();
  card.dateLine.textContent = (
    date.toDateString() == today.toDateString() ? NEWEST_TIME : NEWEST_DATE
  ).format(date);
}

// Wraps the card's fillRow, once per window, to call showNewestInCard
// after Thunderbird fills the card.
function patchCards(about3Pane) {
  const ThreadCard = about3Pane.customElements.get("thread-card");
  if (!ThreadCard || ThreadCard.prototype._outlookModeFillRow) {
    return;
  }
  const original = ThreadCard.prototype.fillRow;
  ThreadCard.prototype._outlookModeFillRow = original;
  ThreadCard.prototype.fillRow = function (...args) {
    const result = original.apply(this, args);
    try {
      showNewestInCard(about3Pane, this);
    } catch (error) {
      console.error("[outlook-mode] newest mail in card", error);
    }
    return result;
  };
}

function unpatchCards(about3Pane) {
  const ThreadCard = about3Pane.customElements.get("thread-card");
  if (ThreadCard?.prototype._outlookModeFillRow) {
    ThreadCard.prototype.fillRow = ThreadCard.prototype._outlookModeFillRow;
    delete ThreadCard.prototype._outlookModeFillRow;
  }
}

function listUpdate(about3Pane) {
  const doc = about3Pane.document;
  doc.documentElement.toggleAttribute(PIN_ATTRIBUTE, listState.pinButton);
  doc.documentElement.toggleAttribute(UNREAD_ATTRIBUTE, listState.unreadHighlight);
  doc.documentElement.toggleAttribute(CONVERSATION_ATTRIBUTE, listState.oneCardPerConversation);
  keepConversationsClosed(about3Pane);
  layoutSections(about3Pane);
  const button = doc.getElementById(REFRESH_BUTTON_ID);
  if (button) {
    button.hidden = !listState.refreshButton;
    applyRefreshState(button);
  }
}

function listInject(about3Pane) {
  if (!about3Pane || about3Pane.closed || listState.panes.has(about3Pane)) {
    if (about3Pane && listState.panes.has(about3Pane)) {
      listUpdate(about3Pane);
    }
    return;
  }
  const doc = about3Pane.document;
  const end = doc.querySelector("#threadPaneHeaderBar .list-header-bar-container-end");
  if (!end || !doc.head) {
    // The list header is not built yet: try again once the page has loaded.
    if (doc.readyState != "complete" && !listState.waiting.has(about3Pane)) {
      listState.waiting.add(about3Pane);
      about3Pane.addEventListener(
        "load",
        () => {
          listState.waiting.delete(about3Pane);
          listInject(about3Pane);
        },
        { once: true }
      );
    }
    return;
  }
  const style = doc.createElement("style");
  style.id = LIST_STYLE_ID;
  style.textContent = listCss();
  doc.head.appendChild(style);

  // Runs before about:3pane's own "toggle-flag" listener on the tree.
  // Thunderbird raises "toggle-flag" on the list's table body when the
  // star of a row is clicked; anything else (another target, an event not
  // made by Thunderbird, a row that does not exist) is ignored.
  const onToggleFlag = (event) => {
    if (!listState.pinButton || !isCardsView(about3Pane)) {
      return;
    }
    event.stopImmediatePropagation();
    const index = event.detail?.index;
    if (
      !event.isTrusted ||
      event.target != about3Pane.threadTree?.table?.body ||
      !Number.isInteger(index) ||
      index < 0 ||
      index >= (about3Pane.gDBView?.rowCount ?? 0)
    ) {
      return;
    }
    try {
      const hdr = about3Pane.gDBView?.getMsgHdrAt(index);
      const id = hdr && listState.extension?.messageManager.convert(hdr).id;
      if (id) {
        for (const fire of listState.listeners.pin) {
          fire.async(id);
        }
      }
    } catch (error) {
      console.error("[outlook-mode] pin button", error);
    }
  };
  doc.addEventListener("toggle-flag", onToggleFlag, true);
  patchCards(about3Pane);

  // Section headings, made and placed by layoutSections.
  listState.sectionLabels.set(about3Pane, new Map());
  // A new day moves mail from Today to Yesterday.
  let day = new Date().toDateString();
  const dayTimer = about3Pane.setInterval(() => {
    if (new Date().toDateString() != day) {
      day = new Date().toDateString();
      layoutSections(about3Pane);
    }
  }, 60e3);
  // Rows are drawn as the list scrolls: put new ones in their section. A
  // card whose tags change (pinned or unpinned, also by pin sync) makes the
  // pinned rows counted again.
  let recountTimer = null;
  const rowObserver = new about3Pane.MutationObserver((mutations) => {
    if (mutations.some((m) => m.type == "attributes")) {
      about3Pane.clearTimeout(recountTimer);
      recountTimer = about3Pane.setTimeout(() => layoutSections(about3Pane), 50);
    } else {
      layoutSections(about3Pane, false);
    }
  });
  const observeRows = () => {
    rowObserver.disconnect();
    const body = about3Pane.threadTree?.table?.body;
    if (body) {
      rowObserver.observe(body, { childList: true, subtree: true, attributes: true, attributeFilter: ["tags"] });
    }
  };
  observeRows();

  // Checked once things settle, after whatever opened a thread.
  let closeTimer = null;
  const onListChange = (event) => {
    if (event.type == "viewchange" && about3Pane.gFolder) {
      listRedraw(about3Pane.gFolder); // a folder shown: its counts checked
    }
    about3Pane.clearTimeout(closeTimer);
    closeTimer = about3Pane.setTimeout(() => {
      try {
        keepConversationsClosed(about3Pane);
      } catch (error) {
        console.error("[outlook-mode] one card per conversation", error);
      }
      try {
        observeRows();
        layoutSections(about3Pane, event.type != "select"); // a selection moves no section
        // Rows sit lower than the list thinks: keep the selected one in view.
        if (event.type == "select" && about3Pane.document.documentElement.hasAttribute(SECTIONS_ATTRIBUTE)) {
          const tree = about3Pane.threadTree;
          tree.getRowAtIndex(tree.currentIndex)?.scrollIntoView({ block: "nearest" });
        }
      } catch (error) {
        console.error("[outlook-mode] pinned section", error);
      }
    }, 0);
  };
  const LIST_EVENTS = ["expanded", "rowcountchange", "viewchange", "select"];
  for (const name of LIST_EVENTS) {
    doc.addEventListener(name, onListChange, true);
  }

  const button = doc.createElement("button");
  button.id = REFRESH_BUTTON_ID;
  button.className = "button button-flat icon-button icon-only";
  button.type = "button";
  button.addEventListener("click", () => {
    if (listState.refresh.busy) {
      return;
    }
    for (const fire of listState.listeners.refresh) {
      fire.async();
    }
  });
  end.prepend(button);

  listState.panes.set(about3Pane, () => {
    doc.removeEventListener("toggle-flag", onToggleFlag, true);
    for (const name of LIST_EVENTS) {
      doc.removeEventListener(name, onListChange, true);
    }
    rowObserver.disconnect();
    unpatchCards(about3Pane);
    about3Pane.clearInterval(dayTimer);
    for (const label of listState.sectionLabels.get(about3Pane)?.values() || []) {
      label.remove();
    }
    listState.sectionLabels.delete(about3Pane);
    listState.sections.delete(about3Pane);
    for (const row of about3Pane.threadTree?.table?.body?.children || []) {
      row.removeAttribute("data-outlook-mode-section");
      row.style.transform = "";
    }
    if (about3Pane.threadTree?.table) {
      about3Pane.threadTree.table.style.marginBottom = "";
    }
    style.remove();
    button.remove();
    doc.documentElement.removeAttribute(PIN_ATTRIBUTE);
    doc.documentElement.removeAttribute(UNREAD_ATTRIBUTE);
    doc.documentElement.removeAttribute(CONVERSATION_ATTRIBUTE);
    doc.documentElement.removeAttribute(SECTIONS_ATTRIBUTE);
  });
  about3Pane.addEventListener("unload", () => listState.panes.delete(about3Pane), { once: true });
  listUpdate(about3Pane);
}

// Runs fn(win) once the document of win is parsed. A document still loading
// when the add-on starts (Thunderbird opening its window as the add-on
// loads) was created before the document-element-inserted observer was
// added, so it would otherwise be missed until Thunderbird restarts.
function whenParsed(win, fn) {
  if (win.document.readyState == "loading") {
    win.addEventListener("DOMContentLoaded", () => fn(win), { once: true });
  } else {
    fn(win);
  }
}

function listScan() {
  for (const win of Services.wm.getEnumerator("mail:3pane")) {
    for (const tab of win.gTabmail?.tabInfo || []) {
      const about3Pane = tab.mode?.name == "mail3PaneTab" ? tab.chromeBrowser?.contentWindow : null;
      if (about3Pane && !about3Pane.closed) {
        whenParsed(about3Pane, listInject);
      }
    }
  }
}

const listObserver = {
  observe(doc) {
    if (doc.documentURI == "about:3pane" && doc.defaultView) {
      const win = doc.defaultView;
      win.addEventListener("DOMContentLoaded", () => listInject(win), { once: true });
    }
  },
};

// A closed thread's card is not drawn again when mail hidden in it changes
// (read here, read on another device and synced, or marked read only once
// the server answers, as for Microsoft 365 accounts): the list showing the
// folder is drawn again whenever the read state of its mail changes.
const listFolderListener = {
  onFolderPropertyFlagChanged(item, property) {
    if (property == "Status") {
      listRedraw(item.QueryInterface(Ci.nsIMsgDBHdr).folder);
    }
  },
  onFolderIntPropertyChanged(folder, property) {
    if (property == "TotalUnreadMessages") {
      listRedraw(folder);
    }
  },
};

// Thunderbird keeps, per thread, a count of its unread mail, and a closed
// thread's card says "unread" from that count. For Microsoft 365 (Graph)
// accounts the count is not always lowered when mail is read, so a
// conversation read long ago still shows as unread. Counted again from the
// mail itself and set right; the mail is not touched.
function repairThreadCounts(folder) {
  let db;
  try {
    db = folder.msgDatabase;
  } catch (error) {
    return 0; // not opened yet
  }
  let fixed = 0;
  for (const thread of db.enumerateThreads()) {
    let unread = 0;
    for (let i = 0; i < thread.numChildren; i++) {
      const hdr = thread.getChildHdrAt(i);
      if (hdr && !hdr.isRead) {
        unread++;
      }
    }
    let diff = thread.numUnreadChildren - unread;
    if (diff) {
      fixed++;
    }
    for (; diff > 0; diff--) {
      thread.markChildRead(true);
    }
    for (; diff < 0; diff++) {
      thread.markChildRead(false);
    }
  }
  if (fixed) {
    db.commit(Ci.nsMsgDBCommitType.kLargeCommit);
    listState.threadCountsRepaired += fixed;
  }
  return fixed;
}

function listRedraw(folder) {
  if (!listState.repairs.has(folder)) {
    listState.repairs.add(folder);
    Services.tm.dispatchToMainThread(() => {
      listState.repairs.delete(folder);
      try {
        repairThreadCounts(folder);
      } catch (error) {
        console.error("[outlook-mode] thread counts", error);
      }
    });
  }
  for (const about3Pane of listState.panes.keys()) {
    if (about3Pane.closed || !(about3Pane.gFolder == folder || about3Pane.gViewWrapper?.isSynthetic)) {
      continue;
    }
    if (!listState.redraws.has(about3Pane)) {
      listState.redraws.add(about3Pane);
      about3Pane.setTimeout(() => {
        listState.redraws.delete(about3Pane);
        about3Pane.threadTree?.invalidate();
      }, 50);
    }
  }
}

function listRemoveAll() {
  if (listState.observing) {
    Services.obs.removeObserver(listObserver, "document-element-inserted");
    MailServices.mailSession.RemoveFolderListener(listFolderListener);
    listState.observing = false;
  }
  for (const cleanup of listState.panes.values()) {
    try {
      cleanup();
    } catch (error) {
      // the window is closing
    }
  }
  listState.panes.clear();
}

// Thunderbird's Get Messages for every signed-in account, and a sync of the
// folders shown in the mail tabs (for folders other than the Inbox).
function getNewMessages() {
  let started = false;
  for (const win of Services.wm.getEnumerator("mail:3pane")) {
    if (!started) {
      win.goDoCommand("cmd_getMsgsForAuthAccounts");
      started = true;
    }
    for (const tab of win.gTabmail?.tabInfo || []) {
      const folder = tab.mode?.name == "mail3PaneTab" ? tab.chromeBrowser?.contentWindow?.gFolder : null;
      if (folder && !folder.isServer) {
        try {
          folder.updateFolder(win.msgWindow);
        } catch (error) {
          // a folder that cannot be updated (virtual, offline)
        }
      }
    }
  }
  return started;
}

this.outlookPin = class extends ExtensionCommon.ExtensionAPI {
  onShutdown(isAppShutdown) {
    if (isAppShutdown) {
      return;
    }
    if (registered) {
      ThreadPaneColumns.removeCustomColumn(COLUMN_ID);
      registered = false;
    }
    if (listenerName) {
      ExtensionSupport.unregisterWindowListener(listenerName);
      listenerName = "";
    }
    if (resortListening) {
      MailServices.mfn.removeListener(resortListener);
      MailServices.mailSession.RemoveFolderListener(resortFolderListener);
      resortListening = false;
    }
    for (const [about3Pane, timer] of resortTimers) {
      about3Pane.clearTimeout(timer);
    }
    resortTimers.clear();
    for (const win of [...windowsWithListener]) {
      removeFromWindow(win);
    }
    listRemoveAll();
    // Make sure a new version of this file is loaded after an update.
    Services.obs.notifyObservers(null, "startupcache-invalidate");
  }

  getAPI(context) {
    const aboutPaneForTab = (tabId) => {
      const tab = context.extension.tabManager.get(tabId);
      const native = tab?.nativeTab;
      if (native?.mode?.name != "mail3PaneTab") {
        return null;
      }
      return native.chromeBrowser?.contentWindow || null;
    };
    const event = (name, set) =>
      new ExtensionCommon.EventManager({
        context,
        name: `outlookPin.${name}`,
        register: (fire) => {
          set.add(fire);
          return () => set.delete(fire);
        },
      }).api();
    return {
      outlookPin: {
        async configureMailList(options) {
          listState.extension = context.extension;
          listState.pinButton = Boolean(options.pinButton);
          listState.refreshButton = Boolean(options.refreshButton);
          listState.refreshKey = Boolean(options.refreshKey);
          listState.unreadHighlight = Boolean(options.unreadHighlight);
          listState.oneCardPerConversation = Boolean(options.oneCardPerConversation);
          listState.pinnedSection = Boolean(options.pinnedSection);
          listState.dateSections = Boolean(options.dateSections);
          const newestChanged = listState.newestInCard != Boolean(options.newestInCard);
          listState.newestInCard = Boolean(options.newestInCard);
          if (
            !listState.pinButton &&
            !listState.refreshButton &&
            !listState.unreadHighlight &&
            !listState.oneCardPerConversation &&
            !listState.pinnedSection &&
            !listState.dateSections &&
            !listState.newestInCard
          ) {
            listRemoveAll();
            return;
          }
          if (!listState.observing) {
            Services.obs.addObserver(listObserver, "document-element-inserted");
            MailServices.mailSession.AddFolderListener(
              listFolderListener,
              Ci.nsIFolderListener.propertyFlagChanged | Ci.nsIFolderListener.intPropertyChanged
            );
            listState.observing = true;
          }
          listScan();
          for (const about3Pane of listState.panes.keys()) {
            listUpdate(about3Pane);
            if (newestChanged) {
              about3Pane.threadTree?.invalidate(); // cards drawn again
            }
          }
        },
        async setRefreshState(state) {
          listState.refresh = {
            busy: Boolean(state.busy),
            title: String(state.title || ""),
            error: Boolean(state.error),
          };
          for (const about3Pane of listState.panes.keys()) {
            const button = about3Pane.document.getElementById(REFRESH_BUTTON_ID);
            if (button) {
              applyRefreshState(button);
            }
          }
        },
        async getNewMessages() {
          return getNewMessages();
        },
        onPinButton: event("onPinButton", listState.listeners.pin),
        onRefresh: event("onRefresh", listState.listeners.refresh),
        async register(options) {
          pinKeyword = options.pinKeyword || pinKeyword;
          menuIdPrefix = `${ExtensionCommon.makeWidgetId(context.extension.id)}-menuitem-`;
          menuLabel = options.menuLabel || menuLabel;
          if (!registered) {
            ThreadPaneColumns.addCustomColumn(COLUMN_ID, {
              name: options.columnName || "Pinned",
              hidden: true,
              resizable: false,
              sortable: true,
              textCallback: (hdr) => (hasPin(hdr) ? "📌" : ""),
              sortCallback: sortKey,
            });
            registered = true;
          }
          if (!resortListening) {
            MailServices.mfn.addListener(resortListener, MailServices.mfn.msgAdded);
            MailServices.mailSession.AddFolderListener(resortFolderListener, Ci.nsIFolderListener.intPropertyChanged);
            resortListening = true;
          }
          if (!listenerName) {
            listenerName = `ext-${context.extension.id}-pin-sort`;
            ExtensionSupport.registerWindowListener(listenerName, {
              chromeURLs: ["chrome://messenger/content/messenger.xhtml"],
              onLoadWindow: addWindowListener,
            });
          }
        },
        async sortByPinned(tabId) {
          return sortPinned(aboutPaneForTab(tabId));
        },
        async isSortedByPinned(tabId) {
          return isSorted(aboutPaneForTab(tabId));
        },
        async listGraphAccounts() {
          return MailServices.accounts.accounts
            .filter((account) => account.incomingServer?.type == "graph")
            .map((account) => account.key);
        },
        async getGraphToken(accountId) {
          return graphToken(accountId);
        },
        // The messages of the conversation (thread, in its folder) of a
        // message, the message itself included.
        async conversationMessageIds(messageId) {
          const hdr = context.extension.messageManager.get(messageId);
          if (!hdr) {
            return [];
          }
          const ids = new Set([messageId]);
          try {
            const thread = hdr.folder.msgDatabase.getThreadContainingMsgHdr(hdr);
            for (let i = 0; thread && i < thread.numChildren; i++) {
              const child = thread.getChildHdrAt(i);
              if (child) {
                ids.add(context.extension.messageManager.convert(child).id);
              }
            }
          } catch (error) {
            // no thread information: the message alone
          }
          return [...ids];
        },
        async threadCountsRepaired() {
          return listState.threadCountsRepaired;
        },
        // Marks a conversation as read, as Outlook does when it is opened.
        // messageIds are the selected messages of a mail tab: one row
        // selected is one conversation (a closed thread selects all its
        // mail; for a single mail, the rest of its thread is added). Rows
        // selected on purpose are left alone. Follows Thunderbird's "mark
        // as read when displayed". Returns what happened, for Diagnostics.
        async markThreadRead(tabId, messageIds) {
          const result = { selected: messageIds.length, found: 0, rows: 0, unreadBefore: 0, marked: 0, card: "", reason: "" };
          if (!Services.prefs.getBoolPref("mailnews.mark_message_read.auto", true)) {
            result.reason = "Thunderbird is set not to mark displayed mail as read";
            return result;
          }
          const about3Pane = aboutPaneForTab(tabId);
          const tree = about3Pane?.threadTree;
          result.rows = tree ? tree.selectedIndices.length : 0;
          if (result.rows != 1) {
            result.reason = result.rows ? "several rows selected" : "no row selected";
            return result;
          }
          // The mail marked is the mail the tab really has selected, read
          // here; the ids given only have to be that selection (a caller
          // cannot have other mail marked read).
          const hdrs = about3Pane.gDBView?.getSelectedMsgHdrs() || [];
          const selectedIds = new Set(hdrs.map((hdr) => context.extension.messageManager.convert(hdr).id));
          result.found = hdrs.length;
          if (!hdrs.length || messageIds.length != selectedIds.size || !messageIds.every((id) => selectedIds.has(id))) {
            result.reason = "not the mail selected in the list";
            return result;
          }
          if (hdrs.length == 1) {
            try {
              const thread = hdrs[0].folder.msgDatabase.getThreadContainingMsgHdr(hdrs[0]);
              for (let i = 0; thread && i < thread.numChildren; i++) {
                const child = thread.getChildHdrAt(i);
                if (child && child != hdrs[0]) {
                  hdrs.push(child);
                }
              }
            } catch (error) {
              // no thread information: the one mail only
            }
          }
          const byFolder = new Map();
          for (const hdr of hdrs) {
            if (!hdr.isRead && hdr.folder) {
              byFolder.set(hdr.folder, [...(byFolder.get(hdr.folder) || []), hdr]);
            }
          }
          for (const [folder, unread] of byFolder) {
            result.unreadBefore += unread.length;
            try {
              folder.markMessagesRead(unread, true);
              result.marked += unread.length;
            } catch (error) {
              result.reason = `${folder.localizedName}: ${error.message || error}`;
            }
          }
          // A closed thread's card does not always notice that mail hidden
          // in it was read: draw it again (and once more after a while, for
          // servers that answer late).
          const index = tree.selectedIndex;
          tree.invalidateRow(index);
          await new Promise((resolve) => about3Pane.setTimeout(resolve, 1500));
          if (!about3Pane.closed && about3Pane.threadTree) {
            about3Pane.threadTree.invalidate();
            result.card = about3Pane.threadTree.getRowAtIndex(about3Pane.threadTree.selectedIndex)?.dataset.properties || "";
          }
          return result;
        },
        // From the flag Thunderbird keeps in the message database, so no
        // message has to be read (messages.query({attachment}) parses, and
        // may download, every message).
        async withAttachments(messageIds) {
          const messageManager = context.extension.messageManager;
          return messageIds.filter(
            (id) => messageManager.get(id)?.flags & Ci.nsMsgMessageFlags.Attachment
          );
        },
      },
    };
  }
};
