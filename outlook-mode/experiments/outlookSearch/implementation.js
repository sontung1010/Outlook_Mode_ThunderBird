"use strict";

/* global ExtensionAPI, ExtensionCommon, ChromeUtils, Services, Cc, Ci, Cr */

// Experiment: Outlook-style search in the main window.
//
// Thunderbird's search box at the top (global-search-bar) suggests Gloda
// results and opens every search in a new tab. Here it works like Outlook's:
// - while typing, a list under the box suggests people, matching mail and
//   recent searches;
// - Enter shows the results in the message list of the mail tab itself
//   (a synthetic view, as Gloda's "show as list" uses), with a bar to pick
//   all folders or the current one, quick filters and Exit;
// - Esc, Exit, or opening a folder goes back.
//
// The search reads the message headers of every mail folder (subject,
// sender, To/Cc names and addresses), ignoring case and diacritics, and adds
// Gloda's full-text index for the message text. It understands Outlook's
// keywords: from: to: cc: subject: body: has:attachment is:unread is:read
// is:flagged is:pinned before: after: (plus participants:, used by people
// suggestions).

var { ExtensionSupport } = ChromeUtils.importESModule(
  "resource:///modules/ExtensionSupport.sys.mjs"
);
var { MailServices } = ChromeUtils.importESModule("resource:///modules/MailServices.sys.mjs");

const SEARCH_LISTENER = "outlookMode-search";
const RECENT_PREF = "extensions.outlookmode.search.recent";
const SEARCH_ENABLED_PREF = "extensions.outlookmode.search.enabled";
const MAX_RESULTS = 2000;
const CORPUS_TTL_MS = 60 * 1000; // a search reads the folders again after this
const SUGGEST_TTL_MS = 5 * 60 * 1000; // suggestions may use an older read
const SKIPPED_FOLDERS = Ci.nsMsgFolderFlags.Trash | Ci.nsMsgFolderFlags.Junk | Ci.nsMsgFolderFlags.Virtual;
const STYLE_ID = "outlookModeSearchStyle";
const POPUP_ID = "outlookModeSearchSuggest";
const BAR_ID = "outlookModeSearchBar";
const PIN_KEYWORD = "olpinned";
const FILTERS = [
  ["is:unread", "Unread"],
  ["has:attachment", "📎 Attachments"],
  ["is:pinned", "📌 Pinned"],
  ["is:tome", "To me"],
  ["is:fromme", "From me"],
];

const searchState = {
  enabled: false,
  windows: new Map(), // messenger window -> cleanup
  modes: new Map(), // about:3pane window -> search mode
  corpus: null,
  corpusAt: 0,
  building: null,
  lastError: "",
};

/* ------------------------------------------------------------------ */
/* Text                                                                */
/* ------------------------------------------------------------------ */

function fold(text) {
  return String(text || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .toLowerCase();
}

function idle() {
  return new Promise((resolve) => Services.tm.dispatchToMainThread(resolve));
}

function note(where, error) {
  searchState.lastError = `${new Date().toISOString()} ${where}: ${error && error.message ? error.message : error}`;
  console.error(`[outlook-mode] search: ${where}`, error);
}

function addresses(raw) {
  try {
    return MailServices.headerParser.parseEncodedHeaderW(raw || "");
  } catch (error) {
    return [];
  }
}

function myAddresses() {
  return new Set(MailServices.accounts.allIdentities.map((i) => (i.email || "").toLowerCase()).filter(Boolean));
}

/* ------------------------------------------------------------------ */
/* The messages: one entry per message, kept for a minute              */
/* ------------------------------------------------------------------ */

function entryOf(hdr, folder, mine) {
  const author = hdr.mime2DecodedAuthor || "";
  const to = addresses(hdr.recipients);
  const cc = [...addresses(hdr.ccList), ...addresses(hdr.bccList)];
  const text = (list) => list.map((a) => `${a.name || ""} ${a.email || ""}`).join(" ");
  const authorAddress = addresses(hdr.author)[0];
  const authorEmail = (authorAddress?.email || "").toLowerCase();
  return {
    // Everyone in it, for the People suggestions.
    people: [authorAddress, ...to, ...cc].filter((a) => a && a.email),
    folder,
    key: hdr.messageKey,
    date: hdr.dateInSeconds,
    subject: hdr.mime2DecodedSubject || "",
    author,
    fSubject: fold(hdr.mime2DecodedSubject),
    fAuthor: fold(`${author} ${authorEmail}`),
    fTo: fold(text(to)),
    fCc: fold(text(cc)),
    flags: hdr.flags,
    pinned: (hdr.getStringProperty("keywords") || "").split(" ").includes(PIN_KEYWORD),
    fromMe: mine.has(authorEmail),
    toMe: [...to, ...cc].some((a) => mine.has((a.email || "").toLowerCase())),
  };
}

async function buildCorpus() {
  const mine = myAddresses();
  const entries = [];
  const everyone = new Map(); // email -> { name, email, count }
  for (const server of MailServices.accounts.allServers) {
    if (["rss", "nntp"].includes(server.type)) {
      continue;
    }
    for (const folder of server.rootFolder.descendants) {
      if (folder.flags & SKIPPED_FOLDERS) {
        continue;
      }
      let db;
      try {
        db = folder.msgDatabase;
      } catch (error) {
        continue; // no summary for this folder
      }
      for (const hdr of db.enumerateMessages()) {
        const entry = entryOf(hdr, folder, mine);
        for (const person of entry.people) {
          const email = person.email.toLowerCase();
          if (mine.has(email)) {
            continue;
          }
          const known = everyone.get(email) || { name: "", email, count: 0 };
          known.count++;
          if (person.name && person.name.toLowerCase() != email) {
            known.name = person.name;
          }
          everyone.set(email, known);
        }
        delete entry.people;
        entries.push(entry);
        if (entries.length % 2000 == 0) {
          await idle();
        }
      }
      try {
        folder.closeDBIfFolderNotOpen(false);
      } catch (error) {
        // still in use
      }
    }
  }
  entries.sort((a, b) => b.date - a.date);
  searchState.correspondents = [...everyone.values()]
    .sort((a, b) => b.count - a.count)
    .map((p) => ({ ...p, key: fold(`${p.name} ${p.email}`) }));
  return entries;
}

function getCorpus(maxAge) {
  if (searchState.corpus && Date.now() - searchState.corpusAt < maxAge) {
    return Promise.resolve(searchState.corpus);
  }
  if (!searchState.building) {
    const started = Date.now();
    searchState.building = buildCorpus()
      .then((entries) => {
        searchState.corpus = entries;
        searchState.corpusAt = started;
        return entries;
      })
      .finally(() => {
        searchState.building = null;
      });
  }
  return searchState.building;
}

function hdrOf(entry) {
  try {
    return entry.folder.msgDatabase.getMsgHdrForKey(entry.key);
  } catch (error) {
    return null;
  }
}

/* ------------------------------------------------------------------ */
/* Queries                                                             */
/* ------------------------------------------------------------------ */

function parseDate(text) {
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  const dmy = text.match(/^(\d{1,2})[/.](\d{1,2})[/.](\d{4})$/);
  const date = iso ? new Date(+iso[1], iso[2] - 1, +iso[3]) : dmy ? new Date(+dmy[3], dmy[2] - 1, +dmy[1]) : null;
  return date && !Number.isNaN(date.getTime()) ? Math.floor(date.getTime() / 1000) : null;
}

// Outlook's search keywords, quoted phrases and free words.
function parseQuery(text) {
  const query = {
    free: [], // folded words to find in subject, people or text
    raw: [], // the same words as typed, for the full-text index
    from: [],
    to: [],
    cc: [],
    people: [],
    subject: [],
    body: [],
    has: new Set(),
    is: new Set(),
    before: null,
    after: null,
  };
  const re = /(\w+):("[^"]*"|\S+)|"([^"]*)"|(\S+)/g;
  let m;
  while ((m = re.exec(text))) {
    if (m[1]) {
      const field = m[1].toLowerCase();
      const value = m[2].replace(/^"|"$/g, "");
      const folded = fold(value).trim();
      if (!folded) {
        continue;
      }
      if (field == "from") {
        query.from.push(folded);
      } else if (field == "to") {
        query.to.push(folded);
      } else if (field == "cc") {
        query.cc.push(folded);
      } else if (field == "participants" || field == "with") {
        query.people.push(folded);
      } else if (field == "subject") {
        query.subject.push(folded);
      } else if (field == "body") {
        query.body.push(value);
      } else if (field == "has" || field == "hasattachments") {
        query.has.add(["yes", "true", "attachments"].includes(folded) ? "attachment" : folded);
      } else if (field == "is") {
        query.is.add(folded);
      } else if (field == "before" || field == "received<") {
        query.before = parseDate(value);
      } else if (field == "after" || field == "received>") {
        query.after = parseDate(value);
      } else {
        query.free.push(fold(m[0]));
        query.raw.push(m[0]);
      }
    } else {
      const word = m[3] ?? m[4];
      if (fold(word).trim()) {
        query.free.push(fold(word).trim());
        query.raw.push(m[3] != null ? `"${word}"` : word);
      }
    }
  }
  return query;
}

function isEmptyQuery(query) {
  return !(
    query.free.length || query.from.length || query.to.length || query.cc.length || query.people.length ||
    query.subject.length || query.body.length || query.has.size || query.is.size || query.before || query.after
  );
}

function headerHas(entry, word) {
  return (
    entry.fSubject.includes(word) || entry.fAuthor.includes(word) || entry.fTo.includes(word) || entry.fCc.includes(word)
  );
}

function matches(entry, query, textHits, bodyHits) {
  const key = `${entry.folder.URI}#${entry.key}`;
  if (query.from.some((w) => !entry.fAuthor.includes(w))) {
    return false;
  }
  if (query.to.some((w) => !entry.fTo.includes(w) && !entry.fCc.includes(w))) {
    return false;
  }
  if (query.cc.some((w) => !entry.fCc.includes(w))) {
    return false;
  }
  if (query.people.some((w) => !entry.fAuthor.includes(w) && !entry.fTo.includes(w) && !entry.fCc.includes(w))) {
    return false;
  }
  if (query.subject.some((w) => !entry.fSubject.includes(w))) {
    return false;
  }
  if (query.has.has("attachment") && !(entry.flags & Ci.nsMsgMessageFlags.Attachment)) {
    return false;
  }
  const read = Boolean(entry.flags & Ci.nsMsgMessageFlags.Read);
  if ((query.is.has("unread") && read) || (query.is.has("read") && !read)) {
    return false;
  }
  if (query.is.has("flagged") && !(entry.flags & Ci.nsMsgMessageFlags.Marked)) {
    return false;
  }
  if (query.is.has("pinned") && !entry.pinned) {
    return false;
  }
  if ((query.is.has("tome") && !entry.toMe) || (query.is.has("fromme") && !entry.fromMe)) {
    return false;
  }
  if ((query.before && entry.date >= query.before) || (query.after && entry.date < query.after)) {
    return false;
  }
  if (query.body.length && !(bodyHits && bodyHits.has(key))) {
    return false;
  }
  if (query.free.length && !query.free.every((w) => headerHas(entry, w)) && !(textHits && textHits.has(key))) {
    return false;
  }
  return true;
}

// Gloda's full-text index (message text, attachment names): the keys of
// the messages that have all the words. null when it is not available.
function glodaKeys(words) {
  if (!words.length || !Services.prefs.getBoolPref("mailnews.database.global.indexer.enabled", true)) {
    return Promise.resolve(null);
  }
  let GlodaMsgSearcher;
  try {
    ({ GlodaMsgSearcher } = ChromeUtils.importESModule("resource:///modules/gloda/GlodaMsgSearcher.sys.mjs"));
  } catch (error) {
    return Promise.resolve(null);
  }
  return new Promise((resolve) => {
    let done = false;
    const timer = Cc["@mozilla.org/timer;1"].createInstance(Ci.nsITimer);
    const finish = (keys) => {
      if (!done) {
        done = true;
        timer.cancel();
        resolve(keys);
      }
    };
    timer.initWithCallback({ notify: () => finish(null) }, 8000, Ci.nsITimer.TYPE_ONE_SHOT);
    try {
      new GlodaMsgSearcher(
        {
          onItemsAdded() {},
          onItemsModified() {},
          onItemsRemoved() {},
          onQueryCompleted(collection) {
            finish(new Set(collection.items.map((item) => `${item.folderURI}#${item.messageKey}`)));
          },
        },
        words.join(" "),
        true
      ).getCollection();
    } catch (error) {
      note("full-text search", error);
      finish(null);
    }
  });
}

async function search(text, { folder = null } = {}) {
  const query = parseQuery(text);
  if (isEmptyQuery(query)) {
    return [];
  }
  const [corpus, textHits, bodyHits] = await Promise.all([
    getCorpus(CORPUS_TTL_MS),
    glodaKeys(query.raw),
    glodaKeys(query.body),
  ]);
  const hits = [];
  for (const entry of corpus) {
    if ((!folder || entry.folder == folder) && matches(entry, query, textHits, bodyHits)) {
      hits.push(entry);
      if (hits.length >= MAX_RESULTS) {
        break;
      }
    }
  }
  return hits;
}

/* ------------------------------------------------------------------ */
/* Recent searches and people                                          */
/* ------------------------------------------------------------------ */

function recentSearches() {
  try {
    return JSON.parse(Services.prefs.getStringPref(RECENT_PREF, "[]"));
  } catch (error) {
    return [];
  }
}

function rememberSearch(text) {
  const list = [text, ...recentSearches().filter((t) => t != text)].slice(0, 8);
  Services.prefs.setStringPref(RECENT_PREF, JSON.stringify(list));
}

let peopleCache = null;
function people() {
  if (peopleCache && Date.now() - peopleCache.at < SUGGEST_TTL_MS) {
    return peopleCache.list;
  }
  const seen = new Set();
  const list = [];
  for (const dir of MailServices.ab.directories) {
    if (dir.isRemote) {
      continue;
    }
    let cards = [];
    try {
      cards = dir.childCards;
    } catch (error) {
      continue;
    }
    for (const card of cards) {
      const email = (card.primaryEmail || "").toLowerCase();
      if (card.isMailList || !email || seen.has(email)) {
        continue;
      }
      seen.add(email);
      const name = card.displayName || [card.firstName, card.lastName].filter(Boolean).join(" ");
      list.push({ name, email, key: fold(`${name} ${email}`) });
    }
  }
  peopleCache = { at: Date.now(), list };
  return list;
}

// Address book contacts and everyone in the mail (most mail first); names
// with a word starting with what was typed come first.
function peopleMatching(text) {
  const words = fold(text).split(/\s+/).filter(Boolean);
  if (!words.length) {
    return [];
  }
  const seen = new Set();
  const candidates = [];
  for (const person of [...(searchState.correspondents || []), ...people()]) {
    if (!seen.has(person.email) && words.every((w) => person.key.includes(w))) {
      seen.add(person.email);
      candidates.push(person);
    }
  }
  const rank = (p) => (fold(p.name).split(/\s+/).some((part) => part.startsWith(words[0])) ? 0 : 1);
  return candidates.sort((a, b) => rank(a) - rank(b)).slice(0, 5);
}

/* ------------------------------------------------------------------ */
/* Results in the mail tab                                             */
/* ------------------------------------------------------------------ */

function syntheticViewOf(hdrs) {
  return {
    viewType: "global",
    customColumns: [],
    defaultSort: [[Ci.nsMsgViewSortType.byDate, Ci.nsMsgViewSortOrder.descending]],
    search(listener, completion) {
      listener.onNewSearch();
      for (const hdr of hdrs) {
        listener.onSearchHit(hdr, hdr.folder);
      }
      listener.onSearchDone(Cr.NS_OK);
      completion?.();
    },
    abortSearch() {},
    getMsgHdrForMessageID(messageId) {
      return hdrs.find((hdr) => hdr.messageId == messageId) || null;
    },
  };
}

function folderName(folder) {
  return folder.localizedName || folder.name;
}

// The folder the mail tab shows (or had selected, under a search).
function shownFolder(about3Pane) {
  const view = about3Pane.gViewWrapper;
  if (view && !view.isSynthetic && view.displayedFolder) {
    return view.displayedFolder;
  }
  const uri = about3Pane.folderTree?.selectedRow?.uri;
  return uri ? MailServices.folderLookup.getFolderForURL(uri) : null;
}

function about3PaneOf(win, { switchTab = true } = {}) {
  const tabmail = win.document.getElementById("tabmail");
  if (!tabmail) {
    return null;
  }
  const isMail = (tab) => tab?.mode?.name == "mail3PaneTab";
  let tab = tabmail.currentTabInfo;
  if (!isMail(tab)) {
    tab = tabmail.tabInfo.find(isMail);
    if (tab && switchTab) {
      tabmail.switchToTab(tab);
    }
  }
  return tab?.chromeBrowser?.contentWindow || null;
}

function barCss() {
  return `
#${BAR_ID} {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: 6px 10px;
  padding: 6px 10px;
  border-bottom: 1px solid var(--splitter-color, ThreeDShadow);
  background: var(--layout-background-1, Canvas);
  color: var(--layout-color-1, CanvasText);
  font-size: 0.95em;
}
#${BAR_ID} .ol-s-title { flex: 1 1 100%; font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
#${BAR_ID} .ol-s-group { display: inline-flex; gap: 4px; flex-wrap: wrap; }
#${BAR_ID} button {
  appearance: none;
  padding: 2px 9px;
  border: 1px solid var(--splitter-color, ThreeDShadow);
  border-radius: 12px;
  background: transparent;
  color: inherit;
  font: inherit;
  cursor: pointer;
}
#${BAR_ID} button:hover { background: var(--layout-background-2, ButtonFace); }
#${BAR_ID} button[aria-pressed="true"] {
  border-color: var(--selected-item-color, AccentColor);
  background: color-mix(in srgb, var(--selected-item-color, AccentColor) 22%, transparent);
}
#${BAR_ID} .ol-s-exit { margin-inline-start: auto; }
/* The folder's name and count are not what the list shows now. */
#threadPaneHeaderBar:has(+ #${BAR_ID}) :is(#threadPaneFolderName, #threadPaneFolderCountContainer) { display: none; }
`;
}

function ensureStyle(doc, id, css) {
  if (!doc.getElementById(id)) {
    const style = doc.createElement("style");
    style.id = id;
    style.textContent = css;
    doc.head.appendChild(style);
  }
}

function renderBar(about3Pane) {
  const mode = searchState.modes.get(about3Pane);
  const doc = about3Pane.document;
  let bar = doc.getElementById(BAR_ID);
  if (!mode) {
    bar?.remove();
    return;
  }
  ensureStyle(doc, STYLE_ID, barCss());
  if (!bar) {
    bar = doc.createElement("div");
    bar.id = BAR_ID;
    const header = doc.getElementById("threadPaneHeaderBar");
    if (header) {
      header.after(bar);
    } else {
      doc.getElementById("threadPane")?.prepend(bar);
    }
  }
  const el = (tag, props, ...children) => {
    const node = doc.createElement(tag);
    Object.assign(node, props);
    node.append(...children);
    return node;
  };
  const button = (text, pressed, onclick, title = "", className = "") => {
    const node = el("button", { type: "button", textContent: text, title, onclick, className });
    if (pressed != null) {
      node.setAttribute("aria-pressed", String(pressed));
    }
    return node;
  };
  const count = mode.busy ? "Searching…" : `${mode.count}${mode.count >= MAX_RESULTS ? "+" : ""} result${mode.count == 1 ? "" : "s"}`;
  const where = mode.scope == "current" && mode.prevFolder ? ` in ${folderName(mode.prevFolder)}` : " in all folders";
  bar.replaceChildren(
    el("span", { className: "ol-s-title", textContent: `🔍 ${count} for “${mode.text}”${mode.busy ? "" : where}` }),
    el(
      "span",
      { className: "ol-s-group" },
      button("All folders", mode.scope == "all", () => rerun(about3Pane, { scope: "all" })),
      mode.prevFolder
        ? button(`Current folder (${folderName(mode.prevFolder)})`, mode.scope == "current", () =>
            rerun(about3Pane, { scope: "current" })
          )
        : ""
    ),
    el(
      "span",
      { className: "ol-s-group" },
      ...FILTERS.map(([term, label]) =>
        button(label, mode.filters.has(term), () => {
          const filters = new Set(mode.filters);
          if (filters.has(term)) {
            filters.delete(term);
          } else {
            filters.add(term);
          }
          rerun(about3Pane, { filters });
        })
      )
    ),
    button("✕ Exit search", null, () => exitSearch(about3Pane), "Back to the folder (Esc)", "ol-s-exit")
  );
}

// Opening a folder (or the selected one again) ends the search.
function watchFolders(about3Pane, mode) {
  const onFolder = () => {
    if (!mode.entering && searchState.modes.get(about3Pane) == mode) {
      endMode(about3Pane);
    }
  };
  const active = () => searchState.modes.get(about3Pane) == mode && !mode.entering;
  // Another folder: Thunderbird ignores the choice under a synthetic view
  // (folderPane._onSelect bails out), so leave the search and choose again.
  const onTreeSelect = () => {
    if (active()) {
      endMode(about3Pane);
      reloadSelectedFolder(about3Pane);
    }
  };
  // The folder that is still selected: no "select" event at all.
  const onTreeClick = (event) => {
    const row = event.target.closest?.("li");
    if (row && active() && about3Pane.folderTree?.selectedRow == row) {
      exitSearch(about3Pane);
    }
  };
  const tree = about3Pane.document.getElementById("folderTree");
  about3Pane.addEventListener("folderURIChanged", onFolder);
  tree?.addEventListener("select", onTreeSelect);
  tree?.addEventListener("click", onTreeClick);
  mode.unwatch = () => {
    about3Pane.removeEventListener("folderURIChanged", onFolder);
    tree?.removeEventListener("select", onTreeSelect);
    tree?.removeEventListener("click", onTreeClick);
  };
}

// Leaves the synthetic view for the folder selected in the folder pane.
function reloadSelectedFolder(about3Pane) {
  about3Pane.gViewWrapper?.close();
  about3Pane.gViewWrapper = null;
  about3Pane.folderTree.dispatchEvent(new about3Pane.CustomEvent("select"));
}

function endMode(about3Pane) {
  const mode = searchState.modes.get(about3Pane);
  if (!mode) {
    return;
  }
  mode.unwatch?.();
  searchState.modes.delete(about3Pane);
  renderBar(about3Pane);
  clearSearchBox(about3Pane.browsingContext?.topChromeWindow || about3Pane.top);
}

function exitSearch(about3Pane) {
  const mode = searchState.modes.get(about3Pane);
  if (!mode) {
    return;
  }
  const { prevFolder, prevMessage } = mode;
  endMode(about3Pane);
  if (prevFolder) {
    // Thunderbird does not leave a synthetic view when a folder is selected
    // (folderPane._onSelect bails out), so close it first.
    about3Pane.gViewWrapper?.close();
    about3Pane.gViewWrapper = null;
    about3Pane.restoreState({ folderURI: prevFolder.URI, folderPaneVisible: !mode.folderPaneCollapsed });
    if (!about3Pane.gViewWrapper) {
      // The folder was still the selected row, so displayFolder() changed
      // nothing: load it as Thunderbird does at startup.
      reloadSelectedFolder(about3Pane);
    }
    if (prevMessage) {
      const index = about3Pane.gDBView?.findIndexOfMsgHdr(prevMessage, true);
      if (index >= 0) {
        about3Pane.threadTree.selectedIndex = index;
      }
    }
  }
}

function showHits(about3Pane, mode, hits, selectKey) {
  const hdrs = hits.map(hdrOf).filter(Boolean);
  mode.count = hdrs.length;
  mode.busy = false;
  mode.entering = true;
  try {
    about3Pane.restoreState({
      syntheticView: syntheticViewOf(hdrs),
      folderPaneVisible: !mode.folderPaneCollapsed,
      title: `Search: ${mode.text}`,
    });
  } finally {
    mode.entering = false;
  }
  if (selectKey) {
    const hdr = hdrs.find((h) => `${h.folder.URI}#${h.messageKey}` == selectKey);
    const index = hdr ? about3Pane.gDBView.findIndexOfMsgHdr(hdr, true) : -1;
    if (index >= 0) {
      about3Pane.threadTree.selectedIndex = index;
    }
  }
  renderBar(about3Pane);
}

async function runSearch(win, text, { selectKey = null } = {}) {
  const about3Pane = about3PaneOf(win);
  if (!about3Pane || !about3Pane.restoreState) {
    return false;
  }
  let mode = searchState.modes.get(about3Pane);
  if (!mode) {
    mode = {
      text,
      scope: "all",
      filters: new Set(),
      prevFolder: shownFolder(about3Pane),
      prevMessage: about3Pane.gViewWrapper?.isSynthetic ? null : about3Pane.gDBView?.hdrForFirstSelectedMessage || null,
      folderPaneCollapsed: Boolean(about3Pane.paneLayout?.folderPaneSplitter?.isCollapsed),
      count: 0,
      busy: true,
    };
    searchState.modes.set(about3Pane, mode);
    watchFolders(about3Pane, mode);
  }
  mode.text = text;
  rememberSearch(text);
  return searchInMode(about3Pane, mode, selectKey);
}

async function searchInMode(about3Pane, mode, selectKey = null) {
  mode.busy = true;
  mode.seq = (mode.seq || 0) + 1;
  const seq = mode.seq;
  renderBar(about3Pane);
  try {
    const query = `${mode.text} ${[...mode.filters].join(" ")}`;
    const hits = await search(query, { folder: mode.scope == "current" ? mode.prevFolder : null });
    if (searchState.modes.get(about3Pane) != mode || seq != mode.seq) {
      return true; // left or searched again meanwhile
    }
    showHits(about3Pane, mode, hits, selectKey);
  } catch (error) {
    note("search", error);
    mode.busy = false;
    renderBar(about3Pane);
  }
  return true;
}

function rerun(about3Pane, changes) {
  const mode = searchState.modes.get(about3Pane);
  if (mode) {
    Object.assign(mode, changes);
    searchInMode(about3Pane, mode);
  }
}

/* ------------------------------------------------------------------ */
/* The search box and its suggestions                                  */
/* ------------------------------------------------------------------ */

function popupCss() {
  return `
#${POPUP_ID} {
  position: fixed;
  z-index: 1000;
  max-height: 70vh;
  overflow-y: auto;
  padding: 6px 0;
  border: 1px solid var(--arrowpanel-border-color, ThreeDShadow);
  border-radius: 8px;
  background: var(--arrowpanel-background, Canvas);
  color: var(--arrowpanel-color, CanvasText);
  box-shadow: 0 6px 18px rgba(0, 0, 0, 0.25);
  font: menu;
  font-size: 1em;
}
#${POPUP_ID} .ol-sg-head { padding: 6px 14px 2px; font-size: 0.85em; font-weight: 600; opacity: 0.7; }
#${POPUP_ID} .ol-sg-row { display: flex; align-items: center; gap: 10px; padding: 6px 14px; cursor: pointer; }
#${POPUP_ID} .ol-sg-row[aria-selected="true"], #${POPUP_ID} .ol-sg-row:hover {
  background: var(--selected-item-color, Highlight);
  color: var(--selected-item-text-color, HighlightText);
}
#${POPUP_ID} .ol-sg-icon {
  flex: 0 0 26px; height: 26px; border-radius: 50%;
  display: inline-flex; align-items: center; justify-content: center;
  background: color-mix(in srgb, currentColor 14%, transparent); font-size: 0.8em; font-weight: 600;
}
#${POPUP_ID} .ol-sg-main { flex: 1; min-width: 0; }
#${POPUP_ID} .ol-sg-line1, #${POPUP_ID} .ol-sg-line2 { overflow: hidden; white-space: nowrap; text-overflow: ellipsis; }
#${POPUP_ID} .ol-sg-line2 { font-size: 0.85em; opacity: 0.75; }
#${POPUP_ID} .ol-sg-date { font-size: 0.85em; opacity: 0.75; white-space: nowrap; }
#${POPUP_ID} .ol-sg-tips { margin-top: 4px; padding: 6px 14px 2px; border-top: 1px solid var(--arrowpanel-border-color, ThreeDShadow); font-size: 0.85em; }
#${POPUP_ID} .ol-sg-tips-head { margin-bottom: 4px; font-weight: 600; opacity: 0.7; }
#${POPUP_ID} .ol-sg-tips ul { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 3px 18px; margin: 0; padding: 0; list-style: none; }
#${POPUP_ID} .ol-sg-tips li { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
#${POPUP_ID} .ol-sg-tips li::before { content: "•"; margin-inline-end: 6px; opacity: 0.6; }
#${POPUP_ID} .ol-sg-tips code { padding: 0 4px; border-radius: 3px; background: color-mix(in srgb, currentColor 12%, transparent); font-family: monospace; cursor: pointer; }
#${POPUP_ID} .ol-sg-tips code:hover { background: color-mix(in srgb, currentColor 24%, transparent); }
#${POPUP_ID} .ol-sg-tips .ol-sg-what { opacity: 0.7; }
`;
}

function inputOf(bar) {
  return bar.shadowRoot?.querySelector("input") || null;
}

function clearSearchBox(win) {
  const bar = win?.document?.querySelector("global-search-bar");
  const input = bar && inputOf(bar);
  if (input) {
    input.value = "";
  }
}

const DATE_FORMAT = new Intl.DateTimeFormat(undefined, { day: "numeric", month: "short", year: "numeric" });

function initials(text) {
  const words = String(text || "?").replace(/\(.*?\)/g, "").trim().split(/\s+/);
  return ((words[0]?.[0] || "") + (words.length > 1 ? words[words.length - 1][0] : "")).toUpperCase() || "?";
}

function closePopup(win) {
  const ui = searchState.windows.get(win);
  if (ui) {
    ui.popup?.remove();
    ui.popup = null;
    ui.items = [];
    ui.active = -1;
  }
}

function renderPopup(win, bar, sections) {
  const ui = searchState.windows.get(win);
  const doc = win.document;
  ensureStyle(doc, STYLE_ID, popupCss());
  if (!ui.popup) {
    ui.popup = doc.createElement("div");
    ui.popup.id = POPUP_ID;
    ui.popup.setAttribute("role", "listbox");
    // Keep the focus in the search box.
    ui.popup.addEventListener("mousedown", (event) => event.preventDefault());
    doc.documentElement.appendChild(ui.popup);
  }
  const rect = bar.getBoundingClientRect();
  ui.popup.style.left = `${Math.round(rect.left)}px`;
  ui.popup.style.top = `${Math.round(rect.bottom + 4)}px`;
  ui.popup.style.width = `${Math.round(Math.max(rect.width, 480))}px`;
  ui.items = [];
  ui.active = -1;
  const el = (tag, className, text) => {
    const node = doc.createElement(tag);
    if (className) {
      node.className = className;
    }
    if (text != null) {
      node.textContent = text;
    }
    return node;
  };
  const nodes = [];
  for (const section of sections) {
    if (!section.rows.length) {
      continue;
    }
    if (section.title) {
      nodes.push(el("div", "ol-sg-head", section.title));
    }
    for (const row of section.rows) {
      const node = el("div", "ol-sg-row");
      node.setAttribute("role", "option");
      node.append(el("span", "ol-sg-icon", row.icon));
      const main = el("span", "ol-sg-main");
      main.append(el("div", "ol-sg-line1", row.line1));
      if (row.line2) {
        main.append(el("div", "ol-sg-line2", row.line2));
      }
      node.append(main);
      if (row.date) {
        node.append(el("span", "ol-sg-date", row.date));
      }
      const index = ui.items.length;
      node.addEventListener("click", () => chooseSuggestion(win, bar, index));
      node.addEventListener("mouseenter", () => setActive(win, index));
      ui.items.push({ node, action: row.action });
      nodes.push(node);
    }
  }
  nodes.push(searchTips(doc, bar));
  ui.popup.replaceChildren(...nodes);
}

// Outlook's search keywords; a click puts one in the search box.
const SEARCH_TIPS = [
  ["from:", "name", "sender"],
  ["to:", "name", "recipient (To or Cc)"],
  ["subject:", "word", "in the subject"],
  ["body:", "word", "in the text"],
  ["has:attachment", "", "with attachments"],
  ["is:unread", "", "unread mail"],
  ["is:pinned", "", "pinned mail"],
  ["after:", "2026-10-01", "received after"],
  ["before:", "2026-10-01", "received before"],
  ['"', "exact phrase", "words together"],
];

function searchTips(doc, bar) {
  const box = doc.createElement("div");
  box.className = "ol-sg-tips";
  const head = doc.createElement("div");
  head.className = "ol-sg-tips-head";
  head.textContent = "Search tips";
  const list = doc.createElement("ul");
  for (const [keyword, example, meaning] of SEARCH_TIPS) {
    const item = doc.createElement("li");
    const code = doc.createElement("code");
    code.textContent = keyword == '"' ? `"${example}"` : keyword + example;
    code.title = "Add to the search";
    code.addEventListener("click", () => {
      const input = inputOf(bar);
      if (input) {
        const insert = keyword == '"' ? '""' : keyword;
        input.value = `${input.value.trimEnd()}${input.value.trim() ? " " : ""}${insert}`;
        input.focus();
        const caret = keyword == '"' ? input.value.length - 1 : input.value.length;
        input.setSelectionRange(caret, caret);
      }
    });
    const what = doc.createElement("span");
    what.className = "ol-sg-what";
    what.textContent = ` ${meaning}`;
    item.append(code, what);
    list.append(item);
  }
  box.append(head, list);
  return box;
}

function setActive(win, index) {
  const ui = searchState.windows.get(win);
  ui.active = index;
  ui.items.forEach((item, i) => item.node.setAttribute("aria-selected", String(i == index)));
  ui.items[index]?.node.scrollIntoView({ block: "nearest" });
}

function chooseSuggestion(win, bar, index) {
  const ui = searchState.windows.get(win);
  const item = ui.items[index];
  closePopup(win);
  item?.action(bar);
}

async function suggest(win, bar, text) {
  const ui = searchState.windows.get(win);
  const seq = (ui.suggestSeq = (ui.suggestSeq || 0) + 1);
  const value = text.trim();
  if (!value) {
    const recent = recentSearches();
    renderPopup(win, bar, [
      {
        title: recent.length ? "Recent searches" : "",
        rows: recent.map((q) => ({ icon: "🕘", line1: q, action: (b) => setBoxAndSearch(win, b, q) })),
      },
    ]);
    return;
  }
  const sections = [
    {
      title: "",
      rows: [{ icon: "🔍", line1: `Search for “${value}”`, line2: "Enter", action: (b) => setBoxAndSearch(win, b, value) }],
    },
  ];
  renderPopup(win, bar, sections);
  // People and messages: from the mail read a moment ago (or now, the first time).
  const corpus = await getCorpus(SUGGEST_TTL_MS).catch(() => null);
  if (ui.suggestSeq != seq || !ui.popup) {
    return;
  }
  const lastWord = value.split(/\s+/).pop().replace(/^(from|to|cc|participants|with):/i, "");
  sections.push({
    title: "People",
    rows: peopleMatching(lastWord).map((p) => ({
      icon: initials(p.name || p.email),
      line1: p.name || p.email,
      line2: `${p.email} · all mail with this person`,
      action: (b) => setBoxAndSearch(win, b, `participants:${p.email}`),
    })),
  });
  const query = parseQuery(value);
  const top = [];
  for (const entry of corpus || []) {
    if (matches(entry, query, null, null)) {
      top.push(entry);
      if (top.length >= 5) {
        break;
      }
    }
  }
  sections.push({
    title: "Messages",
    rows: top.map((entry) => ({
      icon: "✉",
      line1: entry.subject || "(no subject)",
      line2: entry.author,
      date: DATE_FORMAT.format(new Date(entry.date * 1000)),
      action: (b) => setBoxAndSearch(win, b, value, `${entry.folder.URI}#${entry.key}`),
    })),
  });
  renderPopup(win, bar, sections);
}

function setBoxAndSearch(win, bar, text, selectKey = null) {
  const input = inputOf(bar);
  if (input) {
    input.value = text;
  }
  runSearch(win, text, { selectKey }).catch((error) => note("search", error));
}

function attachWindow(win) {
  if (searchState.windows.has(win)) {
    return;
  }
  const isBar = (event) => event.target?.localName == "global-search-bar";
  const onAutocomplete = (event) => {
    if (!searchState.enabled || !isBar(event)) {
      return;
    }
    event.stopPropagation(); // not Gloda's suggestions
    suggest(win, event.target, String(event.detail || "")).catch((error) => note("suggest", error));
  };
  const onSearch = (event) => {
    if (!searchState.enabled || !isBar(event)) {
      return;
    }
    event.stopPropagation(); // not Gloda's results tab
    event.preventDefault(); // keep the words in the box
    closePopup(win);
    runSearch(win, String(event.detail || "").trim()).catch((error) => note("search", error));
  };
  const onKeydown = (event) => {
    if (!searchState.enabled || !isBar(event)) {
      return;
    }
    const ui = searchState.windows.get(win);
    if (ui.popup && ["ArrowDown", "ArrowUp"].includes(event.key)) {
      event.preventDefault();
      event.stopPropagation();
      const count = ui.items.length;
      if (count) {
        const down = event.key == "ArrowDown";
        setActive(win, down ? (ui.active + 1) % count : ui.active <= 0 ? count - 1 : ui.active - 1);
      }
    } else if (ui.popup && event.key == "Enter" && ui.active >= 0) {
      event.preventDefault();
      event.stopPropagation();
      chooseSuggestion(win, event.target, ui.active);
    } else if (event.key == "Escape") {
      if (ui.popup) {
        closePopup(win);
        event.preventDefault();
        event.stopPropagation();
      } else {
        const about3Pane = about3PaneOf(win, { switchTab: false });
        if (about3Pane && searchState.modes.has(about3Pane)) {
          exitSearch(about3Pane);
        }
      }
    }
  };
  const onFocus = (event) => {
    if (searchState.enabled && isBar(event) && !inputOf(event.target)?.value) {
      suggest(win, event.target, "").catch((error) => note("suggest", error));
      getCorpus(SUGGEST_TTL_MS).catch(() => {}); // ready for the first keys
    }
  };
  const onBlur = (event) => {
    if (isBar(event)) {
      win.setTimeout(() => closePopup(win), 150);
    }
  };
  win.addEventListener("autocomplete", onAutocomplete, true);
  win.addEventListener("search", onSearch, true);
  win.addEventListener("keydown", onKeydown, true);
  win.addEventListener("focus", onFocus, true);
  win.addEventListener("blur", onBlur, true);
  searchState.windows.set(win, {
    popup: null,
    items: [],
    active: -1,
    cleanup() {
      win.removeEventListener("autocomplete", onAutocomplete, true);
      win.removeEventListener("search", onSearch, true);
      win.removeEventListener("keydown", onKeydown, true);
      win.removeEventListener("focus", onFocus, true);
      win.removeEventListener("blur", onBlur, true);
      closePopup(win);
      win.document.getElementById(STYLE_ID)?.remove();
    },
  });
  win.addEventListener("unload", () => searchState.windows.delete(win), { once: true });
}

function detachAll() {
  for (const about3Pane of [...searchState.modes.keys()]) {
    try {
      exitSearch(about3Pane);
      about3Pane.document.getElementById(STYLE_ID)?.remove();
    } catch (error) {
      // the window is closing
    }
  }
  for (const ui of searchState.windows.values()) {
    try {
      ui.cleanup();
    } catch (error) {
      // the window is closing
    }
  }
  searchState.windows.clear();
  searchState.corpus = null;
}

this.outlookSearch = class extends ExtensionCommon.ExtensionAPI {
  onShutdown(isAppShutdown) {
    if (isAppShutdown) {
      return;
    }
    if (searchState.enabled) {
      ExtensionSupport.unregisterWindowListener(SEARCH_LISTENER);
      searchState.enabled = false;
    }
    detachAll();
    Services.obs.notifyObservers(null, "startupcache-invalidate");
  }

  getAPI() {
    return {
      outlookSearch: {
        // The user's choice is a Thunderbird preference kept here, in the
        // privileged code: it changes only when Preferences passes
        // options.enabled; otherwise the saved choice (default on) is applied.
        // Settings sync does not cover "extensions." preferences.
        async configure(options) {
          if (typeof options.enabled === "boolean") {
            Services.prefs.setBoolPref(SEARCH_ENABLED_PREF, options.enabled);
          }
          const enabled = Services.prefs.getBoolPref(SEARCH_ENABLED_PREF, true);
          if (enabled && !searchState.enabled) {
            searchState.enabled = true;
            ExtensionSupport.registerWindowListener(SEARCH_LISTENER, {
              chromeURLs: ["chrome://messenger/content/messenger.xhtml"],
              onLoadWindow: attachWindow,
            });
          } else if (!enabled && searchState.enabled) {
            ExtensionSupport.unregisterWindowListener(SEARCH_LISTENER);
            searchState.enabled = false;
            detachAll();
          }
          return { enabled: searchState.enabled };
        },
        async status() {
          return {
            enabled: searchState.enabled,
            messagesRead: searchState.corpus ? searchState.corpus.length : 0,
            lastError: searchState.lastError,
          };
        },
      },
    };
  }
};
