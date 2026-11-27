"use strict";
// Run: node tests/own-replies.test.js
// experiments/outlookPin: your replies (in Sent Items) count in their
// conversation, as in Outlook. A conversation of the Inbox that you answered
// is sorted ("📌 Pinned first", newest first) by your reply and its card shows
// your reply (sender and time); the Sent folder itself is left as it is, and
// a new mail in Sent Items makes the Inbox list sort again.
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");

const code = fs.readFileSync(path.join(__dirname, "../outlook-mode/experiments/outlookPin/implementation.js"), "utf8");
const SENT = 0x200;
const sec = (iso) => Math.floor(Date.parse(iso) / 1000);

function hdr(key, folder, messageId, refs, iso, author) {
  return {
    messageKey: key, folder, messageId, mime2DecodedAuthor: author,
    dateInSeconds: sec(iso), date: sec(iso) * 1e6,
    numReferences: refs.length, getStringReference: (i) => refs[i],
    getStringProperty: () => "",
  };
}
function folder(name, flags, server) {
  const f = {
    URI: `graph://me/${name}`, name, flags, server, msgs: [],
    isSpecialFolder: (flag) => Boolean(f.flags & flag),
    getTotalMessages: () => f.msgs.length,
    msgDatabase: {
      enumerateMessages: () => f.msgs.values(),
      // One thread per root, children by thread id.
      getThreadContainingMsgHdr: (h) => {
        const members = f.msgs.filter((m) => (m.thread || m) === (h.thread || h));
        const root = members[0];
        return {
          numChildren: members.length,
          getChildHdrAt: (i) => members[i],
          getRootHdr: () => root,
          newestMsgDate: Math.max(...members.map((m) => m.dateInSeconds)),
        };
      },
    },
  };
  return f;
}
const server = { key: "server2" };
const inbox = folder("Inbox", 0x1000, server);
const sent = folder("Sent Items", SENT, server);
server.rootFolder = { getFoldersWithFlags: (flag) => (flag === SENT ? [sent] : []) };

// Inbox: conversation A (two mails, newest 20/11 11:49), conversation B (a
// single mail, 21/11), conversation C (24/11).
const a1 = hdr(1, inbox, "a1@x", [], "2026-11-20T11:38:00Z", "Ann <a@x>");
const a2 = hdr(2, inbox, "a2@x", ["a1@x"], "2026-11-20T11:49:00Z", "Bob <b@x>");
a2.thread = a1;
const b1 = hdr(3, inbox, "b1@x", [], "2026-11-21T09:00:00Z", "Carol <c@x>");
const c1 = hdr(4, inbox, "c1@x", [], "2026-11-24T09:00:00Z", "Dan <d@x>");
inbox.msgs.push(a1, a2, b1, c1);
// Sent Items: an old reply in A, and a mail answering nothing in the Inbox.
sent.msgs.push(hdr(10, sent, "s0@me", ["a1@x"], "2026-11-20T11:45:00Z", "Tùng <me@x>"));
sent.msgs.push(hdr(11, sent, "s1@me", ["zzz@y"], "2026-11-26T09:00:00Z", "Tùng <me@x>"));

// A mail tab showing the Inbox, sorted "Pinned first".
const resorts = [];
const inboxPane = {
  closed: false,
  gFolder: inbox,
  gViewWrapper: { primarySortType: 22, isVirtual: false, dbView: {} },
  gDBView: { curCustomColumn: "outlookModePinnedCol" },
  setTimeout: (fn) => {
    resorts.push(fn);
    return resorts.length;
  },
};
const mailWindow = { gTabmail: { tabInfo: [{ mode: { name: "mail3PaneTab" }, chromeBrowser: { contentWindow: inboxPane } }] } };

const ctx = {
  ChromeUtils: {
    importESModule: (url) =>
      url.endsWith("MailServices.sys.mjs")
        ? {
            MailServices: {
              headerParser: {
                parseDecodedHeader: (text) => {
                  const m = /^\s*"?([^"<]*?)"?\s*<([^>]+)>/.exec(text);
                  return m ? [{ name: m[1], email: m[2] }] : [{ name: "", email: text }];
                },
              },
            },
          }
        : new Proxy({}, { get: () => ({}) }),
    generateQI: () => () => {},
  },
  Ci: { nsMsgFolderFlags: { SentMail: SENT }, nsMsgViewSortType: { byCustom: 22 } },
  Services: { wm: { getEnumerator: () => [mailWindow] } },
  ExtensionCommon: { ExtensionAPI: class {} },
  console, Date, Math, JSON, Promise, Map, Set, Object, Array, String, Number, Boolean, Error, Intl,
};
vm.createContext(ctx);
vm.runInContext(code, ctx);
vm.runInContext("this.api = { sortDate, newestOwnReply, showNewestInCard, resortShowing, listState };", ctx);
const { api } = ctx;
const realNow = Date.now;
let clock = realNow();
Date.now = () => clock;

// 1. An older reply of yours does not move A; a mail answering nothing changes nothing.
assert.equal(api.sortDate(a1), sec("2026-11-20T11:49:00Z"), "A: by its newest mail (your reply was earlier)");
assert.equal(api.sortDate(b1), sec("2026-11-21T09:00:00Z"));

// 2. You answer B (the reply reaches Sent Items through the sync): B moves
// above C, its card shows your reply.
sent.msgs.push(hdr(12, sent, "s2@me", ["b1@x"], "2026-11-26T10:00:00Z", "Tùng <me@x>"));
api.resortShowing(sent); // the msgAdded notification
assert.equal(resorts.length, 1, "the Inbox list is sorted again");
assert.equal(api.sortDate(b1), sec("2026-11-26T10:00:00Z"), "B: by your reply");
assert.ok(api.sortDate(b1) > api.sortDate(c1), "B above C");
assert.deepEqual(api.newestOwnReply([b1]), { date: sec("2026-11-26T10:00:00Z"), author: "Tùng <me@x>" });

api.listState.newestInCard = true;
const card = (index, rows) => ({
  index,
  senderLine: { textContent: "first", title: "" },
  dateLine: { textContent: "first date" },
  view: rows,
});
const viewOf = (rows, open = false) => ({
  rowCount: rows.length,
  isContainer: (i) => rows[i] === a1,
  isContainerOpen: () => open,
  getMsgHdrAt: (i) => rows[i],
});
const about3Pane = { gDBView: viewOf([b1, c1, a1]) };
const cardB = card(0);
api.showNewestInCard(about3Pane, cardB);
assert.equal(cardB.senderLine.textContent, "Tùng", "a single mail you answered: your reply on the card");
assert.equal(cardB.senderLine.title, "Tùng <me@x>");
assert.notEqual(cardB.dateLine.textContent, "first date");

const cardC = card(1);
api.showNewestInCard(about3Pane, cardC);
assert.equal(cardC.senderLine.textContent, "first", "not answered: as Thunderbird shows it");

const cardA = card(2);
api.showNewestInCard(about3Pane, cardA);
assert.equal(cardA.senderLine.textContent, "Bob", "A: its newest mail (Bob), your reply was earlier");

// 3. Your new reply in A wins over Bob's mail.
sent.msgs.push(hdr(13, sent, "s3@me", ["a1@x", "a2@x"], "2026-11-27T08:00:00Z", "Tùng <me@x>"));
api.resortShowing(sent);
assert.equal(api.sortDate(a1), sec("2026-11-27T08:00:00Z"));
const cardA2 = card(2);
api.showNewestInCard(about3Pane, cardA2);
assert.equal(cardA2.senderLine.textContent, "Tùng");

// 4. An open thread: each card is its own mail.
const openPane = { gDBView: viewOf([a1], true) };
const cardOpen = card(0);
api.showNewestInCard(openPane, cardOpen);
assert.equal(cardOpen.senderLine.textContent, "first");

// 5. In the Sent folder itself nothing changes.
const sentRoot = sent.msgs[0];
assert.equal(api.sortDate(sentRoot), sentRoot.dateInSeconds);

// 6. Without the notification, the index notices a new sent mail by its count
// (looked at again after SENT_INDEX_CHECK_MS).
sent.msgs.push(hdr(14, sent, "s4@me", ["c1@x"], "2026-11-27T09:00:00Z", "Tùng <me@x>"));
assert.equal(api.sortDate(c1), sec("2026-11-24T09:00:00Z"), "cached for a moment");
clock += 2500;
assert.equal(api.sortDate(c1), sec("2026-11-27T09:00:00Z"), "then seen");

Date.now = realNow;
console.log("own-replies: all checks passed");
