"use strict";
// Run: node tests/handlers.test.js
// Loads common.js with undo-send.js, mention.js, schedule-send.js and pin.js
// against a fake Thunderbird, and checks the checks of the message
// handlers: who may call them, and what they accept.
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");

const BG = path.join(__dirname, "../outlook-mode/background");
const BASE = "moz-extension://own-uuid/";
const OWN_ID = "outlook-mention@tung.local";

const sent = []; // compose.sendMessage calls
const composeTo = { 7: ["Ann <ann@x.com>"] }; // tab id -> To of the compose window
const stored = { scheduled: [{ id: "s1", headerMessageId: "d1@x", sendAt: new Date(Date.now() + 3600e3).toISOString(), status: "pending" }] };
const queried = []; // messages.query calls

const messenger = {
  runtime: { id: OWN_ID, getURL: (p) => BASE + String(p).replace(/^\//, ""), onInstalled: { addListener() {} } },
  storage: {
    local: {
      async get(key) {
        const keys = typeof key === "string" ? [key] : Array.isArray(key) ? key : Object.keys(key || stored);
        return Object.fromEntries(keys.filter((k) => k in stored).map((k) => [k, structuredClone(stored[k])]));
      },
      async set(items) {
        Object.assign(stored, structuredClone(items));
      },
    },
    onChanged: { addListener() {} },
  },
  compose: {
    onBeforeSend: { addListener() {} },
    async getComposeDetails(tabId) {
      return { to: [...(composeTo[tabId] || [])], cc: [], bcc: [] };
    },
    async setComposeDetails(tabId, details) {
      composeTo[tabId] = details.to;
    },
    async sendMessage(tabId) {
      sent.push(tabId);
      return {};
    },
  },
  tabs: { onRemoved: { addListener() {} }, sendMessage: async () => ({ ok: true }) },
  alarms: { create() {}, clear: async () => true, onAlarm: { addListener() {} } },
  messages: {
    onNewMailReceived: { addListener() {} },
    async query(q) {
      queried.push(q.headerMessageId);
      return { messages: [], id: null };
    },
    async getHeaders() {
      return { references: ["<good.id@x> <bad id@x> <evil\u0001@x>"], "in-reply-to": [] };
    },
  },
  mailTabs: { onDisplayedFolderChanged: { addListener() {} }, query: async () => [] },
  outlookPin: { register: async () => {}, onPinButton: { addListener() {} } },
  contacts: {},
  addressBooks: {},
};

const ctx = { messenger, console, setTimeout, clearTimeout, structuredClone, URL, Promise, Date, Math, JSON };
vm.createContext(ctx);
for (const file of ["common.js", "undo-send.js", "mention.js", "schedule-send.js", "pin.js"]) {
  vm.runInContext(fs.readFileSync(path.join(BG, file), "utf8"), ctx, { filename: file });
}
// Address books: one person.
vm.runInContext(
  `getContacts = async () => [{ name: 'Bob "the builder" <b>', email: "bob@x.com", label: "Bob" }];`,
  ctx
);

const compose = (tabId) => ({ id: OWN_ID, url: "about:blank?compose", tab: { id: tabId, type: "messageCompose" } });
const page = (p) => ({ id: OWN_ID, url: BASE + p, tab: { id: 50, type: "content" } });
const handlers = vm.runInContext("handlers", ctx);
const call = (type, message, sender) => Promise.resolve().then(() => handlers[type]({ type, ...message }, sender));

(async () => {
  // 1. Undo Send "send now": only a compose tab with a countdown running.
  assert.deepEqual(await call("undo-send-now", {}, compose(7)), { ok: false }, "no countdown: nothing sent");
  assert.deepEqual(await call("undo-send-now", {}, page("options/options.html")), { ok: false }, "not a compose tab");
  vm.runInContext(`pendingSends.set(7, { timer: 0, endsAt: Date.now() + 5000 });`, ctx);
  assert.deepEqual(await call("undo-send-now", {}, compose(8)), { ok: false }, "another tab's countdown does not count");
  assert.deepEqual(await call("undo-send-now", {}, compose(7)), { ok: true });
  await new Promise((resolve) => setTimeout(resolve, 10));
  assert.deepEqual(sent, [7], "only tab 7, the one counting down, was sent");

  // 2. @mention: only an address book person, with the address book's name;
  //    an address that would add more recipients is refused.
  assert.deepEqual(await call("mention-add", { name: "x", email: "a@x.com>, evil@y.com <b" }, compose(7)), { added: false });
  assert.deepEqual(await call("mention-add", { name: "x", email: "stranger@y.com" }, compose(7)), { added: false }, "not in the address books");
  assert.deepEqual(await call("mention-add", { name: "x", email: "bob@x.com" }, page("options/options.html")), { added: false }, "not a compose tab");
  assert.deepEqual(await call("mention-add", { name: '"><img>', email: "BOB@x.com" }, compose(7)), { added: true });
  assert.deepEqual(composeTo[7], ["Ann <ann@x.com>", '"Bob the builder b" <bob@x.com>'], "address book name, no quotes or brackets");
  assert.deepEqual(await call("mention-add", { email: "bob@x.com" }, compose(7)), { added: false }, "already there");

  // 3. Scheduled mail: each handler only from the page that uses it.
  await assert.rejects(call("schedule-cancel", { id: "s1" }, page("reading-pane/dock.html")), /Not allowed/);
  await assert.rejects(call("schedule-send-now", { id: "s1" }, page("popup/quicksteps.html")), /Not allowed/);
  await assert.rejects(call("schedule-list", {}, page("cleanup/cleanup.html")), /Not allowed/);
  await assert.rejects(call("schedule-list", {}, { id: "other@addon", url: BASE + "options/options.html" }), /Not allowed/, "another add-on");
  assert.equal((await call("schedule-list", {}, page("popup/quicksteps.html"))).length, 1);
  assert.equal((await call("schedule-list", {}, page("options/options.html?x=1#y"))).length, 1, "query and hash ignored");
  assert.deepEqual(await call("schedule-cancel", { id: "nope" }, page("options/options.html")), { ok: false });
  assert.deepEqual(await call("schedule-cancel", { id: "s1" }, page("options/options.html")), { ok: true });
  assert.equal(stored.scheduled.length, 0);

  // 4. Thread root lookup: only well-formed Message-IDs from the headers.
  await ctx.threadRootOf({ id: 1, headerMessageId: "self@x", folder: { id: "f1" } });
  assert.deepEqual(queried, ["good.id@x"], "ids with spaces or control characters are not looked up");

  console.log("all handlers scenarios passed");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
