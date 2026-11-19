"use strict";
// Run: node tests/sent-copies.test.js
// One copy of sent mail on Microsoft Graph accounts:
// 1. background/sent-copies.js against an in-memory Sent Items: the second
//    copies (same Message-ID, created seconds apart; the draft one first) go
//    to Deleted Items, the rest stays;
// 2. experiments/outlookPane: a send through a Graph outgoing server files no
//    copy of its own in a folder of a Graph account (reading pane and compose
//    window), and other sends are left as they are.
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");

const ROOT = path.join(__dirname, "../outlook-mode/");
const code = ["common.js", "graph.js", "sent-copies.js"].map((f) => fs.readFileSync(ROOT + "background/" + f, "utf8")).join("\n;\n");

// ---- fake Graph mailbox ----
const server = { items: new Map(), moves: [], requests: [] };
function addSent(id, messageId, created, isDraft = false) {
  server.items.set(id, { id, internetMessageId: `<${messageId}>`, createdDateTime: created, isDraft, folder: "sentitems" });
}

async function fakeFetch(url, init = {}) {
  const u = new URL(url);
  const urlPath = u.pathname.replace("/v1.0", "");
  const method = init.method || "GET";
  server.requests.push(`${method} ${urlPath}${u.search}`);
  const json = (status, body) => ({ ok: status < 400, status, json: async () => body });
  let m;
  if (urlPath === "/me/mailFolders/sentitems/messages" && method === "GET") {
    assert.equal(u.searchParams.get("$orderby"), "createdDateTime desc");
    const filter = u.searchParams.get("$filter") || "";
    const since = (m = filter.match(/^createdDateTime ge (\S+)$/)) ? Date.parse(m[1]) : 0;
    if (filter && !m) {
      return json(400, { error: { message: `bad filter ${filter}` } });
    }
    const skip = Number(u.searchParams.get("$skip") || 0);
    const all = [...server.items.values()]
      .filter((i) => i.folder === "sentitems" && Date.parse(i.createdDateTime) >= since)
      .sort((a, b) => b.createdDateTime.localeCompare(a.createdDateTime));
    const page = all.slice(skip, skip + 3); // tiny pages, to exercise nextLink
    const next = new URL(url);
    next.searchParams.set("$skip", String(skip + 3));
    return json(200, {
      value: page.map(({ id, internetMessageId, isDraft, createdDateTime }) => ({ id, internetMessageId, isDraft, createdDateTime })),
      "@odata.nextLink": skip + 3 < all.length ? next.toString() : undefined,
    });
  }
  if ((m = urlPath.match(/^\/me\/messages\/([^/]+)\/move$/)) && method === "POST") {
    const item = server.items.get(decodeURIComponent(m[1]));
    assert.ok(item, `move of unknown ${m[1]}`);
    assert.deepEqual(JSON.parse(init.body), { destinationId: "deleteditems" });
    item.folder = "deleteditems";
    server.moves.push(item.id);
    return json(201, { id: item.id });
  }
  throw new Error(`unexpected ${method} ${url}`);
}

function makeBackground() {
  const storage = {};
  const notices = [];
  const setCalls = [];
  const timers = [];
  const messenger = {
    storage: {
      local: {
        async get(keys) {
          const wanted = keys == null ? Object.keys(storage) : Array.isArray(keys) ? keys : [keys];
          return Object.fromEntries(wanted.filter((k) => k in storage).map((k) => [k, structuredClone(storage[k])]));
        },
        async set(obj) {
          Object.assign(storage, structuredClone(obj));
        },
      },
      onChanged: { addListener() {} },
    },
    accounts: { get: async (id) => ({ id, name: id }) },
    alarms: { create() {}, onAlarm: { addListener() {} } },
    runtime: { onInstalled: { addListener() {} } },
    notifications: { create: async (options) => notices.push(options) },
    outlookPin: {
      listGraphAccounts: async () => ["account1"],
      getGraphToken: async () => "token",
    },
    outlookPane: {
      setSingleSentCopy: async (enabled) => {
        setCalls.push(enabled);
        return { enabled, skipped: 0 };
      },
    },
  };
  const ctx = {
    messenger, fetch: fakeFetch, console, URLSearchParams, URL,
    setTimeout: (fn) => timers.push(fn), clearTimeout: () => {},
    Promise, JSON, Date, Math, Object, Array, Map, Set, String, Number, Error, TypeError, structuredClone,
  };
  vm.createContext(ctx);
  vm.runInContext(code + "\n;this.api = { checkSentCopies, secondSentCopies, sentCopiesStatus, configureSentCopies };", ctx);
  return { api: ctx.api, storage, notices, setCalls };
}

async function testBackground() {
  // Two mails sent from Thunderbird before the fix: Exchange's copy and
  // Thunderbird's draft-marked upload, seconds apart.
  addSent("a-server", "a@vinmotion.example", "2026-11-19T10:04:02Z");
  addSent("a-fcc", "a@vinmotion.example", "2026-11-19T10:04:04Z", true);
  addSent("b-fcc", "b@vinmotion.example", "2026-11-19T10:18:40Z", true); // the upload came first this time
  addSent("b-server", "b@vinmotion.example", "2026-11-19T10:18:43Z");
  addSent("c", "c@outlook.example", "2026-11-17T04:53:25Z"); // sent from Outlook: one copy
  addSent("d1", "d@x.example", "2026-11-01T08:00:00Z"); // same id weeks apart: left alone
  addSent("d2", "d@x.example", "2026-11-15T08:00:00Z");
  addSent("e", "", "2026-11-18T08:00:00Z"); // no Message-ID

  const bg = makeBackground();
  await new Promise((resolve) => setImmediate(resolve));
  assert.deepEqual(bg.setCalls, [true], "turned on at startup (on by default)");

  let r = await bg.api.checkSentCopies();
  assert.ok(r.ok, JSON.stringify(r));
  assert.equal(r.moved, 2);
  assert.deepEqual(server.moves.sort(), ["a-fcc", "b-fcc"], "the draft copies go, Exchange's own copies stay");
  for (const id of ["a-server", "b-server", "c", "d1", "d2", "e"]) {
    assert.equal(server.items.get(id).folder, "sentitems", `${id} kept`);
  }
  assert.equal(bg.notices.length, 1, "one notice says what was moved");
  assert.match(bg.notices[0].message, /2 mail\(s\) were twice in Sent Items/);
  assert.ok(bg.storage.sentCopiesChecked.account1, "time of the check kept");
  assert.equal(bg.api.sentCopiesStatus.account1.moved, 2);
  assert.ok(
    server.requests.some((q) => q.includes("%24skip=3")),
    "paging links followed"
  );

  // 2. The next check reads only what was created lately, finds nothing.
  server.requests.length = 0;
  r = await bg.api.checkSentCopies();
  assert.equal(r.moved, 0);
  assert.ok(server.requests[0].includes("createdDateTime+ge+"), server.requests[0]);
  assert.equal(bg.notices.length, 1, "no notice when nothing moved");

  // 3. Two non-draft copies: the older one is kept.
  const now = Date.now();
  addSent("f-new", "f@x.example", new Date(now - 1000).toISOString().replace(/\.\d+Z$/, "Z"));
  addSent("f-old", "f@x.example", new Date(now - 4000).toISOString().replace(/\.\d+Z$/, "Z"));
  r = await bg.api.checkSentCopies();
  assert.equal(r.moved, 1);
  assert.equal(server.items.get("f-new").folder, "deleteditems");
  assert.equal(server.items.get("f-old").folder, "sentitems");

  // 4. Off in Preferences: nothing read, nothing moved.
  bg.storage.settings = { singleSentCopy: false };
  server.requests.length = 0;
  r = await bg.api.checkSentCopies();
  assert.equal(r.ok, false);
  assert.match(r.error, /is off/);
  assert.equal(server.requests.length, 0);
  await bg.api.configureSentCopies();
  assert.deepEqual(bg.setCalls, [true, false]);

  // secondSentCopies on its own: all drafts -> the oldest stays.
  const extra = bg.api.secondSentCopies([
    { id: "x2", internetMessageId: "<x@y>", isDraft: true, createdDateTime: "2026-11-19T10:00:05Z" },
    { id: "x1", internetMessageId: "<x@y>", isDraft: true, createdDateTime: "2026-11-19T10:00:01Z" },
  ]);
  assert.deepEqual(extra.map((i) => i.id), ["x2"]);
}

// ---- the experiment: which sends file a copy of their own ----
function makeExperiment() {
  const folders = {
    "graph://me@graph.microsoft.com/Sent%20Items": { server: { type: "graph" } },
    "mailbox://nobody@Local%20Folders/Sent": { server: { type: "none" } },
  };
  const outgoing = { graphId: { type: "graph" }, smtpId: { type: "smtp" } };
  const listeners = [];
  const deep = () =>
    new Proxy(function () {}, {
      get: (target, key) => (key === Symbol.toPrimitive ? () => 0 : deep()),
      apply: () => deep(),
      construct: () => deep(),
    });
  const modules = {
    "resource:///modules/MailServices.sys.mjs": {
      MailServices: { outgoingServer: { getServerByIdentity: (identity) => outgoing[identity.key] } },
    },
    "resource:///modules/MailUtils.sys.mjs": { MailUtils: { getExistingFolder: (uri) => folders[uri] || null } },
    "resource:///modules/ExtensionSupport.sys.mjs": {
      ExtensionSupport: {
        registerWindowListener: (name, options) => listeners.push({ name, options }),
        unregisterWindowListener: (name) => listeners.splice(listeners.findIndex((l) => l.name === name), 1),
      },
    },
  };
  const ctx = {
    ChromeUtils: { importESModule: (url) => modules[url] || new Proxy({}, { get: () => deep() }), generateQI: () => () => {} },
    Ci: new Proxy(
      { nsIMsgCompDeliverMode: { Now: 0, Later: 1, Save: 2, SaveAs: 3, SaveAsDraft: 4, SaveAsTemplate: 5, SendUnsent: 6, AutoSaveAsDraft: 7, Background: 8 } },
      { get: (target, key) => target[key] || deep() }
    ),
    Cc: deep(),
    Cr: deep(),
    Components: deep(),
    Services: deep(),
    ExtensionCommon: { ExtensionAPI: class {} },
    console, Date, Math, JSON, Promise, Map, Set, Object, Array, String, Number, Boolean, Error, URL,
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(ROOT + "experiments/outlookPane/implementation.js", "utf8"), ctx);
  vm.runInContext("this.api = { paneFccIsSecondCopy, paneSetSingleSentCopy, paneState };", ctx);
  return { api: ctx.api, listeners };
}

function testExperiment() {
  const { api, listeners } = makeExperiment();
  const graphIdentity = { key: "graphId", doFcc: true, fccFolderURI: "graph://me@graph.microsoft.com/Sent%20Items" };
  const smtpIdentity = { key: "smtpId", doFcc: true, fccFolderURI: "graph://me@graph.microsoft.com/Sent%20Items" };

  assert.equal(api.paneFccIsSecondCopy(graphIdentity, ""), false, "off until the background turns it on");
  api.paneSetSingleSentCopy(true);
  assert.equal(listeners.length, 1, "compose windows watched");
  assert.equal(api.paneFccIsSecondCopy(graphIdentity, ""), true, "Graph send, default Sent Items: second copy");
  assert.equal(api.paneFccIsSecondCopy(graphIdentity, "graph://me@graph.microsoft.com/Sent%20Items"), true);
  assert.equal(api.paneFccIsSecondCopy(graphIdentity, "nocopy://"), false);
  assert.equal(api.paneFccIsSecondCopy(graphIdentity, "mailbox://nobody@Local%20Folders/Sent"), false, "a local folder copy is not on the server");
  assert.equal(api.paneFccIsSecondCopy(smtpIdentity, ""), false, "SMTP: the server keeps no copy");
  assert.equal(api.paneFccIsSecondCopy({ ...graphIdentity, doFcc: false }, ""), false, "no copy anyway");

  // The compose window: compose-send-message just before sendMsg().
  const handlers = {};
  const win = {
    addEventListener: (type, fn) => (handlers[type] = fn),
    removeEventListener: (type) => delete handlers[type],
    gCurrentIdentity: graphIdentity,
    gMsgCompose: { compFields: { fcc: "" } },
  };
  listeners[0].options.onLoadWindow(win);
  handlers["compose-send-message"]({ currentTarget: win, detail: { msgType: 4 } }); // SaveAsDraft
  assert.equal(win.gMsgCompose.compFields.fcc, "", "drafts untouched");
  handlers["compose-send-message"]({ currentTarget: win, detail: { msgType: 0 } }); // Now
  assert.equal(win.gMsgCompose.compFields.fcc, "nocopy://", "send now: no second copy");
  assert.equal(api.paneState.sentCopiesSkipped, 1);

  win.gCurrentIdentity = smtpIdentity;
  win.gMsgCompose.compFields.fcc = "";
  handlers["compose-send-message"]({ currentTarget: win, detail: { msgType: 0 } });
  assert.equal(win.gMsgCompose.compFields.fcc, "", "SMTP send keeps its copy");

  api.paneSetSingleSentCopy(false);
  assert.equal(listeners.length, 0, "listener removed");
  assert.equal(handlers["compose-send-message"], undefined, "event listener removed from open windows");
  assert.equal(api.paneFccIsSecondCopy(graphIdentity, ""), false);
}

(async () => {
  await testBackground();
  testExperiment();
  console.log("sent-copies tests passed");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
