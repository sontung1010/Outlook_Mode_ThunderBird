"use strict";
// Run: node tests/pin-sync.test.js
// Runs pin.js and pin-sync.js against an in-memory Microsoft Graph mailbox
// (Outlook's pin: SystemTime 0x0F01 / 0x0F02 in the year 4500) and fake
// Thunderbird message stores, and checks that pins go both ways.
const fs = require("fs");
const vm = require("vm");
const assert = require("assert");

const SRC = require("path").join(__dirname, "../outlook-mode/background/");
const code = ["common.js", "graph.js", "pin.js", "pin-sync.js"].map((f) => fs.readFileSync(SRC + f, "utf8")).join("\n;\n");

const FLAG = "SystemTime 0x0F02";
const TIME = "SystemTime 0x0F01";
const PIN_DATE = "4500-09-01T00:00:00Z";

// ---- fake Graph mailbox, shared by the devices (one Microsoft 365 account) ----
const server = {
  items: new Map(), // id -> { id, internetMessageId, receivedDateTime, folder, props }
  supportsQuery: true, // Graph accepts the cast(...) filter on the pin property
  offline: false,
  patches: [],
};
function addServerItem(id, folder = "inbox", received = "2026-10-01T08:00:00Z") {
  server.items.set(id, { id, internetMessageId: `<${id}@x>`, receivedDateTime: received, modified: received, folder, props: {} });
}
const touch = (item) => (item.modified = new Date().toISOString());
const pinnedOnServer = (id) => Date.parse(server.items.get(id).props[FLAG] || 0) > Date.parse("4000-01-01T00:00:00Z");
function outlookPins(id, pinned) {
  const item = server.items.get(id);
  item.props[TIME] = item.props[FLAG] = pinned ? PIN_DATE : item.receivedDateTime;
  touch(item);
}

async function fakeFetch(url, init = {}) {
  if (server.offline) {
    throw new TypeError("NetworkError when attempting to fetch resource.");
  }
  const u = new URL(url);
  const path = u.pathname.replace("/v1.0", "");
  const method = init.method || "GET";
  const json = (status, body) => ({ ok: status < 400, status, json: async () => body });
  const withProps = (item) => ({
    id: item.id,
    internetMessageId: item.internetMessageId,
    receivedDateTime: item.receivedDateTime,
    // Graph names them its own way: "SystemTime 0xf02".
    singleValueExtendedProperties: Object.entries(item.props).map(([id, value]) => ({ id: id.replace(/0x0/, "0x").toLowerCase().replace("systemtime", "SystemTime"), value })),
  });
  const filter = u.searchParams.get("$filter") || "";
  let m;
  if (path === "/me/messages" && method === "GET") {
    if ((m = filter.match(/^lastModifiedDateTime ge (\S+)$/))) {
      const since = Date.parse(m[1]);
      return json(200, { value: [...server.items.values()].filter((i) => Date.parse(i.modified) >= since).map(withProps) });
    }
    if ((m = filter.match(/^internetMessageId eq '(.*)'$/))) {
      return json(200, { value: [...server.items.values()].filter((i) => i.internetMessageId === m[1]).map(withProps) });
    }
    if (filter.startsWith("singleValueExtendedProperties/any(")) {
      assert.match(filter, /ep\/id eq 'SystemTime 0x0F02'/);
      if (!server.supportsQuery) {
        return json(400, { error: { message: "Invalid filter clause" } });
      }
      const after = filter.match(/gt (\S+?)(?: and|\))/)[1];
      const before = (filter.match(/lt (\S+?)\)/) || [])[1] || "9999-12-31T00:00:00Z";
      const inRange = (i) => i.props[FLAG] && Date.parse(i.props[FLAG]) > Date.parse(after) && Date.parse(i.props[FLAG]) < Date.parse(before);
      return json(200, { value: [...server.items.values()].filter(inRange).map(withProps) });
    }
  }
  if (path === "/me/mailFolders/inbox/messages" && method === "GET") {
    assert.equal(u.searchParams.get("$orderby"), "receivedDateTime desc");
    const skip = Number(u.searchParams.get("$skip") || 0);
    const inbox = [...server.items.values()].filter((i) => i.folder === "inbox").sort((a, b) => b.receivedDateTime.localeCompare(a.receivedDateTime));
    const page = inbox.slice(skip, skip + 2); // tiny pages, to exercise nextLink
    const next = new URL(url);
    next.searchParams.set("$skip", String(skip + 2));
    return json(200, { value: page.map(withProps), "@odata.nextLink": skip + 2 < inbox.length ? next.toString() : undefined });
  }
  if ((m = path.match(/^\/me\/messages\/([^/]+)$/)) && method === "PATCH") {
    const item = server.items.get(decodeURIComponent(m[1]));
    const body = JSON.parse(init.body);
    for (const { id, value } of body.singleValueExtendedProperties) {
      item.props[id] = value;
    }
    touch(item);
    server.patches.push([item.id, body]);
    return json(200, withProps(item));
  }
  throw new Error(`unexpected ${method} ${url}`);
}

// ---- one Thunderbird device ----
function makeDevice(name, localIds) {
  const storage = {};
  const messages = localIds.map((id, index) => ({
    id: index + 1,
    headerMessageId: `${id}@x`,
    subject: id,
    tags: [],
    folder: { accountId: "account1", id: "inbox", type: "inbox" },
  }));
  const list = (items) => ({ id: null, messages: items.map((x) => structuredClone(x)) });
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
        async remove(k) {
          delete storage[k];
        },
      },
      onChanged: { addListener() {} },
    },
    messages: {
      onNewMailReceived: { addListener() {} },
      async query(q) {
        let found = messages;
        if (q.headerMessageId) {
          found = found.filter((x) => x.headerMessageId === q.headerMessageId);
        }
        if (q.tags) {
          const keys = Object.keys(q.tags.tags);
          found = found.filter((x) => keys.some((k) => x.tags.includes(k)));
        }
        return list(found);
      },
      async get(id) {
        return structuredClone(messages.find((x) => x.id === id));
      },
      async update(id, { tags }) {
        messages.find((x) => x.id === id).tags = [...tags];
      },
      async getHeaders() {
        return {}; // no References: every message is its own thread root
      },
      tags: {
        _list: [],
        async list() {
          return this._list;
        },
        async create(key, tag, color) {
          this._list.push({ key, tag, color });
        },
      },
    },
    accounts: { list: async () => [{ id: "account1" }] },
    alarms: { create() {}, onAlarm: { addListener() {} } },
    mailTabs: { query: async () => [], onDisplayedFolderChanged: { addListener() {} } },
    runtime: { onInstalled: { addListener() {} } },
    notifications: { create: async () => {} },
    outlookPin: {
      register: async () => {},
      listGraphAccounts: async () => ["account1"],
      getGraphToken: async () => "token",
    },
  };
  const ctx = {
    messenger, fetch: fakeFetch, console, URLSearchParams, URL, setTimeout: () => 0, clearTimeout: () => {},
    Promise, JSON, Date, Math, Object, Array, Map, Set, String, Number, Error, TypeError, structuredClone,
  };
  vm.createContext(ctx);
  vm.runInContext(code + "\n;this.api = { syncAllPins, togglePin, pinSyncStatus, PIN_TAG_KEY, repairUnpinDates };", ctx);
  const local = (id) => messages.find((x) => x.headerMessageId === `${id}@x`);
  return {
    name,
    api: ctx.api,
    storage,
    local,
    pinnedHere: (id) => local(id).tags.includes(ctx.api.PIN_TAG_KEY),
    pinHere: (id) => (local(id).tags = [...local(id).tags, ctx.api.PIN_TAG_KEY]),
    toggle: (id) => ctx.api.togglePin([structuredClone(local(id))]),
  };
}

// Pin changes made in Thunderbird upload in the background.
const settle = async () => {
  for (let i = 0; i < 50; i++) {
    await new Promise((resolve) => setImmediate(resolve));
  }
};

(async () => {
  for (const id of ["m1", "m2", "m3"]) {
    addServerItem(id, "inbox", `2026-10-0${id.slice(1)}T08:00:00Z`);
  }

  // 1. First sync on a device: pins made before sync go to the server.
  const A = makeDevice("A", ["m1", "m2", "m3"]);
  A.pinHere("m1");
  let r = await A.api.syncAllPins({ force: true });
  assert.ok(r.ok, JSON.stringify(r));
  assert.ok(pinnedOnServer("m1"), "seeded pin uploaded");
  assert.equal(server.items.get("m1").props[TIME], PIN_DATE, "both properties, like Outlook");
  assert.ok(A.pinnedHere("m1"));

  // 2. Pinned in Outlook (Windows, web or mobile) → pinned in Thunderbird.
  outlookPins("m2", true);
  r = await A.api.syncAllPins({ force: true });
  assert.equal(r.changed, 1);
  assert.ok(A.pinnedHere("m2"), "Outlook's pin shows in Thunderbird");

  // 3. Unpinned in Outlook → unpinned in Thunderbird.
  outlookPins("m1", false);
  await A.api.syncAllPins({ force: true });
  assert.ok(!A.pinnedHere("m1"), "Outlook's unpin shows in Thunderbird");

  // 4. Unpinned in Thunderbird → unpinned on the server, back at its received date
  //    (Outlook sorts by 0x0F01: "now" would move an old mail to the top).
  await A.toggle("m2");
  await settle();
  assert.ok(!A.pinnedHere("m2"));
  assert.ok(!pinnedOnServer("m2"), "unpin uploaded at once");
  assert.equal(server.items.get("m2").props[TIME], "2026-10-02T08:00:00Z");
  assert.equal(server.items.get("m2").props[FLAG], "2026-10-02T08:00:00Z");

  // 5. Pinned while offline: kept locally and queued, uploaded when back online.
  server.offline = true;
  await A.toggle("m3");
  await settle();
  assert.ok(A.pinnedHere("m3"), "pinned locally at once");
  r = await A.api.syncAllPins({ force: true });
  assert.equal(r.ok, false, "sync fails offline");
  assert.ok(A.pinnedHere("m3"), "an offline sync never removes a waiting pin");
  assert.ok(Object.keys(A.storage.pinSyncPending.account1).length, "queued");
  server.offline = false;
  await A.api.syncAllPins({ force: true });
  assert.ok(pinnedOnServer("m3"), "queued pin uploaded");
  assert.deepEqual(A.storage.pinSyncPending.account1, {});
  assert.equal(A.api.pinSyncStatus.account1.method, "query");

  // 6. A server that refuses the filter: the Inbox is read instead, and
  //    local pins outside what was read are left alone.
  server.supportsQuery = false;
  addServerItem("m4", "inbox", "2026-10-04T08:00:00Z");
  addServerItem("m5", "archive", "2026-09-01T08:00:00Z");
  outlookPins("m4", true);
  outlookPins("m5", true);
  const B = makeDevice("B", ["m1", "m2", "m3", "m4", "m5"]);
  r = await B.api.syncAllPins({ force: true });
  assert.ok(r.ok, JSON.stringify(r));
  assert.equal(B.api.pinSyncStatus.account1.method, "scan");
  assert.ok(B.pinnedHere("m3") && B.pinnedHere("m4"), "Inbox pins found by reading the Inbox");
  assert.ok(B.pinnedHere("m5"), "a pin made in Outlook on mail outside the Inbox comes through (it changed)");
  outlookPins("m5", false);
  outlookPins("m4", false);
  await B.api.syncAllPins({ force: true });
  assert.ok(!B.pinnedHere("m4"), "an Inbox unpin is seen");
  assert.ok(!B.pinnedHere("m5"), "an unpin outside the Inbox is seen too");
  addServerItem("m6", "archive", "2026-07-01T08:00:00Z"); // old, untouched on the server
  const B2 = makeDevice("B2", ["m6"]);
  B2.storage.pinSyncSeeded = ["account1"]; // past its first sync
  B2.pinHere("m6");
  await B2.api.syncAllPins({ force: true });
  assert.ok(B2.pinnedHere("m6"), "a local pin on mail outside what was read is left alone");

  // 7. Steady state: nothing changes, nothing uploads.
  server.patches.length = 0;
  r = await A.api.syncAllPins({ force: true });
  r = await A.api.syncAllPins({ force: true });
  assert.equal(r.changed, 0);
  assert.equal(server.patches.length, 0, "no uploads without changes");

  // 8. Repair: mail that the old unpin (both properties = the time of the
  //    unpin) moved to the top goes back to its received date; nothing else
  //    is touched.
  const hourAgo = new Date(Date.now() - 3600e3).toISOString().replace(/\.\d+Z$/, "Z");
  addServerItem("r1", "inbox", "2026-10-19T08:12:16Z");
  server.items.get("r1").props = { [TIME]: hourAgo, [FLAG]: hourAgo }; // the old unpin
  touch(server.items.get("r1"));
  addServerItem("r2", "inbox", "2026-10-20T08:00:00Z");
  server.items.get("r2").props = { [TIME]: hourAgo }; // Outlook's own unpin: no 0x0F02
  touch(server.items.get("r2"));
  addServerItem("r3", "inbox", "2026-10-21T08:00:00Z");
  outlookPins("r3", true); // pinned
  addServerItem("r4", "archive", "2026-10-22T08:00:00Z");
  server.items.get("r4").props = { [TIME]: hourAgo, [FLAG]: hourAgo }; // old unpin, not in the Inbox
  touch(server.items.get("r4"));
  server.supportsQuery = true;
  r = await A.api.repairUnpinDates();
  assert.equal(r.fixed, 2, JSON.stringify(r));
  assert.equal(server.items.get("r1").props[FLAG], "2026-10-19T08:12:16Z");
  assert.equal(server.items.get("r1").props[TIME], "2026-10-19T08:12:16Z");
  assert.equal(server.items.get("r4").props[FLAG], "2026-10-22T08:00:00Z");
  assert.equal(server.items.get("r2").props[TIME], hourAgo, "Outlook's own unpin left alone");
  assert.ok(pinnedOnServer("r3"), "pins left alone");
  r = await A.api.repairUnpinDates();
  assert.equal(r.fixed, 0, "once per account");
  //    Mail from months ago, in any folder, is found too (the damage changed it).
  addServerItem("r5", "archive", "2026-07-19T02:06:42Z");
  server.items.get("r5").props = { [TIME]: hourAgo, [FLAG]: hourAgo };
  touch(server.items.get("r5"));
  r = await A.api.repairUnpinDates({ force: true });
  assert.equal(r.fixed, 1);
  assert.equal(server.items.get("r5").props[FLAG], "2026-07-19T02:06:42Z");

  console.log("all pin-sync scenarios passed");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
