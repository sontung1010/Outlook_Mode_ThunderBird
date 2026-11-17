"use strict";
// Run: node tests/server-health.test.js
// Runs common.js and graph.js with a fake outlookServer experiment (folder
// list sync that works or fails on demand) and a fake Graph, and checks:
// the "failing" and "back" notices, Thunderbird's notices held back while
// failing, requests that wait, Refresh-style checks, throttling.
const fs = require("fs");
const vm = require("vm");
const assert = require("assert");

const SRC = require("path").join(__dirname, "../outlook-mode/background/");
const code = ["common.js", "graph.js"].map((f) => fs.readFileSync(SRC + f, "utf8")).join("\n;\n");
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const notices = [];
const quiet = new Map();
const newMail = [];
const refreshStates = [];
let folderSync = { ok: true }; // what the next folder list sync gives
let folderSyncs = 0;
let serverErrorListener = null;
let graph = { status: 200, retryAfter: null }; // what Graph answers
const requests = [];

const messenger = {
  storage: { local: { get: async () => ({}), set: async () => {} }, onChanged: { addListener() {} } },
  notifications: {
    create: async (options) => {
      notices.push(options);
    },
  },
  accounts: { get: async (id) => ({ id, name: "Work", identities: [{ email: "me@corp.example" }] }) },
  outlookPin: {
    listGraphAccounts: async () => ["account1"],
    getGraphToken: async () => "token",
    setRefreshState: async (state) => {
      refreshStates.push(state);
    },
  },
  outlookServer: {
    start: async () => true,
    setQuiet: async (id, value) => {
      quiet.set(id, value);
    },
    syncFolderList: async () => {
      folderSyncs++;
      await sleep(5);
      return { ...folderSync };
    },
    getNewMail: async (id) => {
      newMail.push(id);
      return true;
    },
    status: async () => ({ watching: true }),
    onServerError: {
      addListener(fn) {
        serverErrorListener = fn;
      },
    },
  },
};

async function fakeFetch(url, init = {}) {
  requests.push([init.method || "GET", url]);
  const status = graph.status;
  if (graph.once) {
    graph = graph.once;
  }
  return {
    ok: status < 400,
    status,
    headers: { get: (name) => (name === "Retry-After" ? graph.retryAfterHeader ?? null : null) },
    json: async () => (status < 400 ? { value: [1] } : { error: { message: "UnknownError" } }),
  };
}

const ctx = {
  messenger, fetch: fakeFetch, console, URLSearchParams, URL, setTimeout, clearTimeout,
  Promise, JSON, Date, Math, Object, Array, Map, Set, String, Number, Error, TypeError, structuredClone,
};
vm.createContext(ctx);
vm.runInContext(
  code +
    `
;SERVER_RECHECK_MS = [60, 120];
SERVER_CHECK_GAP_MS = 300;
GRAPH_WAIT_MS = 2000;
GRAPH_THROTTLE_MAX_MS = 50;
this.api = { graphFetch, serverFailing, checkServer, noteServerFailure, serverHealthStatus, serverBackHooks };`,
  ctx
);
const api = ctx.api;
const backHooks = [];
api.serverBackHooks.push((id) => backHooks.push(id));
const titles = () => notices.map((n) => n.title);

(async () => {
  assert.ok(serverErrorListener, "listens to Thunderbird's error notices");

  /* 1. one error notice, the folder list sync works: nothing to tell */
  folderSync = { ok: true };
  serverErrorListener("account1", "Server graph.microsoft.com responded with an error.");
  await sleep(30);
  assert.strictEqual(folderSyncs, 1, "checked once");
  assert.deepStrictEqual(titles(), []);
  assert.ok(!api.serverFailing("account1"));
  // Another notice soon after: not checked again (gap).
  serverErrorListener("account1", "Server graph.microsoft.com responded with an error.");
  await sleep(30);
  assert.strictEqual(folderSyncs, 1, "not checked again within the gap");

  /* 2. this computer's network: not Microsoft's server */
  await sleep(300);
  folderSync = { ok: false, network: true, reason: "Could not connect to server graph.microsoft.com" };
  serverErrorListener("account1", "Server graph.microsoft.com responded with an error.");
  await sleep(30);
  assert.ok(!api.serverFailing("account1"), "network trouble is not a failing server");
  assert.deepStrictEqual(titles(), []);

  /* 3. the server fails: one notice, notices held back, rechecks */
  await sleep(300);
  folderSync = { ok: false, reason: "Server graph.microsoft.com responded with an error." };
  serverErrorListener("account1", "Server graph.microsoft.com responded with an error.");
  await sleep(30);
  assert.ok(api.serverFailing("account1"));
  assert.deepStrictEqual(titles(), ["Microsoft's server is failing"]);
  assert.match(notices[0].message, /me@corp\.example cannot sync: Server graph\.microsoft\.com responded with an error\./);
  assert.strictEqual(quiet.get("account1"), true, "Thunderbird's notices held back");
  assert.strictEqual(refreshStates.at(-1).error, true, "refresh button shows it");
  // More notices while failing: no new check, no new notice.
  const before = folderSyncs;
  serverErrorListener("account1", "Server graph.microsoft.com responded with an error.");
  await sleep(20);
  assert.strictEqual(folderSyncs, before);
  // Recheck after 60 ms: still failing, still one notice.
  await sleep(70);
  assert.strictEqual(folderSyncs, before + 1, "rechecked");
  assert.deepStrictEqual(titles(), ["Microsoft's server is failing"]);

  /* 4. requests wait while failing; a page's request fails at once */
  const requestsBefore = requests.length;
  let fetched = null;
  const waiting = api.graphFetch("account1", "/me/messages?$top=1").then((r) => (fetched = r));
  await assert.rejects(api.graphFetch("account1", "/me/mailFolders", {}, { wait: false }), /failing at the moment/);
  await sleep(30);
  assert.strictEqual(fetched, null, "the request waits");
  assert.strictEqual(requests.length, requestsBefore, "nothing sent while failing");

  /* 5. back: one notice, notices shown again, new mail, waiting request goes on */
  folderSync = { ok: true };
  await sleep(150); // next recheck (120 ms)
  assert.ok(!api.serverFailing("account1"), "back");
  assert.deepStrictEqual(titles(), ["Microsoft's server is failing", "Microsoft's server is back"]);
  assert.match(notices[1].message, /me@corp\.example syncs again \(it was failing for 1 min\)/);
  assert.strictEqual(quiet.get("account1"), false);
  assert.deepStrictEqual(newMail, ["account1"], "the account's new mail is fetched");
  assert.deepStrictEqual(backHooks, ["account1"], "put-off syncs run");
  await waiting;
  assert.ok(fetched, "the waiting request went on");
  assert.strictEqual(refreshStates.at(-1).error, false);

  /* 6. a 503 to the add-on: checked at once; a one-off is retried */
  const syncsBefore = folderSyncs;
  graph = { status: 503, once: { status: 200 } };
  const result = await api.graphFetch("account1", "/me/messages?$top=1");
  assert.ok(result, "retried after the check said the server is fine");
  assert.strictEqual(folderSyncs, syncsBefore + 1, "checked although within the gap");
  assert.strictEqual(titles().length, 2, "no notice for a one-off");

  /* 7. a POST is not sent again after a 503 */
  graph = { status: 503, once: { status: 200 } };
  const posts = requests.filter(([m]) => m === "POST").length;
  await assert.rejects(api.graphFetch("account1", "/me/mailFolders", { method: "POST", body: "{}" }), /Graph 503/);
  assert.strictEqual(requests.filter(([m]) => m === "POST").length, posts + 1, "sent once");

  /* 8. throttled (429): waits as asked (bounded), then once more */
  graph = { status: 429, retryAfterHeader: "1", once: { status: 200 } };
  const t0 = Date.now();
  assert.ok(await api.graphFetch("account1", "/me/messages?$top=1"));
  assert.ok(Date.now() - t0 >= 45, "waited");
  assert.ok(!api.serverFailing("account1"), "throttling is not a failure");

  /* 9. failing longer than a request may wait: it gives up */
  graph = { status: 200 };
  await sleep(350);
  folderSync = { ok: false, reason: "Server graph.microsoft.com responded with an error." };
  await api.noteServerFailure("account1", "x", { force: true });
  assert.ok(api.serverFailing("account1"));
  vm.runInContext("GRAPH_WAIT_MS = 100;", ctx);
  await assert.rejects(api.graphFetch("account1", "/me/messages?$top=1"), /still failing/);
  // Diagnostics keep the history.
  const status = api.serverHealthStatus();
  assert.strictEqual(status.accounts[0].state, "failing");
  assert.ok(status.log.some((e) => /back after/.test(e.text)));

  /* 10. a Refresh-style forced check brings it back at once */
  folderSync = { ok: true };
  const check = await api.checkServer("account1", { force: true });
  assert.ok(check.ok && !api.serverFailing("account1"));
  assert.strictEqual(titles().filter((t) => /back/.test(t)).length, 2);

  console.log("server-health: all checks passed");
  process.exit(0);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
