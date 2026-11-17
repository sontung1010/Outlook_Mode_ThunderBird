"use strict";
// Run: node tests/settings-sync.test.js
// Simulates two Thunderbird devices running settings-sync.js against an
// in-memory Microsoft Graph mailbox.
const fs = require("fs");
const vm = require("vm");
const assert = require("assert");

const SRC = require("path").join(__dirname, "../outlook-mode/background/");
const code = ["common.js", "graph.js", "settings-sync.js"].map((f) => fs.readFileSync(SRC + f, "utf8")).join("\n;\n");

// ---- fake Graph mailbox, shared by all devices ----
const mailbox = { folders: [], messages: new Map(), nextId: 1 };
async function fakeFetch(url, init = {}) {
  const u = new URL(url);
  const path = u.pathname.replace("/v1.0", "");
  const method = init.method || "GET";
  const json = (status, body) => ({
    ok: status < 400,
    status,
    json: async () => body,
  });
  let m;
  if (path === "/me/mailFolders" && method === "GET") {
    assert.equal(u.searchParams.get("includeHiddenFolders"), "true");
    return json(200, { value: mailbox.folders });
  }
  if (path === "/me/mailFolders" && method === "POST") {
    const f = { id: `F${mailbox.nextId++}`, ...JSON.parse(init.body) };
    mailbox.folders.push(f);
    return json(201, f);
  }
  if ((m = path.match(/^\/me\/mailFolders\/([^/]+)\/messages$/))) {
    const folder = decodeURIComponent(m[1]);
    if (method === "GET") {
      return json(200, { value: [...mailbox.messages.values()].filter((x) => x.folder === folder).map((x) => ({ id: x.id })) });
    }
    const msg = { id: `M${mailbox.nextId++}`, folder, ...JSON.parse(init.body) };
    mailbox.messages.set(msg.id, msg);
    return json(201, { id: msg.id });
  }
  if ((m = path.match(/^\/me\/messages\/([^/]+)\/attachments$/))) {
    const msg = mailbox.messages.get(decodeURIComponent(m[1]));
    return json(200, { value: msg.attachments });
  }
  if ((m = path.match(/^\/me\/messages\/([^/]+)\/permanentDelete$/))) {
    const id = decodeURIComponent(m[1]);
    if (!mailbox.messages.has(id)) return json(404, { error: { message: "gone" } });
    mailbox.messages.delete(id);
    return { ok: true, status: 204, json: async () => null };
  }
  throw new Error(`unexpected ${method} ${url}`);
}

// ---- one device ----
function makeDevice(name, { prefs = {}, contacts = {} } = {}) {
  const storage = {};
  const changeListeners = [];
  const dev = {
    name,
    prefs: { ...prefs }, // key -> value, as the experiment returns them
    books: { "Personal Address Book": { id: `${name}-pab`, contacts: { ...contacts } } },
    notes: [],
  };
  const messenger = {
    storage: {
      local: {
        async get(keys) {
          const list = keys == null ? Object.keys(storage) : Array.isArray(keys) ? keys : [keys];
          const out = {};
          for (const k of list) if (k in storage) out[k] = structuredClone(storage[k]);
          return out;
        },
        async set(obj) {
          Object.assign(storage, structuredClone(obj));
          changeListeners.forEach((l) => l(Object.fromEntries(Object.keys(obj).map((k) => [k, {}]))));
        },
        async remove(k) {
          delete storage[k];
        },
      },
      onChanged: { addListener: (l) => changeListeners.push(l) },
    },
    alarms: { create() {}, onAlarm: { addListener() {} } },
    notifications: { create: async (o) => dev.notes.push(o.message) },
    outlookPin: { listGraphAccounts: async () => ["account9"], getGraphToken: async () => "tok" },
    outlookSync: {
      async readSettings() {
        return { values: { ...dev.prefs }, servers: ["graph:me@graph.microsoft.com"], identities: ["me@x.com"], accounts: { account9: "graph:me@graph.microsoft.com" } };
      },
      async supportedKeys(keys) {
        return keys.filter((k) => k.startsWith("pref:") || k.includes("graph:me@graph.microsoft.com") || k.includes("me@x.com"));
      },
      async applySettings(changes) {
        const failed = {};
        for (const [k, v] of Object.entries(changes)) {
          if (k === "pref:mail.broken") { failed[k] = "boom"; continue; }
          if (v === null) delete dev.prefs[k];
          else dev.prefs[k] = v;
        }
        return { failed };
      },
    },
    addressBooks: {
      async list() {
        return Object.entries(dev.books).map(([bname, b]) => ({
          id: b.id, name: bname, remote: false, readOnly: false,
          contacts: Object.entries(b.contacts).map(([id, vCard]) => ({ id, properties: { vCard } })),
        }));
      },
      async create({ name: bname }) {
        dev.books[bname] = { id: `${name}-${bname}`, contacts: {} };
        return dev.books[bname].id;
      },
    },
    contacts: {
      async create(bookId, id, { vCard }) {
        const book = Object.values(dev.books).find((b) => b.id === bookId);
        book.contacts[id] = vCard.replace("END:VCARD", "PRODID:-//Thunderbird\r\nEND:VCARD");
      },
      async update(id, { vCard }) {
        for (const b of Object.values(dev.books)) if (id in b.contacts) b.contacts[id] = vCard;
      },
      async delete(id) {
        for (const b of Object.values(dev.books)) delete b.contacts[id];
      },
    },
    runtime: { onInstalled: { addListener() {} } },
  };
  const ctx = {
    messenger, fetch: fakeFetch, console, URLSearchParams, Blob, Response, CompressionStream, DecompressionStream,
    btoa, atob, TextDecoder, TextEncoder, setTimeout: () => 0, clearTimeout: () => {}, Promise, JSON, Date, Math, Object, Array, Map, Set, String, Number, Error, Uint8Array, structuredClone,
  };
  vm.createContext(ctx);
  vm.runInContext(code + "\n;this.api = { syncSettings, undoLastSettingsSync, settingsSyncStatus, cleanVCard };", ctx);
  dev.api = ctx.api;
  dev.storage = storage;
  dev.enable = async () => messenger.storage.local.set({ settings: { settingsSync: true } });
  return dev;
}

const vc = (fn, email) => `BEGIN:VCARD\r\nVERSION:4.0\r\nFN:${fn}\r\nEMAIL:${email}\r\nEND:VCARD`;

(async () => {
  let t = 1000;
  const realNow = Date.now;
  Date.now = () => (t += 1000);

  const A = makeDevice("A", { prefs: { "pref:mail.uidensity": 1, "pref:mailnews.tags.x.tag": "Work", "identity:me@x.com:reply_on_top": 1, "pref:mail.broken": 1 }, contacts: { u1: vc("Ann", "ann@x.com") } });
  const B = makeDevice("B", { prefs: { "pref:mail.uidensity": 2, "pref:font.size": 15 } });

  // 1. A is first: uploads everything.
  await A.enable();
  let r = await A.api.syncSettings();
  assert.ok(r.ok && r.uploaded, "A uploads");
  assert.equal(mailbox.messages.size, 1);
  assert.equal(mailbox.folders[0].isHidden, true);

  // 2. B joins: A's values win where both have them; B's own extras are kept and uploaded.
  await B.enable();
  r = await B.api.syncSettings();
  assert.equal(B.prefs["pref:mail.uidensity"], 1, "join: server wins");
  assert.equal(B.prefs["pref:mailnews.tags.x.tag"], "Work");
  assert.equal(B.prefs["identity:me@x.com:reply_on_top"], 1);
  assert.equal(B.prefs["pref:font.size"], 15, "B extra kept");
  assert.ok(B.api.settingsSyncStatus.lastFailed["pref:mail.broken"], "failure reported");
  assert.ok(Object.values(B.books["Personal Address Book"].contacts).length === 1, "contact copied");
  assert.equal(mailbox.messages.size, 1, "old state removed");

  // 3. A gets B's extra.
  await A.api.syncSettings();
  assert.equal(A.prefs["pref:font.size"], 15);

  // 4. Steady state: nothing uploads, no ping-pong (contact got a PRODID on B).
  r = await B.api.syncSettings();
  assert.equal(r.uploaded, false, "B steady");
  assert.equal(r.applied, 0);
  r = await A.api.syncSettings();
  assert.equal(r.uploaded, false, "A steady");

  // 5. Change on A, delete on B.
  A.prefs["pref:mail.uidensity"] = 0;
  delete B.prefs["pref:font.size"];
  delete B.books["Personal Address Book"].contacts.u1;
  await A.api.syncSettings();
  await B.api.syncSettings();
  await A.api.syncSettings();
  assert.equal(B.prefs["pref:mail.uidensity"], 0, "A's change reaches B");
  assert.ok(!("pref:font.size" in A.prefs), "B's reset reaches A");
  assert.ok(!("u1" in A.books["Personal Address Book"].contacts), "contact deletion reaches A");

  // 6. Concurrent edits: later one wins.
  A.prefs["pref:mail.uidensity"] = 5;
  B.prefs["pref:mail.uidensity"] = 6; // B's sync happens later -> its change is later
  await A.api.syncSettings();
  await B.api.syncSettings();
  await A.api.syncSettings();
  assert.equal(A.prefs["pref:mail.uidensity"], 6);
  assert.equal(B.prefs["pref:mail.uidensity"], 6);

  // 7. Undo on B puts back the values its last sync replaced.
  A.prefs["pref:mailnews.tags.x.tag"] = "Job";
  await A.api.syncSettings();
  await B.api.syncSettings();
  assert.equal(B.prefs["pref:mailnews.tags.x.tag"], "Job");
  await B.api.undoLastSettingsSync();
  assert.equal(B.prefs["pref:mailnews.tags.x.tag"], "Work", "undo restores");
  await A.api.syncSettings();
  assert.equal(A.prefs["pref:mailnews.tags.x.tag"], "Work", "undo propagates");

  // 8. Two states on the server (devices wrote at the same time) are merged.
  const [only] = mailbox.messages.values();
  mailbox.messages.set("Mdup", { ...structuredClone(only), id: "Mdup" });
  r = await A.api.syncSettings();
  assert.equal(mailbox.messages.size, 1, "duplicates consolidated");

  // 9. Add-on settings and Quick Steps sync, folder ids translated.
  const stepsA = [{ id: "s1", name: "Move", after: "move", moveTo: "account9://Inbox/Done" }];
  await A.api.syncSettings();
  A.storage.quickSteps = stepsA;
  A.storage.settings = { ...A.storage.settings, undoSendSeconds: 30 };
  await A.api.syncSettings();
  await B.api.syncSettings();
  assert.equal(B.storage.settings.undoSendSeconds, 30);
  assert.equal(B.storage.settings.settingsSync, true, "local-only setting untouched");
  assert.equal(B.storage.quickSteps[0].moveTo, "account9://Inbox/Done");
  const state = JSON.parse(await (async () => {
    const att = [...mailbox.messages.values()][0].attachments[0];
    const bytes = Uint8Array.from(atob(att.contentBytes), (c) => c.charCodeAt(0));
    return new Response(new Blob([bytes]).stream().pipeThrough(new DecompressionStream("gzip"))).text();
  })());
  assert.deepEqual(state.entries["addon:quickSteps"].v[0].moveTo, { account: "graph:me@graph.microsoft.com", path: "//Inbox/Done" });

  Date.now = realNow;
  console.log("all settings-sync scenarios passed");
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
