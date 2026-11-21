"use strict";
// Run: node tests/graph-index.test.js
// experiments/outlookServer keeps Thunderbird's index (Gloda) up to date for
// Microsoft Graph accounts: mail the Graph sync adds is given to Gloda's
// indexer, and Graph folders whose message count changed (all on a forced
// check) are marked dirty for Gloda's sweep. Other accounts are left alone.
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");

const code = fs.readFileSync(path.join(__dirname, "../outlook-mode/experiments/outlookServer/implementation.js"), "utf8");

function folder(server, name, count, indexable = true) {
  const glodaFolder = { dirty: 0, _ensureFolderDirty() { this.dirty++; } };
  return { URI: `${server.type}://${name}`, prettyName: name, server, count, indexable, glodaFolder, getTotalMessages: function () { return this.count; } };
}
const graph = { type: "graph" };
const imap = { type: "imap" };
const inbox = folder(graph, "Inbox", 10);
const sent = folder(graph, "Sent Items", 5);
const outbox = folder(graph, "Outbox", 0, false);
const imapInbox = folder(imap, "INBOX", 3);
graph.rootFolder = { descendants: [inbox, sent, outbox] };
imap.rootFolder = { descendants: [imapInbox] };

const mfnListeners = [];
const indexed = [];
const timers = [];
const glodaIndexer = {
  enabled: true,
  indexingSweepNeeded: false,
  shouldIndexFolder: (f) => f.indexable,
  indexMessages: (items) => indexed.push(...items),
};
const ctx = {
  ChromeUtils: {
    importESModule: (url) => {
      if (url.endsWith("MailServices.sys.mjs")) {
        return {
          MailServices: {
            accounts: { allServers: [graph, imap] },
            mfn: {
              addListener: (l, flags) => mfnListeners.push({ l, flags }),
              removeListener: (l) => mfnListeners.splice(mfnListeners.findIndex((x) => x.l === l), 1),
            },
          },
        };
      }
      if (url.endsWith("Timer.sys.mjs")) {
        return { setTimeout: (fn) => timers.push(fn), clearTimeout: () => {} };
      }
      if (url.endsWith("IndexMsg.sys.mjs")) {
        return { GlodaMsgIndexer: glodaIndexer };
      }
      if (url.endsWith("GlodaDatastore.sys.mjs")) {
        return { GlodaDatastore: { _mapFolder: (f) => f.glodaFolder } };
      }
      return {};
    },
    generateQI: () => () => {},
  },
  Ci: { nsIMsgFolderNotificationService: { msgAdded: 0x1 } },
  Services: {},
  ExtensionCommon: { ExtensionAPI: class {} },
  console, Date, Map, Set, JSON, Math, String, Number, Boolean, Object, Array, Error,
};
vm.createContext(ctx);
vm.runInContext(code + "\n;this.api = { srvIndexGraphMail, srvIndexStop, srvIndexFlush, srvState };", ctx);
const { api } = ctx;

// 1. Forced check (startup, Refresh): every indexable Graph folder.
assert.equal(api.srvIndexGraphMail(true), 2);
assert.equal(inbox.glodaFolder.dirty, 1);
assert.equal(sent.glodaFolder.dirty, 1);
assert.equal(outbox.glodaFolder.dirty, 0, "Gloda does not index it");
assert.equal(imapInbox.glodaFolder.dirty, 0, "not a Graph account");
assert.equal(glodaIndexer.indexingSweepNeeded, true);
assert.equal(mfnListeners.length, 1, "new mail watched");
assert.equal(mfnListeners[0].flags, 0x1);

// 2. Regular check: nothing changed -> nothing marked.
glodaIndexer.indexingSweepNeeded = false;
assert.equal(api.srvIndexGraphMail(false), 0);
assert.equal(glodaIndexer.indexingSweepNeeded, false);

// 3. The reply sent from Thunderbird reaches Sent Items through the Graph sync.
sent.count = 6;
assert.equal(api.srvIndexGraphMail(false), 1);
assert.equal(sent.glodaFolder.dirty, 2);
assert.equal(inbox.glodaFolder.dirty, 1);
assert.equal(glodaIndexer.indexingSweepNeeded, true);

// 4. msgAdded on a Graph folder: given to Gloda's indexer in one job.
const listener = mfnListeners[0].l;
listener.msgAdded({ folder: sent, messageKey: 345 });
listener.msgAdded({ folder: inbox, messageKey: 77 });
listener.msgAdded({ folder: imapInbox, messageKey: 9 }); // other account: Gloda handles it
assert.equal(timers.length, 1, "one flush for a batch");
timers.shift()();
assert.deepEqual(indexed.map(([f, k]) => [f.prettyName, k]), [["Sent Items", 345], ["Inbox", 77]]);
assert.equal(api.srvState.index.queued, 2);

// 5. Gloda off in Thunderbird: nothing is done.
glodaIndexer.enabled = false;
sent.count = 7;
assert.equal(api.srvIndexGraphMail(false), 0);
glodaIndexer.enabled = true;

// 6. The add-on stops: the listener is removed.
api.srvIndexStop();
assert.equal(mfnListeners.length, 0);

console.log("graph-index: all checks passed");
