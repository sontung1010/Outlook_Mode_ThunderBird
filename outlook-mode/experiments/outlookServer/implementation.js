"use strict";

/* global ExtensionAPI, ExtensionCommon, ChromeUtils, Services, Cc, Ci */

// Experiment: Microsoft's server failing, for Microsoft Graph accounts.
//
// Thunderbird shows a desktop notice for every failed Graph operation:
// "Server graph.microsoft.com responded with an error." When Microsoft's
// server fails for a while (HTTP 502/503/504 on the folder list sync), that
// notice comes again at every check for mail, for as long as it lasts. This
// experiment:
// - reports those notices of Graph accounts to the add-on (onServerError),
//   which checks whether the server is really failing (background/graph.js);
// - while the add-on says the server of an account is failing (setQuiet),
//   holds back Thunderbird's notices for that account: the add-on shows one
//   notice when it starts failing and one when it is back;
// - runs Thunderbird's folder list sync of an account and tells whether it
//   worked (syncFolderList): used to check the server, and by Refresh to
//   restart a failing sync; Thunderbird's notice of such a run is held back,
//   since the add-on tells the result itself. getNewMail gets the account's
//   new mail.
//
// The notices are seen by wrapping onAlert of Thunderbird's alertHook (the
// module that shows them); the wrapper is removed when the add-on stops.
//
// It also keeps Thunderbird's message index (Gloda) up to date for Graph
// accounts (indexGraphMail). Gloda indexes new mail when Thunderbird reports
// it as classified (msgsClassified), which the Graph sync does not do for
// mail it brings in, e.g. the Sent Items copy of a mail sent from
// Thunderbird. Thunderbird Conversations finds the messages of a
// conversation in other folders through Gloda, so such mail was missing from
// the reading pane (a reply sent from Thunderbird, in Sent Items). Here:
// - mail added to a Graph folder (msgAdded) is given to Gloda's indexer;
// - a Graph folder whose message count changed since the last check (and
//   every Graph folder on a forced check: startup, Refresh) is marked dirty,
//   so Gloda's next sweep indexes what it does not know yet.

var { MailServices } = ChromeUtils.importESModule("resource:///modules/MailServices.sys.mjs");
var { setTimeout, clearTimeout } = ChromeUtils.importESModule("resource://gre/modules/Timer.sys.mjs");

const SRV_ALERT_HOOK_URL = "resource:///modules/activity/alertHook.sys.mjs";
const SRV_SYNC_TIMEOUT_MS = 90000;

const srvState = {
  hook: null, // Thunderbird's alertHook object
  original: null, // its own onAlert
  wrapper: null,
  quiet: new Set(), // account keys whose notices are held back
  captures: new Set(), // { serverKey, alerts: [] } of running folder list syncs
  listeners: new Set(), // onServerError
  seen: 0,
  held: 0,
  lastError: "",
  index: {
    listener: null, // nsIMsgFolderListener for msgAdded
    counts: new Map(), // folder URI -> message count at the last check
    queue: [], // [folder, key] of added mail, given to Gloda in one job
    flushTimer: null,
    queued: 0, // mail given to Gloda's indexer since start
    dirtied: 0, // folders marked for Gloda's sweep since start
    lastCheck: "",
  },
};

function srvNote(where, error) {
  srvState.lastError = `${new Date().toISOString()} ${where}: ${error && error.message ? error.message : error}`;
  console.error(`[outlook-mode] server: ${where}`, error);
}

function srvGraphServer(serverKey) {
  try {
    const server = MailServices.accounts.getIncomingServer(serverKey);
    return server && server.type == "graph" ? server : null;
  } catch (error) {
    return null;
  }
}

function srvAccountOf(server) {
  try {
    return MailServices.accounts.findAccountForServer(server)?.key || "";
  } catch (error) {
    return "";
  }
}

function srvServerOf(accountId) {
  try {
    const server = MailServices.accounts.getAccount(accountId)?.incomingServer;
    return server && server.type == "graph" ? server : null;
  } catch (error) {
    return null;
  }
}

// Thunderbird's own text for an error answer of the server, in the user's
// language: "Server graph.microsoft.com responded with an error."
function srvServerErrorText(server) {
  try {
    return Services.strings
      .createBundle("chrome://messenger/locale/messenger.properties")
      .formatStringFromName("errorResponseError", [server.hostName || server.hostname]);
  } catch (error) {
    return "";
  }
}

// Called for every notice of Thunderbird's mail code. Returns true when the
// notice is held back.
function srvOnAlert(message, serverKey) {
  const server = srvGraphServer(serverKey);
  if (!server) {
    return false;
  }
  srvState.seen++;
  let checking = false;
  for (const capture of srvState.captures) {
    if (capture.serverKey == serverKey) {
      capture.alerts.push(String(message || ""));
      checking = true;
    }
  }
  const accountId = srvAccountOf(server);
  // Only the server's error answers are signs of a failing server; "could
  // not connect" and time-outs are this computer's network as often as not.
  if (accountId && String(message || "") == srvServerErrorText(server)) {
    for (const fire of srvState.listeners) {
      try {
        fire.async(accountId, String(message));
      } catch (error) {
        srvNote("event", error);
      }
    }
  }
  // Held back: while the server is failing, and the notice of a check of
  // the add-on (its result is told by the add-on's own notice, if needed).
  if (accountId && (checking || srvState.quiet.has(accountId))) {
    srvState.held++;
    return true;
  }
  return false;
}

function srvWrap() {
  if (srvState.wrapper) {
    return;
  }
  const { alertHook } = ChromeUtils.importESModule(SRV_ALERT_HOOK_URL);
  if (!alertHook || typeof alertHook.onAlert != "function") {
    throw new Error("Thunderbird's alertHook was not found");
  }
  const original = alertHook.onAlert;
  const wrapper = function (message, serverKey, silent) {
    let held = false;
    try {
      held = srvOnAlert(message, serverKey, silent);
    } catch (error) {
      srvNote("alert", error);
    }
    return held ? true : original.call(this, message, serverKey, silent);
  };
  alertHook.onAlert = wrapper;
  Object.assign(srvState, { hook: alertHook, original, wrapper });
}

function srvUnwrap() {
  if (srvState.hook && srvState.hook.onAlert === srvState.wrapper) {
    srvState.hook.onAlert = srvState.original;
  }
  Object.assign(srvState, { hook: null, original: null, wrapper: null });
  srvState.quiet.clear();
}

function srvMsgWindow() {
  return Services.wm.getMostRecentWindow("mail:3pane")?.msgWindow || null;
}

// Runs Thunderbird's folder list sync of the account. It reports success
// even when it failed; a failure shows as a notice of Thunderbird for the
// account in the meantime (always before the callback).
function srvSyncFolderList(accountId) {
  const server = srvServerOf(accountId);
  if (!server) {
    return Promise.resolve({ ok: false, reason: "not a Microsoft Graph account" });
  }
  if (Services.io.offline) {
    return Promise.resolve({ ok: false, network: true, reason: "Thunderbird is offline" });
  }
  const window = srvMsgWindow();
  if (!window) {
    return Promise.resolve({ ok: false, network: true, reason: "no Thunderbird window" });
  }
  return new Promise((resolve) => {
    const capture = { serverKey: server.key, alerts: [] };
    srvState.captures.add(capture);
    let timer = null;
    const finish = (result) => {
      if (!srvState.captures.has(capture)) {
        return;
      }
      srvState.captures.delete(capture);
      clearTimeout(timer);
      resolve(result);
    };
    const settle = () => {
      if (!capture.alerts.length) {
        finish({ ok: true });
        return;
      }
      const reason = capture.alerts[capture.alerts.length - 1];
      finish({ ok: false, reason, network: reason != srvServerErrorText(server) });
    };
    timer = setTimeout(
      () => finish({ ok: false, reason: `no answer within ${SRV_SYNC_TIMEOUT_MS / 1000} s` }),
      SRV_SYNC_TIMEOUT_MS
    );
    try {
      server.QueryInterface(Ci.IExchangeIncomingServer).syncFolderHierarchy(
        {
          QueryInterface: ChromeUtils.generateQI(["IExchangeSimpleOperationListener"]),
          onOperationSuccess: settle,
        },
        window
      );
    } catch (error) {
      srvNote("folder list sync", error);
      finish({ ok: false, reason: String(error && error.message ? error.message : error).slice(0, 200) });
    }
  });
}

/* ------------------------------------------------------------------ */
/* Gloda index for Graph accounts                                      */
/* ------------------------------------------------------------------ */

const SRV_INDEX_FLUSH_MS = 3000;

function srvGloda() {
  try {
    const { GlodaMsgIndexer } = ChromeUtils.importESModule("resource:///modules/gloda/IndexMsg.sys.mjs");
    const { GlodaDatastore } = ChromeUtils.importESModule("resource:///modules/gloda/GlodaDatastore.sys.mjs");
    return GlodaMsgIndexer && GlodaMsgIndexer.enabled ? { GlodaMsgIndexer, GlodaDatastore } : null;
  } catch (error) {
    return null; // Gloda turned off or not there
  }
}

function srvGraphFolders() {
  const folders = [];
  for (const server of MailServices.accounts.allServers) {
    if (server.type == "graph") {
      try {
        folders.push(...server.rootFolder.descendants);
      } catch (error) {
        srvNote("folders", error);
      }
    }
  }
  return folders;
}

function srvIndexFlush() {
  const index = srvState.index;
  index.flushTimer = null;
  const items = index.queue.splice(0);
  const gloda = srvGloda();
  if (!gloda || !items.length) {
    return;
  }
  try {
    gloda.GlodaMsgIndexer.indexMessages(items.filter(([folder]) => gloda.GlodaMsgIndexer.shouldIndexFolder(folder)));
    index.queued += items.length;
  } catch (error) {
    srvNote("index new mail", error);
  }
}

function srvIndexListen() {
  const index = srvState.index;
  if (index.listener) {
    return;
  }
  index.listener = {
    msgAdded(hdr) {
      try {
        if (hdr && hdr.folder && hdr.folder.server && hdr.folder.server.type == "graph") {
          index.queue.push([hdr.folder, hdr.messageKey]);
          if (!index.flushTimer) {
            index.flushTimer = setTimeout(srvIndexFlush, SRV_INDEX_FLUSH_MS);
          }
        }
      } catch (error) {
        srvNote("new mail", error);
      }
    },
    QueryInterface: ChromeUtils.generateQI(["nsIMsgFolderListener"]),
  };
  MailServices.mfn.addListener(index.listener, Ci.nsIMsgFolderNotificationService.msgAdded);
}

function srvIndexStop() {
  const index = srvState.index;
  if (index.listener) {
    MailServices.mfn.removeListener(index.listener);
    index.listener = null;
  }
  clearTimeout(index.flushTimer);
  index.flushTimer = null;
  index.queue.length = 0;
}

// Marks the Graph folders whose message count changed (all with force) for
// Gloda's sweep. Returns the number of folders marked.
function srvIndexGraphMail(force) {
  srvIndexListen();
  const index = srvState.index;
  index.lastCheck = new Date().toISOString();
  const gloda = srvGloda();
  if (!gloda) {
    return 0;
  }
  let dirtied = 0;
  for (const folder of srvGraphFolders()) {
    try {
      if (!gloda.GlodaMsgIndexer.shouldIndexFolder(folder)) {
        continue;
      }
      const count = folder.getTotalMessages(false);
      if (!force && index.counts.get(folder.URI) === count) {
        continue;
      }
      index.counts.set(folder.URI, count);
      const glodaFolder = gloda.GlodaDatastore._mapFolder(folder);
      if (glodaFolder) {
        glodaFolder._ensureFolderDirty();
        dirtied++;
      }
    } catch (error) {
      srvNote(`index ${folder.prettyName}`, error);
    }
  }
  if (dirtied) {
    gloda.GlodaMsgIndexer.indexingSweepNeeded = true;
    index.dirtied += dirtied;
  }
  return dirtied;
}

this.outlookServer = class extends ExtensionCommon.ExtensionAPI {
  onShutdown(isAppShutdown) {
    if (isAppShutdown) {
      return;
    }
    srvUnwrap();
    srvIndexStop();
    Services.obs.notifyObservers(null, "startupcache-invalidate");
  }

  getAPI(context) {
    return {
      outlookServer: {
        async start() {
          srvWrap();
          return true;
        },
        async setQuiet(accountId, quiet) {
          if (quiet) {
            srvState.quiet.add(String(accountId));
          } else {
            srvState.quiet.delete(String(accountId));
          }
        },
        async syncFolderList(accountId) {
          return srvSyncFolderList(String(accountId));
        },
        async getNewMail(accountId) {
          const server = srvServerOf(String(accountId));
          const window = srvMsgWindow();
          if (!server || !window || Services.io.offline) {
            return false;
          }
          server.getNewMessages(server.rootFolder, window, null);
          return true;
        },
        async indexGraphMail(force) {
          return srvIndexGraphMail(Boolean(force));
        },
        async status() {
          return {
            index: {
              gloda: Boolean(srvGloda()),
              queued: srvState.index.queued,
              dirtied: srvState.index.dirtied,
              lastCheck: srvState.index.lastCheck,
            },
            watching: Boolean(srvState.wrapper),
            quiet: [...srvState.quiet],
            seen: srvState.seen,
            held: srvState.held,
            lastError: srvState.lastError,
          };
        },
        onServerError: new ExtensionCommon.EventManager({
          context,
          name: "outlookServer.onServerError",
          register: (fire) => {
            srvState.listeners.add(fire);
            return () => srvState.listeners.delete(fire);
          },
        }).api(),
      },
    };
  }
};
