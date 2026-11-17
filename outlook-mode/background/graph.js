"use strict";

// Microsoft Graph requests with the access token Thunderbird uses for the
// account (from the outlookPin experiment). Used by pin-sync.js,
// settings-sync.js and cleanup.js.
//
// Also keeps track of whether Microsoft's server is failing, per account:
// - signs of failure: Thunderbird's own "Server graph.microsoft.com responded
//   with an error." for a Graph account (experiment outlookServer), and HTTP
//   500/502/503/504 answers to this add-on's own requests. A lost network on
//   this computer is not counted (that is not Microsoft's server);
// - each sign is checked by running Thunderbird's folder list sync of the
//   account once (outlookServer.syncFolderList). If it fails too, the server
//   is "failing": one notice says so, Thunderbird's repeated error notices
//   for the account are held back, and the add-on's Graph requests wait;
// - while failing, the folder list sync is run again after 30 s, 1, 2, then
//   every 5 minutes. When it works again, one notice says the server is back,
//   Thunderbird gets the account's new mail, and the waiting requests and
//   syncs go on.

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const GRAPH_FAILING_STATUS = new Set([500, 502, 503, 504]);
// let: the tests make these short.
let SERVER_RECHECK_MS = [30000, 60000, 120000, 300000]; // then the last one again
let SERVER_CHECK_GAP_MS = 120000; // at most one check per account in this time while all is well
let GRAPH_WAIT_MS = 15 * 60000; // longest wait of a request for the server to come back
let GRAPH_THROTTLE_MAX_MS = 120000; // longest wait asked by a 429 answer

async function graphAccountIds() {
  if (!messenger.outlookPin || !messenger.outlookPin.listGraphAccounts) {
    return [];
  }
  return messenger.outlookPin.listGraphAccounts();
}

/* ------------------------------------------------------------------ */
/* Server health                                                       */
/* ------------------------------------------------------------------ */

// accountId -> { state: "ok" | "failing", since, reason, lastCheck,
//                checking: Promise | null, timer, attempts, waiters: Set }
const serverHealth = new Map();
const serverHealthLog = []; // for Diagnostics: the last changes
// Functions (accountId) => ... run when an account's server is back: the
// syncs that were put off (pin-sync.js, settings-sync.js add theirs).
const serverBackHooks = [];

function healthOf(accountId) {
  let health = serverHealth.get(accountId);
  if (!health) {
    health = { state: "ok", since: 0, reason: "", lastCheck: 0, checking: null, timer: null, attempts: 0, waiters: new Set() };
    serverHealth.set(accountId, health);
  }
  return health;
}

function serverFailing(accountId) {
  return healthOf(accountId).state === "failing";
}

async function anyServerFailing() {
  return (await graphAccountIds()).some((id) => serverFailing(id));
}

function healthNote(accountId, text) {
  serverHealthLog.push({ at: new Date().toISOString(), account: accountId, text });
  if (serverHealthLog.length > 20) {
    serverHealthLog.shift();
  }
}

async function accountLabel(accountId) {
  try {
    const account = await messenger.accounts.get(accountId, false);
    return (account && (account.identities?.[0]?.email || account.name)) || accountId;
  } catch (error) {
    return accountId;
  }
}

function minutesSince(time) {
  const minutes = Math.max(1, Math.round((Date.now() - time) / 60000));
  return minutes < 60 ? `${minutes} min` : `${Math.floor(minutes / 60)} h ${minutes % 60} min`;
}

// One run of the account's folder list sync: { ok, reason, network }.
// Without the experiment there is nothing to check with: not counted as
// Microsoft's server failing.
async function runServerCheck(accountId) {
  if (messenger.outlookServer && messenger.outlookServer.syncFolderList) {
    return messenger.outlookServer.syncFolderList(accountId);
  }
  return { ok: false, network: true, reason: "cannot check the server" };
}

function scheduleRecheck(accountId) {
  const health = healthOf(accountId);
  clearTimeout(health.timer);
  const delay = SERVER_RECHECK_MS[Math.min(health.attempts, SERVER_RECHECK_MS.length - 1)];
  health.attempts++;
  health.timer = setTimeout(() => {
    health.timer = null;
    checkServer(accountId, { force: true }).catch((error) => logError("server check", error));
  }, delay);
}

async function markFailing(accountId, reason) {
  const health = healthOf(accountId);
  health.reason = reason || health.reason;
  if (health.state === "failing") {
    return;
  }
  health.state = "failing";
  health.since = Date.now();
  health.attempts = 0;
  healthNote(accountId, `failing: ${health.reason}`);
  await messenger.outlookServer?.setQuiet(accountId, true).catch((error) => logError("server quiet", error));
  const who = await accountLabel(accountId);
  notify(
    "Microsoft's server is failing",
    `Mail for ${who} cannot sync: ${(health.reason || "the server answers with errors").replace(/\.$/, "")}. ` +
      "Thunderbird keeps trying; you will be told when it is back."
  );
  showServerState().catch(() => {});
}

async function markBack(accountId) {
  const health = healthOf(accountId);
  clearTimeout(health.timer);
  health.timer = null;
  health.attempts = 0;
  if (health.state !== "failing") {
    return;
  }
  const since = health.since;
  health.state = "ok";
  health.reason = "";
  healthNote(accountId, `back after ${minutesSince(since)}`);
  await messenger.outlookServer?.setQuiet(accountId, false).catch((error) => logError("server quiet", error));
  for (const resolve of health.waiters) {
    resolve();
  }
  health.waiters.clear();
  const who = await accountLabel(accountId);
  notify("Microsoft's server is back", `Mail for ${who} syncs again (it was failing for ${minutesSince(since)}).`);
  // Catch up: the account's new mail, then what the add-on put off.
  messenger.outlookServer?.getNewMail(accountId).catch((error) => logError("new mail after outage", error));
  showServerState().catch(() => {});
  for (const hook of serverBackHooks) {
    Promise.resolve()
      .then(() => hook(accountId))
      .catch((error) => logError("after outage", error));
  }
}

// Checks the server of an account (one check at a time). force: also within
// SERVER_CHECK_GAP_MS of the last one. Returns the result of the check.
function checkServer(accountId, { force = false, reason = "" } = {}) {
  const health = healthOf(accountId);
  if (health.checking) {
    return health.checking;
  }
  if (!force && health.state === "ok" && Date.now() - health.lastCheck < SERVER_CHECK_GAP_MS) {
    return Promise.resolve({ ok: true, skipped: true });
  }
  health.lastCheck = Date.now();
  health.checking = (async () => {
    let result;
    try {
      result = await runServerCheck(accountId);
    } catch (error) {
      result = { ok: false, network: true, reason: String(error && error.message ? error.message : error) };
    }
    if (result.ok) {
      await markBack(accountId);
    } else if (result.network && health.state !== "failing") {
      // Not Microsoft's server: this computer's connection (or Thunderbird
      // is offline). Nothing to tell.
    } else {
      await markFailing(accountId, result.reason || reason);
      scheduleRecheck(accountId);
    }
    return result;
  })().finally(() => {
    health.checking = null;
  });
  return health.checking;
}

// A sign that the server may be failing (an error notice of Thunderbird,
// or a 5xx answer to the add-on).
function noteServerFailure(accountId, reason, { force = false } = {}) {
  const health = healthOf(accountId);
  if (health.state === "failing") {
    health.reason = reason || health.reason;
    return Promise.resolve({ ok: false });
  }
  return checkServer(accountId, { force, reason });
}

// Resolves when the account's server is not known to be failing; rejects
// after `max` ms of waiting.
function waitForServer(accountId, max = GRAPH_WAIT_MS) {
  const health = healthOf(accountId);
  const pending = health.checking || Promise.resolve();
  return pending.then(() => {
    if (health.state !== "failing") {
      return undefined;
    }
    return new Promise((resolve, reject) => {
      const done = () => {
        clearTimeout(timer);
        resolve();
      };
      const timer = setTimeout(() => {
        health.waiters.delete(done);
        reject(userError(`Microsoft's server is still failing (${health.reason || "no answer"}).`));
      }, max);
      health.waiters.add(done);
    });
  });
}

// The ⟳ button's tooltip and colour while a server is failing.
async function showServerState() {
  if (!messenger.outlookPin || !messenger.outlookPin.setRefreshState) {
    return;
  }
  const failing = [...serverHealth.entries()].filter(([, health]) => health.state === "failing");
  if (failing.length) {
    const [, health] = failing[0];
    const time = new Date(health.since).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    await messenger.outlookPin.setRefreshState({
      busy: false,
      error: true,
      title: `Microsoft's server is failing since ${time} (${health.reason}). Thunderbird keeps trying.\nClick to try now.`,
    });
  } else {
    const time = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    await messenger.outlookPin.setRefreshState({
      busy: false,
      error: false,
      title: `Microsoft's server is back (${time}).\nClick to refresh.`,
    });
  }
}

function serverHealthStatus() {
  return {
    accounts: [...serverHealth.entries()].map(([account, health]) => ({
      account,
      state: health.state,
      since: health.since ? new Date(health.since).toISOString() : "",
      reason: health.reason,
      lastCheck: health.lastCheck ? new Date(health.lastCheck).toISOString() : "",
    })),
    log: [...serverHealthLog],
  };
}

if (messenger.outlookServer && messenger.outlookServer.onServerError) {
  messenger.outlookServer.onServerError.addListener((accountId, message) => {
    noteServerFailure(accountId, message).catch((error) => logError("server check", error));
  });
  messenger.outlookServer.start().catch((error) => logError("server watch", error));
}

/* ------------------------------------------------------------------ */
/* Requests                                                            */
/* ------------------------------------------------------------------ */

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// options.wait: false for requests a person is looking at (cleanup page):
// they fail at once while the server is failing instead of waiting.
async function graphFetch(accountId, pathOrUrl, init = {}, { wait = true } = {}) {
  // Paging links come back as full URLs; the token is only sent to Graph.
  const url = pathOrUrl.startsWith("https://") ? pathOrUrl : GRAPH_BASE + pathOrUrl;
  if (!url.startsWith(`${GRAPH_BASE}/`)) {
    throw new Error("Refused a request outside Microsoft Graph.");
  }
  const method = String(init.method || "GET").toUpperCase();
  for (let attempt = 0; ; attempt++) {
    if (serverFailing(accountId)) {
      if (!wait) {
        throw userError("Microsoft's server is failing at the moment; try again when it is back.");
      }
      await waitForServer(accountId);
    }
    const token = await messenger.outlookPin.getGraphToken(accountId);
    const response = await fetch(url, {
      ...init,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...(init.headers || {}),
      },
    });
    if (response.status === 429 && attempt === 0) {
      // Throttled: wait as long as Microsoft asks (bounded), then once more.
      const seconds = Number((response.headers && response.headers.get && response.headers.get("Retry-After")) || 10);
      await sleep(Math.min(Math.max(seconds, 1) * 1000, GRAPH_THROTTLE_MAX_MS));
      continue;
    }
    if (GRAPH_FAILING_STATUS.has(response.status)) {
      await noteServerFailure(accountId, `Microsoft Graph answered ${response.status}`, { force: true });
      // A POST may have been done before the error; it is not sent again.
      if (wait && attempt === 0 && method !== "POST") {
        await waitForServer(accountId);
        continue;
      }
    }
    if (!response.ok) {
      let detail = "";
      try {
        detail = (await response.json()).error.message;
      } catch (error) {
        // no JSON error body
      }
      throw new Error(`Graph ${response.status}${detail ? `: ${detail}` : ""}`);
    }
    return response.status === 204 ? null : response.json();
  }
}

// OData string literal: single quotes are doubled.
function odataString(text) {
  return `'${String(text).replace(/'/g, "''")}'`;
}
