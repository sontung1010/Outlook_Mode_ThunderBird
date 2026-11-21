"use strict";

const DEFAULT_SETTINGS = {
  replyHeader: true,
  removeQuoteBar: true,
  readingPaneNewestFirst: true,
  readingPaneExpandAll: true,
  readingPaneFoldQuotes: true,
  readingPaneAttachmentsTop: true,
  readingPaneLatestOnly: true,
  readingPaneOutlookHeader: true,
  readingPaneLight: true,
  inlineReply: true,
  inlineNewMessage: true,
  pinnedOnTop: true,
  pinSync: true,
  pinButton: true,
  refreshButton: true,
  unreadHighlight: true,
  markConversationRead: true,
  oneCardPerConversation: true,
  pinnedSection: true,
  dateSections: true,
  newestInCard: true,
  outlookSearch: true,
  collectContacts: true,
  contactCompanyFormat: true,
  refreshKey: true,
  undoSendSeconds: 10,
  singleSentCopy: true,
  mentionTagging: true,
  settingsSync: false,
  syncCategories: { prefs: true, accounts: true, filters: true, contacts: true, addon: true },
};
const CHECKBOXES = [
  "replyHeader",
  "removeQuoteBar",
  "readingPaneNewestFirst",
  "readingPaneExpandAll",
  "readingPaneFoldQuotes",
  "readingPaneAttachmentsTop",
  "readingPaneLatestOnly",
  "readingPaneOutlookHeader",
  "readingPaneLight",
  "inlineReply",
  "inlineNewMessage",
  "pinnedOnTop",
  "pinSync",
  "pinButton",
  "refreshButton",
  "refreshKey",
  "unreadHighlight",
  "markConversationRead",
  "oneCardPerConversation",
  "pinnedSection",
  "dateSections",
  "newestInCard",
  "outlookSearch",
  "collectContacts",
  "contactCompanyFormat",
  "mentionTagging",
  "singleSentCopy",
  "settingsSync",
];
const SYNC_CATEGORIES = ["prefs", "accounts", "filters", "contacts", "addon"];

// Outlook shortcut -> how to do it in Thunderbird. `command` rows show the
// add-on's current key; the others are Thunderbird's own keys.
const SHORTCUTS = [
  { action: "New message", outlook: "Ctrl+N / Ctrl+Shift+M", tb: "Ctrl+N or Ctrl+M" },
  { action: "Reply", outlook: "Ctrl+R", tb: "Ctrl+R" },
  { action: "Reply all", outlook: "Ctrl+Shift+R", tb: "Ctrl+Shift+R" },
  { action: "Forward", outlook: "Ctrl+F", tb: "Ctrl+L" },
  { action: "Forward as attachment", outlook: "Ctrl+Alt+F", command: "ol-forward-attachment" },
  { action: "Send", outlook: "Ctrl+Enter / Alt+S", tb: "Ctrl+Enter" },
  { action: "Delete", outlook: "Ctrl+D / Delete", command: "ol-delete", extra: "or Delete" },
  { action: "Mark as read", outlook: "Ctrl+Q", command: "ol-mark-read", extra: "or M (toggle)" },
  { action: "Mark as unread", outlook: "Ctrl+U", command: "ol-mark-unread", extra: "or M (toggle)" },
  { action: "Move to folder", outlook: "Ctrl+Shift+V", command: "ol-move" },
  { action: "Pin / flag", outlook: "Insert", command: "toggle-pin", extra: "or S (star)" },
  { action: "Show @mentions", outlook: "(@ filter)", command: "ol-show-mentions" },
  { action: "Quick Steps 1–5", outlook: "Ctrl+Shift+1…9", command: "quickstep-1", extra: "…5" },
  { action: "Search", outlook: "Ctrl+E", tb: "Ctrl+K (quick filter) / Ctrl+Shift+K" },
  { action: "Address book", outlook: "Ctrl+Shift+B", tb: "Ctrl+Shift+B" },
  { action: "New appointment / task", outlook: "Ctrl+Shift+A / Ctrl+Shift+K", tb: "Ctrl+I / Ctrl+D (Ctrl+D is Delete while this add-on is active)" },
  { action: "Go to mail", outlook: "Ctrl+1", tb: "Alt+1" },
];

let folders = [];

async function loadFolders() {
  const accounts = await messenger.accounts.list(true);
  const out = [];
  const walk = (list, prefix) => {
    for (const folder of list || []) {
      out.push({ id: folder.id, label: `${prefix}${folder.name}` });
      walk(folder.subFolders, `${prefix}${folder.name} / `);
    }
  };
  for (const account of accounts) {
    const top = (account.rootFolder && account.rootFolder.subFolders) || account.folders || [];
    walk(top, `${account.name} / `);
  }
  return out.filter((f) => f.id);
}

function el(tag, props = {}, ...children) {
  const node = document.createElement(tag);
  Object.assign(node, props);
  node.append(...children);
  return node;
}

function option(value, text, selected) {
  const node = el("option", { value, textContent: text });
  node.selected = value === selected;
  return node;
}

/* ------------------------------------------------------------------ */
/* Quick Steps editor                                                  */
/* ------------------------------------------------------------------ */

function addRow(step) {
  const tr = document.createElement("tr");
  const name = el("input", { type: "text", value: step.name || "", placeholder: "e.g. Done: read + archive" });
  name.dataset.field = "name";
  const markRead = el("input", { type: "checkbox", checked: Boolean(step.markRead) });
  markRead.dataset.field = "markRead";
  const pin = el("input", { type: "checkbox", checked: Boolean(step.pin) });
  pin.dataset.field = "pin";
  const compose = el("select");
  compose.dataset.field = "compose";
  compose.append(
    option("none", "—", step.compose || "none"),
    option("reply", "Reply", step.compose),
    option("replyAll", "Reply all", step.compose),
    option("forward", "Forward", step.compose)
  );
  const forwardTo = el("input", { type: "text", value: step.forwardTo || "", placeholder: "a@x.com; b@y.com" });
  forwardTo.dataset.field = "forwardTo";
  const after = el("select");
  after.dataset.field = "after";
  after.append(
    option("none", "—", step.after || "none"),
    option("move", "Move to", step.after),
    option("archive", "Archive", step.after),
    option("delete", "Delete", step.after)
  );
  const moveTo = el("select");
  moveTo.dataset.field = "moveTo";
  moveTo.append(option("", "(choose folder)", step.moveTo || ""));
  for (const folder of folders) {
    moveTo.append(option(folder.id, folder.label, step.moveTo));
  }
  const syncEnabled = () => {
    forwardTo.disabled = compose.value !== "forward";
    moveTo.disabled = after.value !== "move";
  };
  compose.addEventListener("change", syncEnabled);
  after.addEventListener("change", syncEnabled);
  syncEnabled();
  const remove = el("button", { className: "danger", textContent: "✕", title: "Remove step" });
  remove.addEventListener("click", () => tr.remove());

  [name, markRead, pin, compose, forwardTo, after, moveTo, remove].forEach((node, index) => {
    const td = el("td", {}, node);
    if ([1, 2, 7].includes(index)) {
      td.className = "narrow";
    }
    tr.appendChild(td);
  });
  tr.dataset.id = step.id || `step-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
  document.getElementById("steps").appendChild(tr);
}

function readSteps() {
  return [...document.querySelectorAll("#steps tr")].map((tr) => {
    const get = (field) => tr.querySelector(`[data-field="${field}"]`);
    return {
      id: tr.dataset.id,
      name: get("name").value.trim(),
      markRead: get("markRead").checked,
      pin: get("pin").checked,
      compose: get("compose").value,
      forwardTo: get("forwardTo").value.trim(),
      after: get("after").value,
      moveTo: get("moveTo").value,
    };
  });
}

/* ------------------------------------------------------------------ */
/* Scheduled mail                                                      */
/* ------------------------------------------------------------------ */

async function renderScheduled() {
  const body = document.getElementById("scheduled");
  const list = await messenger.runtime.sendMessage({ type: "schedule-list" });
  body.textContent = "";
  if (!list.length) {
    body.append(el("tr", {}, el("td", { colSpan: 5, className: "hint", textContent: "Nothing scheduled." })));
    return;
  }
  list.sort((a, b) => new Date(a.sendAt) - new Date(b.sendAt));
  for (const item of list) {
    const sendNow = el("button", { className: "small", textContent: item.status === "failed" ? "Retry" : "Send now" });
    sendNow.addEventListener("click", async () => {
      sendNow.disabled = true;
      await messenger.runtime.sendMessage({ type: "schedule-send-now", id: item.id }).catch(() => {});
      renderScheduled();
    });
    const cancel = el("button", { className: "small danger", textContent: "Cancel", title: "Stop the schedule; the message stays in Drafts" });
    cancel.addEventListener("click", async () => {
      await messenger.runtime.sendMessage({ type: "schedule-cancel", id: item.id });
      renderScheduled();
    });
    const status = el("td", {
      className: item.status === "failed" ? "failed" : "",
      textContent: item.status === "failed" ? `Failed: ${item.error}` : "Waiting",
    });
    body.append(
      el(
        "tr",
        {},
        el("td", { textContent: new Date(item.sendAt).toLocaleString() }),
        el("td", { textContent: item.subject }),
        el("td", { textContent: item.to }),
        status,
        el("td", { className: "narrow" }, sendNow, cancel)
      )
    );
  }
}

/* ------------------------------------------------------------------ */
/* Shortcuts                                                           */
/* ------------------------------------------------------------------ */

async function renderShortcuts() {
  const commands = await messenger.commands.getAll();
  const byName = new Map(commands.map((c) => [c.name, c.shortcut]));
  const body = document.getElementById("shortcuts");
  body.textContent = "";
  for (const row of SHORTCUTS) {
    const tb = el("td");
    if (row.command) {
      const key = byName.get(row.command);
      tb.append(key ? el("kbd", { textContent: key }) : el("span", { className: "hint", textContent: "(not set)" }));
      if (row.extra) {
        tb.append(` ${row.extra}`);
      }
    } else {
      tb.textContent = row.tb;
    }
    body.append(el("tr", {}, el("td", { textContent: row.action }), el("td", { textContent: row.outlook }), tb));
  }
}

/* ------------------------------------------------------------------ */
/* Diagnostics                                                         */
/* ------------------------------------------------------------------ */

function pinSyncLines(status) {
  const accounts = Object.entries(status || {});
  if (!accounts.length) {
    return ["Pin sync: no Microsoft Graph account synced yet"];
  }
  return accounts.map(
    ([id, s]) =>
      `Pin sync ${id}: last sync ${s.lastSync ? new Date(s.lastSync).toLocaleString() : "never"}, ` +
      `${s.serverPins === null ? "?" : s.serverPins} pinned on server` +
      (s.method === "scan" ? " (read from the newest Inbox mail)" : "") +
      `, ${s.pending} waiting to upload` +
      (s.lastError ? ` (${s.lastError})` : "")
  );
}

// Sent Items on Graph accounts: sends that filed no second copy, and the
// second copies moved to Deleted Items.
function sentCopyLines(accounts, pane) {
  const lines = [
    `One copy of sent mail: ${pane && pane.singleSentCopy ? "on" : "off"}, ` +
      `${(pane && pane.sentCopiesSkipped) || 0} send(s) since start filed no second copy`,
  ];
  for (const [id, s] of Object.entries(accounts || {})) {
    lines.push(
      `  Sent Items ${id}: last check ${s.lastCheck ? new Date(s.lastCheck).toLocaleString() : "never"}, ` +
        `${s.scanned} read, ${s.moved} second cop${s.moved === 1 ? "y" : "ies"} moved to Deleted Items since start` +
        (s.lastError ? ` (${s.lastError})` : "")
    );
  }
  return lines;
}

async function checkPins() {
  const pre = document.getElementById("pin-check");
  pre.textContent = "Checking…";
  try {
    const rows = await messenger.runtime.sendMessage({ type: "pin-sync-check" });
    pre.textContent = rows
      .map((row) => {
        const head = `${row.subject || "(no subject)"} — local 📌: ${row.localPinned ? "yes" : "no"}`;
        if (row.error) {
          return `${head}\n  ${row.error}`;
        }
        if (!row.server.length) {
          return `${head}\n  not found on the server`;
        }
        return [head, ...row.server.map((props, i) => `  server copy ${i + 1}: ${JSON.stringify(props)}`)].join("\n");
      })
      .join("\n\n");
  } catch (error) {
    pre.textContent = `Failed: ${error.message || error}`;
  }
}

async function renderSyncStatus() {
  const pre = document.getElementById("sync-status");
  try {
    const s = await messenger.runtime.sendMessage({ type: "settings-sync-status" });
    const lines = [
      s.lastSync ? `Last sync: ${new Date(s.lastSync).toLocaleString()} (${s.keys} settings, ${Math.round(s.bytes / 1024)} KB)` : "Not synced yet on this device.",
    ];
    if (s.lastApplied) {
      lines.push(`Last sync applied ${s.lastApplied} setting(s) from other devices.`);
    }
    const failed = Object.entries(s.lastFailed || {});
    if (failed.length) {
      lines.push(`Could not apply ${failed.length}:`, ...failed.slice(0, 10).map(([key, why]) => `  ${key}: ${why}`));
    }
    if (s.lastError) {
      lines.push(`Error: ${s.lastError}`);
    }
    if (s.backup) {
      lines.push(`Undo is available for the sync of ${new Date(s.backup).toLocaleString()}.`);
    }
    pre.textContent = lines.join("\n");
    document.getElementById("sync-undo").disabled = !s.backup;
  } catch (error) {
    pre.textContent = `Background script not reachable: ${error.message || error}`;
  }
}

// Scrolls to a section (the popup's ✍ Signature, the composer's Signature
// menu) and marks it for a moment.
function showSection(section) {
  const heading = section && document.getElementById(section);
  if (!heading) {
    return;
  }
  heading.scrollIntoView({ block: "start" });
  heading.classList.add("flash");
  setTimeout(() => heading.classList.remove("flash"), 1500);
}

messenger.runtime.onMessage.addListener((message) => {
  if (message && message.type === "show-settings-section") {
    showSection(message.section);
  }
  return undefined;
});

async function renderCollectStatus() {
  const pre = document.getElementById("collect-status");
  try {
    const s = await messenger.runtime.sendMessage({ type: "collect-contacts-status" });
    pre.textContent = s.lastRun
      ? `Last run: ${new Date(s.lastRun).toLocaleString()}: ${s.lastScanned} messages checked, ${s.people} people found, ${s.lastAdded} added, ${s.lastMerged || 0} duplicate(s) merged.`
      : "Not run yet on this device (it runs a minute after Thunderbird starts, then every 6 hours).";
    if (s.lastError) {
      pre.textContent += `\nError: ${s.lastError}`;
    }
  } catch (error) {
    pre.textContent = "";
  }
}

async function runSyncAction(type, button) {
  const pre = document.getElementById("sync-status");
  button.disabled = true;
  pre.textContent = "Syncing…";
  try {
    const result = await messenger.runtime.sendMessage({ type });
    if (result && result.ok === false && result.error) {
      pre.textContent = result.error;
      button.disabled = false;
      return;
    }
  } catch (error) {
    // shown by renderSyncStatus as lastError, or below
    pre.textContent = `Failed: ${error.message || error}`;
  }
  button.disabled = false;
  await renderSyncStatus();
}

function paneLines(status) {
  if (!status) {
    return [];
  }
  if (status.error) {
    return [`Reading pane: error: ${status.error}`];
  }
  const lines = [
    status.conversations
      ? `Reading pane: Thunderbird Conversations ${status.conversationsVersion}, active in ${status.panes} view(s), ${status.drafts} unsent draft(s), ${status.pendingSends} in Undo Send`
      : "Reading pane: Thunderbird Conversations is not installed (needed for the inline reply and newest-on-top)",
    `New messages: ${status.composeRoutedWindows ? "open in the compose dock" : "open in a compose window"}, ${status.newMessages || 0} open`,
  ];
  if (status.lastError) {
    lines.push(`  last error: ${status.lastError}`);
  }
  return lines;
}

async function renderPaneStatus() {
  const pre = document.getElementById("pane-status");
  try {
    const status = await messenger.runtime.sendMessage({ type: "reading-pane-status" });
    pre.textContent = paneLines(status).join("\n");
    // The add-on does not install Thunderbird Conversations; say what to do.
    const advice = document.getElementById("conversations-advice");
    const state = await messenger.runtime.sendMessage({ type: "conversations-state" });
    advice.textContent = state.advice;
    advice.hidden = !state.advice;
  } catch (error) {
    pre.textContent = "";
  }
}

// Microsoft's server (Graph accounts): failing or not, and the last changes.
function serverLines(server) {
  if (!server) {
    return [];
  }
  const watch = server.watch || {};
  const lines = [
    `Microsoft's server: ${
      server.accounts.filter((a) => a.state === "failing").map((a) => `${a.account} failing since ${a.since} (${a.reason})`).join("; ") ||
      "no failure known"
    }; Thunderbird error notices seen ${watch.seen ?? "?"}, held back ${watch.held ?? "?"}` +
      (watch.lastError ? ` (error: ${watch.lastError})` : ""),
  ];
  if (watch.index) {
    lines.push(
      `Message index (Gloda) for Graph folders: ${watch.index.gloda ? "on" : "OFF in Thunderbird"}, ` +
        `${watch.index.queued} new mail(s) and ${watch.index.dirtied} folder check(s) given to it since start` +
        (watch.index.lastCheck ? `, last check ${new Date(watch.index.lastCheck).toLocaleString()}` : "")
    );
  }
  for (const entry of server.log.slice(-5)) {
    lines.push(`  ${entry.at}  ${entry.account}: ${entry.text}`);
  }
  return lines;
}

async function renderDiagnostics() {
  const pre = document.getElementById("diag");
  try {
    const d = await messenger.runtime.sendMessage({ type: "diagnostics" });
    const lines = [
      `Add-on ${d.version} on ${d.thunderbird}`,
      `Compose script registered: ${d.boot.composeScriptRegistered ? "yes" : `NO (${d.boot.composeScriptError})`}`,
      `Injected into already open compose windows: ${d.boot.injectedIntoOpenTabs}` +
        (d.boot.injectErrors.length ? ` (errors: ${d.boot.injectErrors.join("; ")})` : ""),
      `Compose windows where the script started: ${d.composeWindowsSeen}`,
      `Contacts for @mention: ${d.contacts.contacts} from ${d.contacts.books} address book(s)` +
        (d.contacts.error ? ` (error: ${d.contacts.error})` : ""),
      ...pinSyncLines(d.pinSync),
      ...sentCopyLines(d.sentCopies, d.readingPane),
      ...paneLines(d.readingPane),
      d.markRead.last
        ? `Mark conversation read: last at ${d.markRead.last}, ${d.markRead.selected} mail(s) (${d.markRead.found} found) in ${d.markRead.rows} row(s), ` +
          `${d.markRead.unreadBefore} unread, ${d.markRead.marked} marked read` +
          (d.markRead.reason ? ` (${d.markRead.reason})` : "") +
          `; card now: ${d.markRead.card || "?"}`
        : "Mark conversation read: not run yet",
      `Thread unread counts set right since start: ${d.markRead.threadCountsRepaired ?? "?"}`,
      ...serverLines(d.server),
      "",
      d.errors.length ? "Recent errors:" : "No errors recorded.",
      ...d.errors.map((e) => `  ${e.at}  ${e.where}: ${e.text}`),
    ];
    pre.textContent = lines.join("\n");
  } catch (error) {
    pre.textContent = `Background script not reachable: ${error.message || error}`;
  }
}

/* ------------------------------------------------------------------ */
/* Init                                                                */
/* ------------------------------------------------------------------ */

async function init() {
  document.getElementById("version").textContent = `v${messenger.runtime.getManifest().version}`;
  try {
    folders = await loadFolders();
  } catch (error) {
    console.error("[outlook-mode] could not list folders", error);
  }
  const stored = await messenger.storage.local.get(["settings", "quickSteps"]);
  const settings = { ...DEFAULT_SETTINGS, ...(stored.settings || {}) };
  for (const key of CHECKBOXES) {
    document.getElementById(key).checked = Boolean(settings[key]);
  }
  // Search on/off is kept by Thunderbird (not in the add-on's storage).
  try {
    document.getElementById("outlookSearch").checked = (await messenger.runtime.sendMessage({ type: "search-status" })).enabled;
  } catch (error) {
    // keep the stored value
  }
  for (const category of SYNC_CATEGORIES) {
    document.getElementById(`sync-${category}`).checked = settings.syncCategories[category] !== false;
  }
  document.getElementById("undoSendSeconds").value = String(settings.undoSendSeconds);
  const steps = Array.isArray(stored.quickSteps)
    ? stored.quickSteps
    : await messenger.runtime.sendMessage({ type: "get-steps" });
  steps.forEach(addRow);

  document.getElementById("add").addEventListener("click", () =>
    addRow({ name: "", markRead: true, compose: "none", after: "none" })
  );
  document.getElementById("save").addEventListener("click", async () => {
    // 0 to 60 seconds (the background clamps the stored value the same way).
    const undoSeconds = Math.floor(Number(document.getElementById("undoSendSeconds").value));
    const newSettings = { undoSendSeconds: Number.isFinite(undoSeconds) ? Math.min(Math.max(undoSeconds, 0), 60) : 0 };
    for (const key of CHECKBOXES) {
      newSettings[key] = document.getElementById(key).checked;
    }
    newSettings.syncCategories = Object.fromEntries(
      SYNC_CATEGORIES.map((category) => [category, document.getElementById(`sync-${category}`).checked])
    );
    const newSteps = readSteps().filter((s) => s.name);
    const bad = newSteps.find((s) => s.after === "move" && !s.moveTo);
    const saved = document.getElementById("saved");
    if (bad) {
      saved.style.color = "#b3261e";
      saved.textContent = `"${bad.name}": choose a folder to move to.`;
      return;
    }
    await messenger.storage.local.set({ settings: newSettings, quickSteps: newSteps });
    await messenger.runtime.sendMessage({ type: "search-set", enabled: newSettings.outlookSearch });
    saved.style.color = "#2e7d32";
    saved.textContent = "Saved.";
    setTimeout(() => (saved.textContent = ""), 2000);
  });

  document.getElementById("scan").addEventListener("click", async (event) => {
    const status = document.getElementById("scan-status");
    event.target.disabled = true;
    status.textContent = "Scanning…";
    try {
      const r = await messenger.runtime.sendMessage({ type: "mentions-scan", days: 30 });
      status.textContent = `Checked ${r.scanned} messages, tagged ${r.tagged}.`;
    } catch (error) {
      status.textContent = `Failed: ${error.message || error}`;
    }
    event.target.disabled = false;
  });

  document.getElementById("shortcut-settings").addEventListener("click", () =>
    messenger.commands.openShortcutSettings()
  );
  document.getElementById("diag-refresh").addEventListener("click", renderDiagnostics);
  document.getElementById("pin-check-run").addEventListener("click", checkPins);
  document.getElementById("sync-now").addEventListener("click", (event) =>
    runSyncAction("settings-sync-now", event.target)
  );
  document.getElementById("sync-undo").addEventListener("click", (event) => {
    if (confirm("Put back the settings that the last sync replaced on this device?")) {
      runSyncAction("settings-sync-undo", event.target);
    }
  });
  const identitySelect = document.getElementById("signature-identity");
  for (const account of await messenger.accounts.list(true)) {
    for (const identity of account.identities || []) {
      if (identity.email) {
        identitySelect.append(new Option(identity.name ? `${identity.name} <${identity.email}>` : identity.email, identity.id));
      }
    }
  }
  // The signature as it will look (a sandboxed frame: no scripts, no access).
  const preview = document.getElementById("signature-preview");
  const renderSignature = async () => {
    const html = await messenger.runtime.sendMessage({ type: "get-signature", identityId: identitySelect.value }).catch(() => "");
    // The HTML is the sanitized signature (outlookPane.getSignature). The
    // frame is sandboxed without allow-scripts, and its own policy allows
    // no script, frame, form or connection either, and only style
    // attributes: no <style> sheet (so no CSS selectors on the content).
    preview.srcdoc =
      '<!DOCTYPE html><html><head><meta http-equiv="Content-Security-Policy" content="default-src \'none\'; img-src data: https:; style-src-attr \'unsafe-inline\'"></head>' +
      '<body style="margin:8px;background:#fff;color:#000;font:11pt Calibri, Arial, sans-serif">' +
      (html || '<i style="color:#888">No signature</i>') +
      "</body></html>";
  };
  identitySelect.addEventListener("change", renderSignature);
  renderSignature();
  const signatureAction = (id, type, describe) =>
    document.getElementById(id).addEventListener("click", async (event) => {
      const status = document.getElementById("signature-status");
      event.target.disabled = true;
      status.textContent = "";
      try {
        const result = await messenger.runtime.sendMessage({ type, identityId: identitySelect.value });
        if (!result.cancelled) {
          status.textContent = describe(result);
          renderSignature();
        }
      } catch (error) {
        status.textContent = `Could not set the signature: ${error.message || error}`;
      }
      event.target.disabled = false;
    });
  signatureAction(
    "signature-from-mail",
    "signature-from-mail",
    (r) => `Signature of ${r.email} taken from "${r.subject}", with ${r.images} picture(s).`
  );
  signatureAction(
    "import-signature",
    "import-signature",
    (r) =>
      `Signature of ${r.email} imported, with ${r.images} picture(s).` +
      (r.missing ? `\n${r.missing} picture(s) were not found next to the .htm file; copy its _files folder too and import again.` : "")
  );
  document.getElementById("open-cleanup").addEventListener("click", () =>
    messenger.runtime.sendMessage({ type: "cleanup-open" }).catch(() => {})
  );
  document.getElementById("apply-recommended").addEventListener("click", async (event) => {
    const status = document.getElementById("recommended-status");
    event.target.disabled = true;
    try {
      const result = await messenger.runtime.sendMessage({ type: "apply-recommended" });
      status.textContent = result.changed ? `Changed: ${result.summary}.` : "Already set.";
    } catch (error) {
      status.textContent = `Failed: ${error.message || error}`;
    }
    event.target.disabled = false;
  });
  document.getElementById("collect-now").addEventListener("click", async (event) => {
    const pre = document.getElementById("collect-status");
    event.target.disabled = true;
    pre.textContent = "Collecting… (this can take a minute with a large mailbox)";
    try {
      const result = await messenger.runtime.sendMessage({ type: "collect-contacts-now" });
      if (result && result.ok === false) {
        pre.textContent = `${result.error} Turn it on and click Save first.`;
      } else {
        await renderCollectStatus();
      }
    } catch (error) {
      pre.textContent = `Failed: ${error.message || error}`;
    }
    event.target.disabled = false;
  });
  document.getElementById("pin-sync-run").addEventListener("click", async (event) => {
    event.target.disabled = true;
    await messenger.runtime.sendMessage({ type: "pin-sync-now" }).catch(() => {});
    event.target.disabled = false;
    renderDiagnostics();
  });
  document.getElementById("reload-contacts").addEventListener("click", async () => {
    await messenger.runtime.sendMessage({ type: "reload-contacts" }).catch(() => {});
    renderDiagnostics();
  });

  messenger.storage.onChanged.addListener((changes) => {
    if (changes.scheduled) {
      renderScheduled();
    }
  });

  renderScheduled();
  renderSyncStatus();
  renderCollectStatus();
  const section =
    location.hash.slice(1) || (await messenger.runtime.sendMessage({ type: "take-settings-section" }).catch(() => ""));
  // After the lists above are filled, so the page has its final height.
  setTimeout(() => showSection(section), 300);
  renderPaneStatus();
  renderShortcuts();
  renderDiagnostics();
}

init();
