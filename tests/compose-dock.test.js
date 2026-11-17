"use strict";
// Run: NODE_PATH=<dir with jsdom>/node_modules node tests/compose-dock.test.js
// (jsdom: `npm install jsdom@24` in any temporary folder)
// Runs reading-pane/pane.js in a jsdom copy of reading-pane/dock.html (the
// compose dock for new messages) with a fake bridge and dock
// (experiments/outlookPane), and checks the main flows.
const fs = require("fs");
const path = require("path");
const assert = require("assert");
const { JSDOM } = require("jsdom");

const ROOT = path.join(__dirname, "../outlook-mode/reading-pane");
const dom = new JSDOM(fs.readFileSync(path.join(ROOT, "dock.html"), "utf8"), {
  runScripts: "outside-only",
  pretendToBeVisual: true,
});
const { window } = dom;
const { document } = window;
window.document.execCommand = () => true;
// jsdom has no blob: URLs.
let blobCount = 0;
window.URL.createObjectURL = () => `blob:test-${++blobCount}`;
window.URL.revokeObjectURL = () => {};
const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
window.Element.prototype.setPointerCapture = function () {};
let confirmAnswer = true;
const confirms = [];
// jsdom has no ResizeObserver (Thunderbird has); the pane only needs it to exist.
window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
window.confirm = (text) => {
  confirms.push(text);
  return confirmAnswer;
};

// Fake experiment: drafts, the dock's tab list and its size state.
const drafts = new Map();
const pending = new Map();
const sent = [];
let order = [];
let seq = 0;
let focusUri = null;
const dockState = { minimized: false, maximized: false, hidden: false, height: 300 };
const createNew = () => {
  const uri = `new:${++seq}`;
  drafts.set(uri, { type: "new", identityKey: "id1", to: "", cc: "", bcc: "", subject: "", html: null, attachments: [], files: [], dirty: false });
  order.push(uri);
  return uri;
};
const bridge = {
  config: () => ({ inlineNew: true, loadId: 1 }),
  sanitize: (html) => html, // the experiment runs Thunderbird's sanitizer
  uriForConvId: () => null,
  convIdForUri: () => null,
  prepare: async (uri, type) => {
    assert.equal(type, "new");
    return {
      uri,
      type,
      identityKey: "id1",
      identities: [
        { key: "id1", label: "Me <me@x.com>", signature: `<table cellpadding="4"><tbody><tr><td><img src="${PNG}" width="40"></td><td>Best,<br>Me</td></tr></tbody></table>` },
      ],
      to: "",
      cc: "",
      bcc: "",
      subject: "",
      attachments: [],
      original: null,
      font: "",
    };
  },
  searchAddresses: () => [],
  getDraft: (uri) => (drafts.has(uri) ? structuredClone(drafts.get(uri)) : null),
  saveDraft: (uri, d) => drafts.set(uri, structuredClone(d)),
  discardDraft: (uri) => drafts.delete(uri),
  send: async (uri, d) => {
    drafts.set(uri, { ...structuredClone(d), dirty: true }); // as paneSend does
    pending.set(uri, { until: Date.now() + 10000, draft: structuredClone(d) });
    return { pending: true, until: Date.now() + 10000 };
  },
  pendingSend: (uri) => (pending.has(uri) ? structuredClone(pending.get(uri)) : null),
  cancelSend: (uri) => {
    const p = pending.get(uri);
    pending.delete(uri);
    return p ? p.draft : null;
  },
  sendNow: async (uri) => {
    const p = pending.get(uri);
    pending.delete(uri);
    sent.push(p.draft);
    drafts.delete(uri);
    order = order.filter((u) => u !== uri); // the experiment forgets a sent message
    window.outlookModePane.onSendResult(uri, { ok: true, subject: p.draft.subject });
    return { sent: true };
  },
  popOut: (uri, d) => {
    window.__popOut = d;
    drafts.delete(uri);
    return true;
  },
  reportError: (where, text) => {
    throw new Error(`page error ${where}: ${text}`);
  },
  dock: {
    takeFocus: () => {
      const uri = focusUri;
      focusUri = null;
      return uri;
    },
    messages: () =>
      order.map((uri) => ({ uri, subject: drafts.get(uri)?.subject || "", dirty: Boolean(drafts.get(uri)?.dirty), sending: pending.has(uri) })),
    create: createNew,
    remove: (uri) => {
      if (!pending.has(uri)) {
        drafts.delete(uri);
      }
      order = order.filter((u) => u !== uri);
    },
    state: () => ({ minimized: dockState.minimized, maximized: dockState.maximized }),
    height: () => (dockState.minimized ? 34 : dockState.height),
    setMinimized: (on) => {
      dockState.minimized = on;
      if (on) dockState.maximized = false;
    },
    setMaximized: (on) => {
      dockState.maximized = on;
      if (on) dockState.minimized = false;
    },
    setHeight: (px) => {
      dockState.height = px;
      dockState.minimized = dockState.maximized = false;
    },
    hide: () => {
      dockState.hidden = true;
    },
  },
};
window.outlookModeBridge = bridge;

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const $ = (sel) => document.querySelector(sel);
const tabs = () => [...document.querySelectorAll(".ol-dock-tab")].map((t) => t.querySelector(".ol-dock-title").textContent + (t.classList.contains("ol-active") ? "*" : ""));
const button = (title) => [...document.querySelectorAll(".ol-dock-button")].find((b) => b.title === title);
const subjectInput = () => document.querySelector(".ol-row:not(.ol-hidden) input.ol-input");
const action = (text) => [...document.querySelectorAll(".ol-actions button")].find((b) => b.textContent === text);
const type = (input, text) => {
  input.value = text;
  input.dispatchEvent(new window.Event("input", { bubbles: true }));
};
const key = (target, k, opts = {}) =>
  target.dispatchEvent(new window.KeyboardEvent("keydown", { key: k, bubbles: true, cancelable: true, ...opts }));

(async () => {
  // 1. The experiment created a message, then the page started.
  focusUri = createNew();
  window.eval(fs.readFileSync(path.join(ROOT, "pane.js"), "utf8"));
  window.outlookModePane.dockChanged();
  await wait(20);
  assert.ok($(".ol-composer"), "composer in the dock");
  assert.deepEqual(tabs(), ["New message*"]);
  assert.equal($(".ol-original"), null, "nothing quoted in a new message");
  assert.ok($(".ol-editor .moz-signature").innerHTML.includes("Best,"), "signature");
  // The signature's table stays; its picture shows through a blob: URL with the data kept.
  const sigImg = $(".ol-editor .moz-signature table img");
  assert.ok(sigImg, "signature table with picture");
  assert.match(sigImg.getAttribute("src"), /^blob:test-/);
  assert.equal(sigImg.getAttribute("data-ol-src"), PNG);
  const visibleRows = [...document.querySelectorAll(".ol-row:not(.ol-hidden) .ol-label")].map((l) => l.textContent);
  assert.deepEqual(visibleRows, ["To", "Subject"]);
  assert.equal(document.activeElement, $("textarea.ol-address"), "To is focused");

  // 2. The tab follows the subject; typing makes a draft.
  type($("textarea.ol-address"), "Ann <ann@x.com>");
  type(subjectInput(), "Plan");
  assert.deepEqual(tabs(), ["Plan*"]);
  await wait(900);
  assert.equal(drafts.get("new:1").subject, "Plan", "draft autosaved");

  // 3. Ctrl+N in the dock adds a tab; switching back restores the first one.
  key($(".ol-editor"), "n", { ctrlKey: true });
  await wait(20);
  assert.deepEqual(tabs(), ["Plan", "New message*"]);
  document.querySelectorAll(".ol-dock-tab")[0].click();
  await wait(20);
  assert.deepEqual(tabs(), ["Plan*", "New message"]);
  assert.equal($("textarea.ol-address").value, "Ann <ann@x.com>");

  // 4. Closing an untouched tab asks nothing; a draft asks first.
  confirms.length = 0;
  document.querySelectorAll(".ol-dock-tab")[1].querySelector(".ol-dock-close").click();
  assert.equal(confirms.length, 0);
  assert.deepEqual(tabs(), ["Plan*"]);

  // 5. Minimize / maximize / restore.
  button("Minimize").click();
  assert.ok(dockState.minimized);
  assert.ok(button("Restore"), "restore shown while minimized");
  button("Restore").click();
  button("Maximize").click();
  assert.ok(dockState.maximized);
  button("Restore size").click();
  assert.ok(!dockState.maximized && !dockState.minimized);

  // 6. Empty subject asks first; send with Undo Send, then the dock tucks away.
  type(subjectInput(), "");
  confirmAnswer = false;
  action("Send").click();
  await wait(20);
  assert.match(confirms.at(-1), /without a subject/);
  assert.equal(pending.size, 0, "not sent without the user's OK");
  confirmAnswer = true;
  type(subjectInput(), "Plan v2");
  action("Send").click();
  await wait(20);
  assert.match($(".ol-countdown span").textContent, /Sending in/);
  assert.deepEqual(tabs(), ["⏳ Plan v2*"]);
  [...document.querySelectorAll(".ol-countdown button")].find((b) => b.textContent === "Send now").click();
  await wait(20);
  assert.equal(sent.at(-1).subject, "Plan v2");
  assert.equal(sent.at(-1).type, "new");
  assert.equal($(".ol-composer"), null);
  assert.match($(".ol-dock-notice").textContent, /Message sent/);
  assert.ok(dockState.minimized, "minimized with the notice");
  await wait(4100);
  assert.ok(dockState.hidden, "hidden after the notice");

  // 7. A message sent from a tab that is not shown is reported in the dock.
  dockState.hidden = false;
  focusUri = createNew();
  window.outlookModePane.dockChanged();
  await wait(20);
  type($("textarea.ol-address"), "Bob <bob@x.com>");
  type(subjectInput(), "Background");
  action("Send").click();
  await wait(20);
  const backgroundUri = order[0];
  focusUri = createNew();
  window.outlookModePane.dockChanged();
  await wait(20);
  assert.deepEqual(tabs(), ["⏳ Background", "New message*"]);
  await bridge.sendNow(backgroundUri);
  assert.deepEqual(tabs(), ["New message*"]);
  assert.equal($("#ol-toast").textContent, "Sent.");
  assert.equal(window.outlookModePane.onSendResult("uri:5", { ok: true }), false, "inline replies are not the dock's");

  // 8. Pop out leaves the dock.
  type($("textarea.ol-address"), "Cy <cy@x.com>");
  action("Pop out").click();
  assert.equal(window.__popOut.type, "new");
  assert.ok(!window.__popOut.html.includes("moz-signature"), "signature left to the compose window");
  assert.ok(!window.__popOut.html.includes("blob:"), "no page-only blob: URLs in a popped-out message");
  assert.ok(dockState.hidden, "no tabs left: dock hidden");

  // 9. destroy() cleans up.
  window.outlookModePane.destroy();
  assert.equal(document.querySelector('link[href^="chrome://outlookmode"]'), null);

  console.log("all compose-dock scenarios passed");
  process.exit(0);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
