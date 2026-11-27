"use strict";
// Run: NODE_PATH=<dir with jsdom>/node_modules node tests/reading-pane.test.js
// (jsdom: `npm install jsdom@24` in any temporary folder)
// Runs reading-pane/pane.js in a jsdom copy of the Thunderbird Conversations
// page with a fake bridge (experiments/outlookPane), and checks the main flows.
const fs = require("fs");
const assert = require("assert");
const { JSDOM } = require("jsdom");

const PANE = require("path").join(__dirname, "../outlook-mode/reading-pane/pane.js");
const html = `<!doctype html><html><head></head><body><div id="conversationWrapper"><conversation-header></conversation-header><div id="messageList"></div><conversation-footer></conversation-footer></div></body></html>`;
const dom = new JSDOM(html, { runScripts: "outside-only", pretendToBeVisual: true });
const { window } = dom;
const { document } = window;
window.document.execCommand = (cmd, ui, value) => {
  window.__exec = [...(window.__exec || []), [cmd, value]];
  return true;
};
window.Element.prototype.scrollIntoView = function () {};
// jsdom has no ResizeObserver (Thunderbird has); the pane only needs it to exist.
window.ResizeObserver = class { observe() {} unobserve() {} disconnect() {} };
window.prompt = () => "https://example.com";
window.confirm = () => true;

const composeWindowOpened = [];

// Conversations-like custom elements.
window.eval(`
  class ActionButton extends HTMLElement {
    constructor() {
      super();
      this.attachShadow({ mode: "open" });
      const b = document.createElement("button");
      b.addEventListener("click", (e) => { e.stopPropagation(); window.__composeWindow.push([this.getAttribute("type"), this.getAttribute("msgid")]); });
      this.shadowRoot.appendChild(b);
      this.addEventListener("button-clicked", () => (window.__menuClosed = (window.__menuClosed || 0) + 1));
    }
  }
  customElements.define("action-button", ActionButton);
  class MessageFooter extends HTMLElement {
    connectedCallback() {
      if (this.shadowRoot) return;
      this.attachShadow({ mode: "open" });
      for (const t of ["reply", "replyAll", "forward"]) {
        const a = document.createElement("action-button");
        a.setAttribute("type", t);
        a.setAttribute("msgid", this.getAttribute("msgid"));
        this.shadowRoot.appendChild(a);
      }
    }
  }
  customElements.define("message-footer", MessageFooter);
  customElements.define("message-header-options", class extends HTMLElement {});
`);
window.__composeWindow = composeWindowOpened;

// expanded: ids Conversations shows expanded (with a footer); default all.
function showConversation(ids, expanded = ids) {
  const ul = document.createElement("ul");
  ul.id = "messageList";
  for (const id of ids) {
    const li = document.createElement("li");
    li.className = "message";
    li.tabIndex = 0;
    // Conversations' header with its star button (React flags on click).
    const messageHeader = document.createElement("div");
    messageHeader.className = expanded.includes(id) ? "messageHeader expanded" : "messageHeader";
    const star = document.createElement("button");
    star.className = "star";
    star.append(document.createElement("svg-icon"));
    star.addEventListener("click", () => (window.__flagged = [...(window.__flagged || []), id]));
    messageHeader.append(star);
    li.append(messageHeader);
    const header = document.createElement("message-header-options");
    header.setAttribute("msgid", String(id));
    const footer = document.createElement("message-footer");
    footer.setAttribute("msgid", String(id));
    li.append(header);
    if (expanded.includes(id)) {
      li.append(footer);
    }
    ul.appendChild(li);
  }
  document.getElementById("messageList").replaceChildren(ul);
}

// Fake bridge.
const drafts = new Map();
const pending = new Map();
const sent = [];
const prepared = [];
let lastReport = null;
const bridge = {
  config: () => ({ newestFirst: true, inlineReply: true, loadId: 1 }),
  sanitize: (html) => html, // the experiment runs Thunderbird's sanitizer
  uriForConvId: (id) => `uri:${id}`,
  convIdForUri: (uri) => Number(uri.split(":")[1]),
  prepare: async (uri, type) => {
    prepared.push([uri, type]);
    return {
      uri,
      type,
      identityKey: "id1",
      identities: [
        { key: "id1", label: "Me <me@x.com>", signature: "Best,<br>Me" },
        { key: "id2", label: "Alt <alt@x.com>", signature: "" },
      ],
      to: "Ann <ann@x.com>",
      cc: type === "replyAll" ? "Bob <bob@x.com>" : "",
      bcc: "",
      subject: type === "forward" ? "FW: Hello" : "RE: Hello",
      attachments: type === "forward" ? [{ url: "x", name: "a.pdf", size: 2048 }] : [],
      original: { from: "Ann <ann@x.com>", sent: "Monday", to: "Me", cc: "", subject: "Hello" },
      font: "",
    };
  },
  searchAddresses: (q) => (q.startsWith("ng") ? ["Nguyen A <a@x.com>", "Nguyen B <b@x.com>"] : []),
  getDraft: (uri) => (drafts.has(uri) ? structuredClone(drafts.get(uri)) : null),
  saveDraft: (uri, d) => drafts.set(uri, structuredClone(d)),
  discardDraft: (uri) => drafts.delete(uri),
  send: async (uri, d) => {
    drafts.set(uri, { ...structuredClone(d), dirty: true });
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
    lastReport = window.outlookModePane.onSendResult(uri, { ok: true });
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
  searchContacts: (q) =>
    "tien".startsWith(q.toLowerCase())
      ? [
          { name: "Hoàng Thị Tiên (APP)", email: "tien@x.com", label: "Hoàng Thị Tiên", address: '"Hoàng Thị Tiên (APP)" <tien@x.com>' },
          { name: "Tiến Phạm", email: "le@x.com", label: "Tiến Phạm", address: "Tiến Phạm <le@x.com>" },
        ]
      : [],
  messageDetails: (id) => ({
    to: [{ name: "Ann", email: "ann@x.com" }, { name: "", email: "bob@x.com" }],
    cc: Number(id) === 51 ? [{ name: "Carol", email: "carol@x.com" }] : [],
    bcc: [],
    time: Date.UTC(2026, 10, 2, 5, 0),
  }),
  isPinned: (id) => pinnedIds.has(Number(id)),
  togglePin: (id) => {
    pinToggles.push(Number(id));
    return true;
  },
};
const pinnedIds = new Set();
const pinToggles = [];
window.outlookModeBridge = bridge;

const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const $ = (sel) => document.querySelector(sel);
const clickInner = (actionButton) =>
  actionButton.shadowRoot.querySelector("button").dispatchEvent(new window.MouseEvent("click", { bubbles: true, composed: true, button: 0 }));
const footerButton = (id, type) =>
  document.querySelector(`message-footer[msgid="${id}"]`).shadowRoot.querySelector(`action-button[type="${type}"]`);
const type = (input, text) => {
  input.value = text;
  input.setSelectionRange(text.length, text.length);
  input.dispatchEvent(new window.Event("input", { bubbles: true }));
};
const key = (target, k, opts = {}) =>
  target.dispatchEvent(new window.KeyboardEvent("keydown", { key: k, bubbles: true, composed: true, cancelable: true, ...opts }));

// Conversations scrolls with window.scrollTo (jsdom has none that works).
const scrolls = [];
const pageScrollTo = (x, y) => scrolls.push([x, y]);
window.scrollTo = pageScrollTo;

(async () => {
  showConversation([11, 12, 13]);
  window.eval(fs.readFileSync(PANE, "utf8"));
  assert.ok(window.outlookModePane, "pane installed");
  assert.ok(document.documentElement.hasAttribute("data-ol-newest-first"));
  assert.ok(document.querySelector('link[href^="chrome://outlookmode/content/pane.css"]'));

  // 1. Conversations' Reply all button opens the inline composer, not a window.
  clickInner(footerButton(13, "replyAll"));
  await wait(20);
  assert.equal(composeWindowOpened.length, 0, "compose window not opened");
  assert.equal(window.__menuClosed, 1, "menu told to close");
  let c = $(".ol-composer");
  assert.ok(c, "composer shown");
  assert.equal(c.nextElementSibling.id, "messageList", "composer above the messages");
  assert.equal($(".ol-kind").textContent, "Reply all");
  const inputs = () => [...c.querySelectorAll("textarea.ol-address")];
  assert.equal(inputs()[0].value, "Ann <ann@x.com>");
  assert.equal(inputs()[1].value, "Bob <bob@x.com>");
  assert.ok(!inputs()[1].closest(".ol-row").classList.contains("ol-hidden"), "Cc row shown for reply all");
  assert.ok($(".ol-editor .moz-signature").innerHTML.includes("Best,"), "signature");

  // 2. From change swaps the signature.
  const from = c.querySelector("select");
  from.value = "id2";
  from.dispatchEvent(new window.Event("change"));
  assert.equal($(".ol-editor .moz-signature"), null, "no signature for id2");
  from.value = "id1";
  from.dispatchEvent(new window.Event("change"));

  // 3. Typing marks a draft; keys do not leak to Thunderbird / Conversations.
  let leaked = 0;
  document.addEventListener("keydown", () => leaked++);
  const editor = $(".ol-editor");
  editor.insertBefore(document.createTextNode("Hello Ann"), editor.firstChild);
  editor.dispatchEvent(new window.Event("input", { bubbles: true }));
  const ctrlD = key(editor, "d", { ctrlKey: true });
  assert.equal(ctrlD, false, "Ctrl+D default prevented");
  const ran = (command, value) =>
    window.__exec.some(([c, v]) => c === command && (value === undefined || v === value));
  key(editor, "b", { ctrlKey: true });
  assert.ok(ran("bold"), "Ctrl+B");
  key(editor, "e", { ctrlKey: true });
  assert.ok(ran("justifyCenter"), "Ctrl+E centers, as in Outlook");
  key(editor, " ", { ctrlKey: true });
  assert.ok(ran("removeFormat"), "Ctrl+Space clears formatting");
  assert.equal(leaked, 0, "keys stay in the composer");

  // 3a. Tab stays in the text: a tab mid-line, indent at the start of a
  // paragraph, list levels in a list; "- " / "1. " start a list.
  const caret = (node, offset) => window.getSelection().collapse(node, offset);
  const lastExec = () => window.__exec.filter(([c]) => c !== "styleWithCSS").at(-1);
  const cssMode = () => window.__exec.filter(([c]) => c === "styleWithCSS").at(-2)?.[1];
  window.__exec = [];
  caret(editor.firstChild, 5); // "Hello| Ann"
  assert.equal(key(editor, "Tab"), false, "Tab default prevented (focus stays in the text)");
  assert.deepEqual(lastExec(), ["insertText", "\u00a0\u00a0\u00a0\u00a0"], "mid-line: a tab");
  const para = document.createElement("div");
  para.textContent = "Second paragraph";
  editor.append(para);
  caret(para.firstChild, 0);
  key(editor, "Tab");
  assert.deepEqual(lastExec(), ["indent", null], "start of a paragraph: indented");
  assert.equal(cssMode(), true, "as a margin, not a quote");
  key(editor, "Tab", { shiftKey: true });
  assert.deepEqual(lastExec(), ["outdent", null], "Shift+Tab: back");
  const list = document.createElement("ul");
  list.innerHTML = "<li>First</li><li>Second</li>";
  editor.append(list);
  caret(list.lastChild.firstChild, 3); // mid-item
  key(editor, "Tab");
  assert.deepEqual(lastExec(), ["indent", null], "in a list: one level in");
  assert.equal(cssMode(), false, "a list inside the list");
  key(editor, "Tab", { shiftKey: true });
  assert.deepEqual(lastExec(), ["outdent", null], "Shift+Tab: one level out");
  const typed = (text) => {
    const line = document.createElement("div");
    line.textContent = text;
    editor.append(line);
    caret(line.firstChild, text.length);
    window.__exec = [];
    editor.dispatchEvent(new window.InputEvent("input", { bubbles: true, inputType: "insertText", data: " " }));
    return window.__exec.map(([c]) => c);
  };
  assert.deepEqual(typed("- "), ["delete", "insertUnorderedList"], '"- " starts a bulleted list');
  assert.deepEqual(typed("* "), ["delete", "insertUnorderedList"]);
  assert.deepEqual(typed("1. "), ["delete", "insertOrderedList"], '"1. " starts a numbered list');
  assert.deepEqual(typed("Note - "), [], "not at the start: plain text");
  assert.equal(leaked, 0, "Tab does not reach Thunderbird either");

  // 3b. Outlook's Format Text tools.
  const tool = (name) => $(`.ol-composer .ol-tool-${name}`);
  for (const name of ["undo", "font", "size", "strikeThrough", "superscript", "colorMenu", "highlightMenu",
    "justifyRight", "style", "tableMenu", "insertHorizontalRule", "emojiMenu", "picture", "attach"]) {
    assert.ok(tool(name), `toolbar has ${name}`);
  }
  window.__exec = [];
  tool("colorMenu").click();
  assert.ok($(".ol-popup .ol-swatches"), "color menu open");
  $('.ol-popup .ol-swatch[title="#C00000"]').click();
  assert.deepEqual(window.__exec.slice(0, 3), [["styleWithCSS", true], ["foreColor", "#C00000"], ["styleWithCSS", false]]);
  assert.equal($(".ol-popup"), null, "menu closed");
  assert.equal(tool("color").style.getPropertyValue("--ol-swatch"), "#C00000", "the split button remembers the color");
  tool("color").click();
  assert.equal(window.__exec.filter(([c]) => c === "foreColor").length, 2);
  tool("tableMenu").click();
  $('.ol-popup .ol-table-cell[aria-label="3 × 2"]').click();
  const [, tableHtml] = window.__exec.filter(([c]) => c === "insertHTML").at(-1);
  assert.equal((tableHtml.match(/<tr>/g) || []).length, 2, "2 rows");
  assert.equal((tableHtml.match(/<td /g) || []).length, 6, "3 columns");
  // Sizes in points: the size-7 marks execCommand makes become font-size spans.
  const size = tool("size");
  size.value = "14";
  size.dispatchEvent(new window.Event("change"));
  assert.ok(ran("fontSize", "7"));
  const mark = document.createElement("font");
  mark.setAttribute("size", "7");
  mark.textContent = "big";
  editor.append(mark);
  editor.dispatchEvent(new window.Event("input", { bubbles: true }));
  assert.equal(editor.querySelector('font[size="7"]'), null);
  assert.ok(editor.innerHTML.includes('<span style="font-size: 14pt;">big</span>'));
  editor.lastChild.remove();
  assert.match($(".ol-status").textContent, /Draft/);
  await wait(900);
  assert.ok(drafts.get("uri:13").html.includes("Hello Ann"), "draft autosaved");

  // 4. Address suggestions.
  const cc = inputs()[1];
  type(cc, "Bob <bob@x.com>, ng");
  assert.ok($(".ol-suggest"), "suggestions shown");
  key(cc, "ArrowDown");
  key(cc, "Enter");
  assert.equal(cc.value, "Bob <bob@x.com>, Nguyen B <b@x.com>, ");
  assert.equal($(".ol-suggest"), null);

  // 4b. Address fields wrap (textarea); pasted lines become one list, Enter adds no line break.
  const ccBefore = cc.value;
  type(cc, "Ann <ann@x.com>\nBob <bob@x.com>\r\n");
  assert.equal(cc.value, "Ann <ann@x.com>, Bob <bob@x.com>, ");
  assert.equal(key(cc, "Enter"), false, "Enter default prevented");
  type(cc, ccBefore);

  // 4c. @mention in the body: pick a contact, a mailto link is inserted, To gets them.
  {
    const editor = $(".ol-editor");
    const line = document.createElement("div");
    const text = document.createTextNode("Thanks @ti");
    line.append(text);
    editor.insertBefore(line, editor.firstChild);
    const selection = window.getSelection();
    const range = document.createRange();
    range.setStart(text, text.length);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
    editor.dispatchEvent(new window.Event("input", { bubbles: true }));
    assert.equal(document.querySelectorAll(".ol-mention .ol-mention-item").length, 2, "mention list shown");
    key(editor, "ArrowDown");
    assert.ok(document.querySelectorAll(".ol-mention-item")[1].classList.contains("ol-active"));
    key(editor, "ArrowUp");
    const toBefore = inputs()[0].value;
    assert.equal(key(editor, "Enter"), false, "Enter used by the list");
    const [command, html] = window.__exec.at(-1);
    assert.equal(command, "insertHTML");
    assert.match(html, /^<a href="mailto:tien@x\.com">@Hoàng Thị Tiên<\/a>&nbsp;$/);
    assert.equal($(".ol-mention"), null, "list closed");
    assert.equal(inputs()[0].value, `${toBefore}, "Hoàng Thị Tiên (APP)" <tien@x.com>, `, "added to To");
    // Mentioning someone already in To/Cc does not add them again.
    text.data = "Thanks @ti";
    range.setStart(text, text.length);
    selection.removeAllRanges();
    selection.addRange(range);
    editor.dispatchEvent(new window.Event("input", { bubbles: true }));
    key(editor, "Enter");
    assert.equal((inputs()[0].value.match(/tien@x\.com/g) || []).length, 1, "not added twice");
    // Escape closes the list; text without "@" shows none.
    editor.dispatchEvent(new window.Event("input", { bubbles: true }));
    key(editor, "Escape");
    assert.equal($(".ol-mention"), null);
    line.remove();
    inputs()[0].value = toBefore;
  }

  // 5. Another conversation: composer closes, draft kept; coming back restores it.
  showConversation([21, 22]);
  await wait(250);
  assert.equal($(".ol-composer"), null, "closed for another conversation");
  assert.ok(drafts.get("uri:13").cc.includes("Nguyen B"), "draft has the latest fields");
  showConversation([11, 12, 13]);
  await wait(250);
  c = $(".ol-composer");
  assert.ok(c, "draft restored");
  assert.ok($(".ol-editor").innerHTML.includes("Hello Ann"));
  assert.match($(".ol-status").textContent, /Draft/);

  // 6. Send -> Undo Send countdown -> Undo -> Send -> Send now.
  [...c.querySelectorAll(".ol-actions button")].find((b) => b.textContent === "Send").click();
  await wait(20);
  assert.ok(!$(".ol-countdown").classList.contains("ol-hidden"), "countdown shown");
  assert.match($(".ol-countdown span").textContent, /Sending in \d+s/);
  [...c.querySelectorAll(".ol-countdown button")].find((b) => b.textContent === "Undo").click();
  assert.ok($(".ol-countdown").classList.contains("ol-hidden"), "undo hides countdown");
  assert.equal(pending.size, 0);
  key($(".ol-editor"), "Enter", { ctrlKey: true }); // Ctrl+Enter sends
  await wait(20);
  [...c.querySelectorAll(".ol-countdown button")].find((b) => b.textContent === "Send now").click();
  await wait(20);
  assert.equal(sent.length, 1, "sent");
  assert.equal(sent[0].type, "replyAll");
  assert.ok(sent[0].html.includes("Hello Ann"));
  assert.equal(lastReport, true, "result shown in the pane");
  assert.equal($(".ol-composer"), null, "composer closed after sending");

  // 7. Ctrl+Shift+R on a message, then Thunderbird's command (newest).
  key(document.querySelectorAll("li.message")[0], "R", { ctrlKey: true, shiftKey: true });
  await wait(20);
  assert.deepEqual(prepared.at(-1), ["uri:11", "replyAll"]);
  assert.equal(window.outlookModePane.openFor("uri:12", "reply", { newest: true }), true);
  await wait(20);
  assert.deepEqual(prepared.at(-1), ["uri:13", "reply"]);
  assert.equal($(".ol-kind").textContent, "Reply");
  assert.equal(document.querySelectorAll(".ol-composer").length, 1, "one composer at a time");

  // 8. Forward shows the original's attachments; pop out drops the signature.
  window.outlookModePane.openFor("uri:12", "forward");
  await wait(20);
  assert.match($(".ol-attachments").textContent, /a\.pdf \(2 KB\)/);
  [...$(".ol-composer").querySelectorAll(".ol-actions button")].find((b) => b.textContent === "Pop out").click();
  assert.equal(window.__popOut.type, "forward");
  assert.ok(!window.__popOut.html.includes("moz-signature"), "signature left to the compose window");
  assert.equal($(".ol-composer"), null);

  // 9. Turning inline reply off lets Conversations open its compose window.
  bridge.config = () => ({ newestFirst: false, inlineReply: false, loadId: 1 });
  window.outlookModePane.configChanged();
  assert.ok(!document.documentElement.hasAttribute("data-ol-newest-first"));
  clickInner(footerButton(12, "reply"));
  assert.deepEqual(composeWindowOpened.at(-1), ["reply", "12"]);
  assert.equal(window.outlookModePane.openFor("uri:12", "reply"), false);

  // 10. Newest message only: earlier collapsed messages behind a link, reset per conversation.
  bridge.config = () => ({ newestFirst: true, latestOnly: true, inlineReply: true, loadId: 1 });
  window.outlookModePane.configChanged();
  showConversation([31, 32, 33, 34], [32, 34]);
  await wait(250);
  const earlier = () => document.getElementById("ol-earlier");
  assert.ok(document.documentElement.hasAttribute("data-ol-latest-only"));
  assert.equal(earlier().textContent, "Show 2 earlier messages", "expanded 32 is not counted");
  assert.equal(earlier().previousElementSibling, document.querySelector("div#messageList"), "outside React's root");
  earlier().click();
  assert.ok(document.documentElement.hasAttribute("data-ol-show-earlier"));
  assert.equal(earlier().textContent, "Hide earlier messages");
  showConversation([31, 32, 33, 34, 35], [35]); // a new message in the same conversation
  await wait(250);
  assert.ok(document.documentElement.hasAttribute("data-ol-show-earlier"), "still shown");
  showConversation([41, 42], [42]);
  await wait(250);
  assert.ok(!document.documentElement.hasAttribute("data-ol-show-earlier"), "reset for another conversation");
  assert.equal(earlier().textContent, "Show 1 earlier message");
  showConversation([51]);
  await wait(250);
  assert.equal(earlier(), null, "no link for a single message");

  // 10b. The star is a pin button: it pins (through the background), never flags.
  bridge.config = () => ({ newestFirst: true, latestOnly: true, inlineReply: true, pinButton: true, loadId: 1 });
  window.outlookModePane.configChanged();
  const starOf = (id) => document.querySelector(`message-header-options[msgid="${id}"]`).closest("li").querySelector("button.star");
  assert.ok(document.documentElement.hasAttribute("data-ol-pin-button"));
  assert.equal(starOf(51).title, "Pin");
  starOf(51).querySelector("svg-icon").dispatchEvent(new window.MouseEvent("click", { bubbles: true, composed: true, button: 0 }));
  assert.deepEqual(pinToggles, [51]);
  assert.equal(window.__flagged, undefined, "Conversations' star handler did not run");
  pinnedIds.add(51); // the background tagged it
  await wait(450);
  assert.ok(starOf(51).hasAttribute("data-ol-pinned"));
  assert.equal(starOf(51).title, "Unpin");

  // 10c. Outlook header: To and Cc lines under an expanded message's header.
  bridge.config = () => ({ inlineReply: true, outlookHeader: true, lightPane: true, loadId: 1 });
  window.outlookModePane.configChanged();
  assert.ok(document.documentElement.hasAttribute("data-ol-light"));
  const rcpt = $(".ol-recipients");
  assert.ok(rcpt, "recipient lines");
  assert.equal(rcpt.previousElementSibling.className, "messageHeader expanded");
  assert.deepEqual([...rcpt.querySelectorAll(".ol-rcpt-label")].map((l) => l.textContent), ["To", "Cc"]);
  assert.equal(rcpt.querySelector(".ol-rcpt-list").textContent, "Ann <ann@x.com>; bob@x.com");
  assert.equal(rcpt.querySelector(".ol-rcpt-toggle").textContent, "▾ 3");
  rcpt.querySelector(".ol-rcpt-toggle").click();
  assert.ok(rcpt.classList.contains("ol-open"), "click shows everyone");
  window.outlookModePane.configChanged(); // the same message: the lines stay as they are
  assert.equal($(".ol-recipients"), rcpt);
  bridge.config = () => ({ inlineReply: true, outlookHeader: false, loadId: 1 });
  window.outlookModePane.configChanged();
  assert.equal($(".ol-recipients"), null, "off: Conversations' own header");

  // 10d. Every message open: closed ones opened through Conversations' store
  // (plain msgExpand, nothing marked read), one at a time, newest first;
  // opened again when Conversations shows the conversation anew, but one
  // the user closes stays closed.
  const dispatched = [];
  window.eval(`customElements.define("conversation-header", class extends HTMLElement { static dispatch; });`);
  function customElementsHeader() {
    return window.customElements.get("conversation-header");
  }
  const headerOf = (id) => document.querySelector(`message-header-options[msgid="${id}"]`)?.closest("li").querySelector(".messageHeader");
  // Like Conversations: the message shows as open after msgExpand.
  const conversationsStore = (action) => {
    dispatched.push([action.type, action.payload.id, action.payload.expand ?? action.payload.hasRemoteContent]);
    if (action.type === "messages/msgExpand") {
      headerOf(action.payload.id)?.classList.add("expanded");
    }
  };
  customElementsHeader().dispatch = conversationsStore;
  bridge.config = () => ({ newestFirst: true, expandAll: true, latestOnly: true, inlineReply: true, loadId: 1 });
  window.outlookModePane.configChanged();
  showConversation([61, 62, 63], [63]);
  await wait(250);
  assert.deepEqual(dispatched, [["messages/msgExpand", 62, true]], "one at a time, newest first");
  await wait(500);
  assert.deepEqual(dispatched, [["messages/msgExpand", 62, true], ["messages/msgExpand", 61, true]], "then the next");
  assert.ok(!document.documentElement.hasAttribute("data-ol-latest-only"), "nothing hidden while all are open");
  assert.equal(earlier(), null, "no earlier-messages link");
  showConversation([61, 62, 63]); // Conversations shows them open
  await wait(600);
  assert.equal(dispatched.length, 2);
  showConversation([61, 62, 63], [63]); // shown anew (selected again): Conversations closed them
  await wait(1000);
  assert.deepEqual(dispatched.slice(2), [["messages/msgExpand", 62, true], ["messages/msgExpand", 61, true]]);
  showConversation([61, 62, 63]);
  await wait(250);
  headerOf(61).dispatchEvent(new window.MouseEvent("click", { bubbles: true })); // the user closes 61
  showConversation([61, 62, 63], [62, 63]);
  await wait(600);
  assert.equal(dispatched.length, 4, "a message the user closes stays closed");
  showConversation([61, 62, 63, 64], [62, 63]); // a new mail in the same conversation
  await wait(600);
  assert.deepEqual(dispatched.slice(4), [["messages/msgExpand", 64, true]], "only the new mail is opened");
  showConversation([71, 72], [72]); // another conversation
  await wait(600);
  assert.deepEqual(dispatched.at(-1), ["messages/msgExpand", 71, true]);
  assert.equal(dispatched.length, 6);
  customElementsHeader().dispatch = undefined; // no store: the header is clicked as the user would
  const headerClicks = [];
  showConversation([81, 82], [82]);
  headerOf(81).addEventListener("click", () => headerClicks.push(81));
  await wait(1800);
  assert.deepEqual(headerClicks, [81, 81, 81], "tried a few times, then left alone");
  assert.equal(dispatched.length, 6);
  customElementsHeader().dispatch = conversationsStore;
  bridge.config = () => ({ newestFirst: true, expandAll: false, latestOnly: true, inlineReply: true, loadId: 1 });
  window.outlookModePane.configChanged();
  showConversation([91, 92], [92]);
  await wait(600);
  assert.equal(dispatched.length, 6, "off: Conversations decides what is open");
  assert.ok(document.documentElement.hasAttribute("data-ol-latest-only"), "off: newest-only applies again");

  // 10e. The earlier mail quoted below each message (but the oldest) is folded.
  const bodies = {
    101: '<p>First</p><div id="divRplyFwdMsg"><b>From:</b> Zed</div><div>Not in this conversation</div>',
    102:
      '<div class="WordSection1"><p>Outlook reply text</p><div><div style="border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0in 0in 0in">' +
      "<p><b>From:</b> Ann<br><b>Sent:</b> Monday</p></div></div><p>Outlook original</p></div>",
    103: '<div>Web reply</div><div id="appendonsend"></div><hr><div id="divRplyFwdMsg"><b>Từ:</b> Bob</div><div>Web original</div>',
    104: '<div id="divRplyFwdMsg"><b>From:</b> Carol</div><div>Forwarded only</div>',
    105: "<pre>Plain reply\n________________________________\nFrom: Bob\nSent: Monday\n\nPlain original</pre>",
    106: '<p>TB reply</p><div class="moz-cite-prefix">On Monday Ann wrote:</div><button class="showhidequote">-</button><blockquote type="cite">TB original</blockquote>',
    107: '<p>Inline intro</p><div class="moz-cite-prefix">On Monday Ann wrote:</div><blockquote type="cite">Short quote</blockquote><p>Inline answer</p>',
    // New Outlook (Japanese), no #divRplyFwdMsg: a line, then the header.
    108:
      '<div class="elementToProof">New Outlook reply</div><div><br></div>' +
      '<hr style="display: inline-block; width: 98%;"><div style="font-size: 12pt;"><b>差出人:</b> Thuyen<br><b>送信:</b> 2026 年 11 月 20 日<br><b>件名:</b> RE: Hello</div>' +
      "<div>Japanese original</div>",
    // A line in the text that is not a quote stays.
    109: '<p>Above the line</p><hr><p>Below the line</p>',
  };
  const frames = {};
  bridge.config = () => ({ newestFirst: true, expandAll: true, foldQuotes: true, latestOnly: true, inlineReply: true, loadId: 1 });
  window.outlookModePane.configChanged();
  showConversation(Object.keys(bodies).map(Number));
  for (const [id, body] of Object.entries(bodies)) {
    const frame = document.createElement("iframe");
    frame.style.height = "100px";
    document.querySelector(`message-header-options[msgid="${id}"]`).closest("li").append(frame);
    const frameDoc = frame.contentDocument;
    Object.defineProperty(frameDoc, "URL", { value: `mailbox:///Inbox?number=${id}` }); // the message, loaded
    frameDoc.body.innerHTML = body;
    const style = frameDoc.createElement("style");
    style.textContent = "span.moz-txt-formfeed { height: auto; }"; // Conversations is done with it
    frameDoc.head.append(style);
    frames[id] = frameDoc;
  }
  await wait(450);
  const hiddenText = (id, text) => {
    const walker = frames[id].createTreeWalker(frames[id].body, 4);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      if (node.data.includes(text)) {
        return node.parentElement.closest("[data-ol-quoted]")?.style.display === "none";
      }
    }
    throw new Error(`no ${text} in ${id}`);
  };
  const toggleOf = (id) => frames[id].querySelector(".ol-quote-toggle");
  assert.equal(toggleOf(101), null, "the oldest message keeps its quote");
  assert.ok(!hiddenText(102, "Outlook reply text") && hiddenText(102, "Outlook original") && hiddenText(102, "Ann"), "Outlook desktop");
  assert.equal(toggleOf(102).textContent, "- show quoted text -");
  assert.ok(!hiddenText(103, "Web reply") && hiddenText(103, "Web original"), "Outlook on the web");
  assert.equal(frames[103].querySelector("hr").closest("[data-ol-quoted]").style.display, "none", "with its line");
  assert.equal(toggleOf(104), null, "a forward without text stays as it is");
  assert.ok(!hiddenText(105, "Plain reply") && hiddenText(105, "Plain original"), "plain-text Outlook");
  assert.equal(toggleOf(106), null, "Conversations' own fold stays");
  assert.ok(!hiddenText(107, "Inline intro") && hiddenText(107, "Short quote") && hiddenText(107, "On Monday") && !hiddenText(107, "Inline answer"),
    "only the quote box, answers after it stay");
  assert.ok(!hiddenText(108, "New Outlook reply") && hiddenText(108, "Thuyen") && hiddenText(108, "Japanese original"), "new Outlook, header after a line");
  assert.equal(frames[108].querySelector("hr").closest("[data-ol-quoted]").style.display, "none", "with its line");
  assert.equal(toggleOf(109), null, "a plain line is not a quote");
  toggleOf(102).click();
  assert.ok(!hiddenText(102, "Outlook original"), "show quoted text");
  assert.equal(toggleOf(102).textContent, "- hide quoted text -");
  bridge.config = () => ({ newestFirst: true, expandAll: true, foldQuotes: false, inlineReply: true, loadId: 1 });
  window.outlookModePane.configChanged();
  assert.ok(!hiddenText(103, "Web original") && toggleOf(103).hidden, "off: everything shown");

  // 10f. One message loading at a time (Conversations' own included), newest
  // first; a message that loaded empty (a Graph download cut off) is loaded
  // again at once, then after 3 s and 10 s; it says it is loading, then that
  // it could not be loaded, and a click on that loads it again.
  const frameState = {};
  const showLoading = (ids, expanded) => {
    showConversation(ids, expanded);
    for (const id of ids) {
      const body = document.createElement("div");
      body.className = "messageBody";
      const wrap = document.createElement("div");
      wrap.className = `iframewrap${id}`;
      const frame = document.createElement("iframe");
      Object.defineProperty(frame, "contentDocument", { configurable: true, get: () => frameState[id] });
      frameState[id] ||= { URL: "about:blank", readyState: "complete", body: null };
      wrap.append(frame);
      body.append(wrap);
      headerOf(id).after(body);
    }
  };
  const liOf = (id) => headerOf(id).closest("li");
  const fakeBody = (text) => ({ textContent: text, querySelector: () => null });
  const loaded = (id, text = "mail text") =>
    (frameState[id] = { URL: `mailbox:///Inbox?number=${id}`, readyState: "complete", body: fakeBody(text) });
  customElementsHeader().dispatch = conversationsStore;
  bridge.config = () => ({ newestFirst: true, expandAll: true, attachmentsTop: true, inlineReply: true, loadId: 1 });
  window.outlookModePane.configChanged();
  assert.ok(document.documentElement.hasAttribute("data-ol-attachments-top"), "attachments above the text (pane.css)");
  dispatched.length = 0;
  showLoading([201, 202, 203, 204, 205], [205]); // Conversations is loading the newest
  await wait(600);
  assert.deepEqual(dispatched, [], "nothing else while the newest loads");
  assert.equal(liOf(205).getAttribute("data-ol-load"), null, "no notice for a quick load");
  await wait(700);
  assert.equal(liOf(205).getAttribute("data-ol-load"), "loading", "a slow one says it is loading");
  loaded(205);
  await wait(500);
  assert.equal(liOf(205).getAttribute("data-ol-load"), null);
  assert.deepEqual(dispatched, [["messages/msgExpand", 204, true]], "the next one when it has loaded");
  await wait(500);
  assert.equal(dispatched.length, 1, "204 still loading");
  loaded(204, ""); // loaded empty: its download was cut off
  await wait(500);
  assert.deepEqual(dispatched.slice(1), [
    ["messages/setHasRemoteContent", 204, true],
    ["messages/setHasRemoteContent", 204, false],
  ], "loaded again at once, before the next one");
  await wait(500);
  assert.equal(dispatched.length, 3, "the empty document no longer counts as loaded");
  loaded(204);
  await wait(500);
  assert.deepEqual(dispatched.at(-1), ["messages/msgExpand", 203, true]);
  const kinds = () => dispatched.slice(4).map(([type, id]) => `${type.split("/")[1]} ${id}`);
  loaded(203, "");
  await wait(500); // reload 1 at once
  loaded(203, "");
  await wait(500); // reload 2 waits 3 s: the rest go on meanwhile
  assert.equal(liOf(203).getAttribute("data-ol-load"), "loading");
  assert.deepEqual(kinds(), [
    "setHasRemoteContent 203", "setHasRemoteContent 203", "msgExpand 202",
  ], "the next message while 203 waits");
  loaded(202);
  await wait(3000);
  assert.deepEqual(kinds().slice(3), ["msgExpand 201", "setHasRemoteContent 203", "setHasRemoteContent 203"].slice(0, kinds().length - 3));
  loaded(201);
  await wait(500);
  loaded(203, "");
  await wait(10500); // reload 3 after 10 s
  loaded(203, "");
  await wait(500);
  assert.equal(kinds().filter((k) => k === "setHasRemoteContent 203").length, 6, "three reloads");
  assert.equal(liOf(203).getAttribute("data-ol-load"), "failed", "then: could not be loaded");
  await wait(500);
  assert.equal(kinds().filter((k) => k === "setHasRemoteContent 203").length, 6, "no more by itself");
  liOf(203).querySelector(".iframewrap203").dispatchEvent(new window.MouseEvent("click", { bubbles: true }));
  await wait(300);
  assert.equal(kinds().filter((k) => k === "setHasRemoteContent 203").length, 8, "a click loads it again");
  assert.equal(liOf(203).getAttribute("data-ol-load"), "loading");

  // 10b. Date order: Conversations appends a message that comes in later
  // (here 74, a reply indexed after the conversation was shown) at the end
  // of its list, whatever its date. The pane orders by date: 74 goes between
  // 72 and 73, and 73 stays the newest.
  const dates = { 71: 1000, 72: 2000, 73: 4000, 74: 3000 };
  bridge.dateForConvId = (id) => dates[id] ?? null;
  bridge.config = () => ({ newestFirst: true, latestOnly: true, inlineReply: true, loadId: 1 });
  window.outlookModePane.configChanged();
  showConversation([71, 72, 73, 74], [73]);
  await wait(250);
  const orderOf = (id) => document.querySelector(`message-header-options[msgid="${id}"]`).closest("li").style.order;
  assert.deepEqual([71, 72, 74, 73].map(orderOf), ["0", "1", "2", "3"], "oldest 0 ... newest last");
  const newest = [...document.querySelectorAll("li.message[data-ol-newest]")].map((li) => li.querySelector("message-header-options").getAttribute("msgid"));
  assert.deepEqual(newest, ["73"], "the newest by date, not the last in the page");
  assert.match(earlier().textContent, /Show 3 earlier messages/);
  // A date not known yet: the page's order.
  dates[75] = undefined;
  showConversation([71, 72, 73, 75], [73]);
  await wait(250);
  assert.deepEqual([71, 72, 73, 75].map(orderOf), ["0", "1", "2", "3"]);
  delete bridge.dateForConvId;

  // 10c. A conversation opens at its newest message (the top), not where
  // Conversations would scroll (the selected or first unread message, down
  // the page); once the user scrolls, Conversations' later scrolls of that
  // opening (a frame done loading) leave the page alone.
  bridge.config = () => ({ newestFirst: true, inlineReply: true, loadId: 1 });
  window.outlookModePane.configChanged();
  scrolls.length = 0;
  showConversation([81, 82, 83], [83]);
  await wait(250);
  assert.deepEqual(scrolls, [[0, 0]], "another conversation: to the top");
  window.scrollTo(0, 900); // Conversations, to the selected message
  assert.deepEqual(scrolls.at(-1), [0, 0], "Conversations' scroll goes to the top");
  document.dispatchEvent(new window.WheelEvent("wheel", { bubbles: true }));
  scrolls.length = 0;
  window.scrollTo(0, 900); // its frame finished loading
  assert.deepEqual(scrolls, [], "the user scrolled: left alone");
  showConversation([84, 85], [85]);
  await wait(250);
  assert.deepEqual(scrolls, [[0, 0]], "the next conversation starts at the top again");
  showConversation([84, 85, 86], [85]); // a message comes in: no jump
  await wait(250);
  assert.equal(scrolls.length, 1);
  bridge.config = () => ({ newestFirst: false, inlineReply: true, loadId: 1 });
  window.outlookModePane.configChanged();
  scrolls.length = 0;
  window.scrollTo(0, 900);
  assert.deepEqual(scrolls, [[0, 900]], "oldest first: Conversations scrolls as it likes");

  // 10d. The toolbar's save button: every message of the conversation, in
  // date order, to the experiment (which checks and saves them).
  const convHeader = document.createElement("conversation-header");
  const actions = document.createElement("conv-actions-buttons");
  convHeader.attachShadow({ mode: "open" }).append(actions);
  const bar = actions.attachShadow({ mode: "open" });
  for (const name of ["expand", "archive", "trash"]) {
    const b = document.createElement("button");
    b.className = `button-flat actions-button ${name}`;
    bar.append(b);
  }
  document.body.prepend(convHeader);
  const saveCalls = [];
  bridge.saveConversation = async (ids) => {
    saveCalls.push(ids);
    return { saved: ids.length, file: "/home/me/Mail/Plan - 2026-11-24.eml" };
  };
  bridge.dateForConvId = (id) => ({ 91: 1000, 92: 3000, 93: 2000 })[id] ?? null;
  showConversation([91, 92, 93], [92]);
  await wait(250);
  const saveButton = bar.querySelector(".ol-save");
  assert.ok(saveButton, "save button added");
  assert.equal(saveButton.nextElementSibling.classList.contains("archive"), true, "before Archive");
  assert.equal(saveButton.querySelector("svg-icon").getAttribute("hash"), "save_alt");
  saveButton.click();
  await wait(20);
  assert.deepEqual(saveCalls, [[91, 93, 92]], "all messages, oldest first");
  assert.match(document.getElementById("ol-toast").textContent, /Saved the conversation \(3 messages\) to \/home\/me\/Mail\/Plan - 2026-11-24\.eml/);
  showConversation([91, 92, 93], [92]);
  await wait(250);
  assert.equal(bar.querySelectorAll(".ol-save").length, 1, "added once");
  bridge.saveConversation = async () => ({ saved: 0, file: "" }); // cancelled
  saveButton.click();
  await wait(20);
  assert.ok(!document.getElementById("ol-toast").classList.contains("ol-visible"), "cancelled: the Saving… notice goes");
  delete bridge.dateForConvId;

  // 11. destroy() cleans up.
  showConversation([41, 42], [42]);
  await wait(250);
  assert.equal(document.querySelector("li.message[data-ol-newest] message-header-options").getAttribute("msgid"), "42");
  window.outlookModePane.destroy();
  assert.equal(document.querySelector("li.message[data-ol-newest]"), null, "marks removed");
  assert.equal(document.querySelector("li.message").style.order, "", "order removed");
  assert.equal(window.scrollTo, pageScrollTo, "page's scrollTo back");
  assert.equal(bar.querySelector(".ol-save"), null, "save button removed");
  assert.equal(earlier(), null);
  assert.ok(!document.documentElement.hasAttribute("data-ol-latest-only"));
  assert.ok(!document.documentElement.hasAttribute("data-ol-attachments-top"));
  assert.equal(document.querySelector('link[href^="chrome://outlookmode"]'), null);

  console.log("all reading-pane scenarios passed");
  process.exit(0);
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
