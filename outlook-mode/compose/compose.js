// Outlook Mode - compose script (runs inside the message body editor):
// @mentions, Outlook-style reply/forward header and the Undo Send banner.
(() => {
  "use strict";
  // Guard per extension context: a re-injection after an add-on update gets a
  // fresh sandbox, so it is not blocked by a dead copy of an older version.
  if (window.__outlookModeLoaded) {
    return;
  }
  window.__outlookModeLoaded = true;

  const api = globalThis.messenger || globalThis.browser;
  const doc = document;
  const UI_ATTR = "data-outlook-mode-ui";
  // "@" at the start of a line or after whitespace/bracket, followed by the query.
  const TRIGGER = /(^|[\s (\[{>"'])@([^\s @]{1,40})$/;

  // Elements we add live under <html>, outside <body>, so they are never part
  // of the sent message; ui-cleanup removes them anyway before every send.
  function uiHost() {
    return doc.documentElement;
  }

  /* ---------------------------------------------------------------- */
  /* @mention                                                           */
  /* ---------------------------------------------------------------- */

  let popup = null;
  let items = [];
  let active = 0;
  let requestId = 0;

  function caretContext() {
    const selection = doc.getSelection();
    if (!selection || selection.rangeCount === 0 || !selection.isCollapsed) {
      return null;
    }
    const node = selection.anchorNode;
    const offset = selection.anchorOffset;
    if (!node || node.nodeType !== Node.TEXT_NODE) {
      return null;
    }
    if (node.parentElement && node.parentElement.closest("a")) {
      return null;
    }
    const match = node.data.slice(0, offset).match(TRIGGER);
    if (!match) {
      return null;
    }
    return { node, start: offset - match[2].length - 1, end: offset, query: match[2] };
  }

  function closePopup() {
    requestId++;
    items = [];
    if (popup) {
      popup.remove();
      popup = null;
    }
  }

  function ensurePopup() {
    if (popup) {
      return popup;
    }
    popup = doc.createElement("div");
    popup.setAttribute("contenteditable", "false");
    popup.setAttribute(UI_ATTR, "mention");
    Object.assign(popup.style, {
      position: "fixed",
      zIndex: "2147483647",
      background: "#ffffff",
      color: "#1f1f1f",
      border: "1px solid #c8c8c8",
      borderRadius: "6px",
      boxShadow: "0 4px 14px rgba(0, 0, 0, 0.18)",
      font: "13px/1.35 system-ui, sans-serif",
      minWidth: "260px",
      maxWidth: "440px",
      padding: "4px 0",
      userSelect: "none",
      cursor: "default",
    });
    // mousedown + preventDefault keeps the caret in the editor.
    popup.addEventListener("mousedown", (event) => {
      event.preventDefault();
      event.stopPropagation();
      const row = event.target.closest("[data-index]");
      if (row) {
        choose(Number(row.dataset.index));
      }
    });
    uiHost().appendChild(popup);
    return popup;
  }

  function render() {
    const box = ensurePopup();
    box.textContent = "";
    items.forEach((item, index) => {
      const row = doc.createElement("div");
      row.dataset.index = String(index);
      Object.assign(row.style, {
        padding: "5px 12px",
        background: index === active ? "#e5f1fb" : "transparent",
        whiteSpace: "nowrap",
        overflow: "hidden",
        textOverflow: "ellipsis",
      });
      const name = doc.createElement("div");
      name.textContent = item.name;
      name.style.fontWeight = "600";
      const email = doc.createElement("div");
      email.textContent = item.email;
      Object.assign(email.style, { fontSize: "11.5px", color: "#666666" });
      row.append(name, email);
      row.addEventListener("mouseenter", () => {
        if (active !== index) {
          active = index;
          render();
        }
      });
      box.appendChild(row);
    });
  }

  function position(context) {
    if (!popup) {
      return;
    }
    const range = doc.createRange();
    range.setStart(context.node, context.start);
    range.setEnd(context.node, context.end);
    const rects = range.getClientRects();
    const rect = rects.length
      ? rects[rects.length - 1]
      : context.node.parentElement.getBoundingClientRect();
    const height = popup.offsetHeight;
    const width = popup.offsetWidth;
    let top = rect.bottom + 4;
    if (top + height > window.innerHeight && rect.top - height - 4 > 0) {
      top = rect.top - height - 4;
    }
    const left = Math.max(4, Math.min(rect.left, window.innerWidth - width - 4));
    popup.style.top = `${top}px`;
    popup.style.left = `${left}px`;
  }

  async function update() {
    const context = caretContext();
    if (!context) {
      closePopup();
      return;
    }
    const id = ++requestId;
    let results;
    try {
      results = await api.runtime.sendMessage({ type: "mention-search", query: context.query });
    } catch (error) {
      console.error("[outlook-mode] mention search failed", error);
      closePopup();
      return;
    }
    if (id !== requestId) {
      return; // a newer keystroke already started another search
    }
    if (!results || !results.length) {
      closePopup();
      return;
    }
    items = results;
    active = 0;
    render();
    position(context);
  }

  function escapeHtml(text) {
    return text
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  const MENTION_EMAIL = /^[^\s@<>(),;:"'\\]+@[^\s@<>(),;:"'\\]+\.[^\s@<>(),;:"'\\]+$/;

  async function choose(index) {
    const item = items[index];
    const context = caretContext();
    closePopup();
    if (!item || !context || !MENTION_EMAIL.test(String(item.email || ""))) {
      return; // only a plain address goes into the link
    }
    const selection = doc.getSelection();
    const range = doc.createRange();
    range.setStart(context.node, context.start);
    range.setEnd(context.node, context.end);
    selection.removeAllRanges();
    selection.addRange(range);
    const html = `<a href="mailto:${escapeHtml(item.email)}">@${escapeHtml(String(item.label || item.email))}</a>&nbsp;`;
    doc.execCommand("insertHTML", false, html);
    try {
      await api.runtime.sendMessage({ type: "mention-add", name: item.name, email: item.email });
    } catch (error) {
      console.error("[outlook-mode] could not add recipient", error);
    }
  }

  let sendLaterKeyAt = 0;

  doc.addEventListener(
    "keydown",
    (event) => {
      // Ctrl+Shift+Enter = Send Later: Undo Send skips its countdown for it.
      if (event.key === "Enter" && event.shiftKey && (event.ctrlKey || event.metaKey)) {
        sendLaterKeyAt = Date.now();
      }
      if (!popup || !items.length) {
        return;
      }
      if (event.key === "ArrowDown") {
        active = (active + 1) % items.length;
        render();
      } else if (event.key === "ArrowUp") {
        active = (active - 1 + items.length) % items.length;
        render();
      } else if (event.key === "Enter" || event.key === "Tab") {
        choose(active);
      } else if (event.key === "Escape") {
        closePopup();
      } else {
        return;
      }
      event.preventDefault();
      event.stopPropagation();
    },
    true
  );

  doc.addEventListener("input", () => update(), true);
  doc.addEventListener(
    "keyup",
    (event) => {
      if (["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) {
        update();
      }
    },
    true
  );
  doc.addEventListener(
    "mousedown",
    (event) => {
      if (popup && !popup.contains(event.target)) {
        closePopup();
      }
    },
    true
  );
  window.addEventListener("blur", closePopup);

  /* ---------------------------------------------------------------- */
  /* Undo Send banner                                                   */
  /* ---------------------------------------------------------------- */

  let banner = null;
  let bannerTimer = null;

  function hideBanner() {
    clearInterval(bannerTimer);
    bannerTimer = null;
    if (banner) {
      banner.remove();
      banner = null;
    }
  }

  function bannerButton(label, primary, onClick) {
    const button = doc.createElement("button");
    button.type = "button";
    button.textContent = label;
    Object.assign(button.style, {
      font: "600 13px system-ui, sans-serif",
      padding: "5px 14px",
      marginLeft: "8px",
      borderRadius: "4px",
      border: primary ? "1px solid #ffffff" : "1px solid transparent",
      background: primary ? "#ffffff" : "transparent",
      color: primary ? "#0f6cbd" : "#ffffff",
      cursor: "pointer",
    });
    button.addEventListener("mousedown", (event) => event.preventDefault());
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      onClick();
    });
    return button;
  }

  function showBanner(seconds) {
    hideBanner();
    closePopup();
    banner = doc.createElement("div");
    banner.setAttribute("contenteditable", "false");
    banner.setAttribute(UI_ATTR, "undo-send");
    Object.assign(banner.style, {
      position: "fixed",
      top: "0",
      left: "0",
      right: "0",
      zIndex: "2147483647",
      display: "flex",
      alignItems: "center",
      gap: "4px",
      padding: "8px 14px",
      background: "#0f6cbd",
      color: "#ffffff",
      font: "14px/1.3 system-ui, sans-serif",
      boxShadow: "0 2px 8px rgba(0, 0, 0, 0.25)",
      userSelect: "none",
    });
    const text = doc.createElement("span");
    text.style.flex = "1";
    const endsAt = Date.now() + seconds * 1000;
    const tick = () => {
      const left = Math.max(0, Math.ceil((endsAt - Date.now()) / 1000));
      text.textContent = left > 0 ? `✉ Sending in ${left}s…` : "✉ Sending…";
    };
    tick();
    bannerTimer = setInterval(tick, 250);
    banner.append(
      text,
      bannerButton("Undo", true, () => {
        hideBanner();
        api.runtime.sendMessage({ type: "undo-cancel" }).catch(() => {});
      }),
      bannerButton("Send now", false, () => {
        text.textContent = "✉ Sending…";
        api.runtime.sendMessage({ type: "undo-send-now" }).catch(() => {});
      })
    );
    uiHost().appendChild(banner);
  }

  function cleanupUi() {
    closePopup();
    hideBanner();
    for (const el of doc.querySelectorAll(`[${UI_ATTR}]`)) {
      el.remove();
    }
  }

  api.runtime.onMessage.addListener((message) => {
    if (!message || typeof message.type !== "string") {
      return undefined;
    }
    if (message.type === "ui-cleanup" || message.type === "mention-cleanup") {
      cleanupUi();
      return Promise.resolve({ ok: true, sendLaterKeyAt });
    }
    if (message.type === "undo-show") {
      showBanner(Number(message.seconds) || 10);
      return Promise.resolve({ ok: true });
    }
    if (message.type === "undo-hide") {
      hideBanner();
      return Promise.resolve({ ok: true });
    }
    return undefined;
  });

  /* ---------------------------------------------------------------- */
  /* Outlook-style reply / forward header                              */
  /* ---------------------------------------------------------------- */

  function buildHeader(header) {
    const wrap = doc.createElement("div");
    wrap.style.cssText =
      "border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0 0 0;margin-top:6px";
    const p = doc.createElement("p");
    p.style.margin = "0";
    const rows = [
      ["From", header.from],
      ["Sent", header.sent],
      ["To", header.to],
      ["Cc", header.cc],
      ["Subject", header.subject],
    ].filter(([, value]) => value);
    rows.forEach(([label, value], index) => {
      if (index) {
        p.appendChild(doc.createElement("br"));
      }
      const bold = doc.createElement("b");
      bold.textContent = `${label}:`;
      p.append(bold, doc.createTextNode(` ${value}`));
    });
    wrap.appendChild(p);
    return wrap;
  }

  function transformReply(info) {
    const prefix = doc.querySelector("div.moz-cite-prefix");
    const quote = doc.querySelector("blockquote[type=cite]");
    if (!prefix && !quote) {
      return false;
    }
    if (prefix && info.settings.replyHeader) {
      prefix.replaceWith(buildHeader(info.header));
    }
    if (quote && info.settings.removeQuoteBar) {
      const plain = doc.createElement("div");
      while (quote.firstChild) {
        plain.appendChild(quote.firstChild);
      }
      quote.replaceWith(plain);
    }
    return true;
  }

  function transformForward(info) {
    const table = doc.querySelector("div.moz-forward-container table.moz-email-headers-table");
    if (!table) {
      return false;
    }
    if (!info.settings.replyHeader) {
      return true;
    }
    const container = table.closest("div.moz-forward-container");
    // Drop the "-------- Forwarded Message --------" line and the <br>s before the table.
    while (container.firstChild && container.firstChild !== table) {
      container.firstChild.remove();
    }
    table.replaceWith(buildHeader(info.header));
    return true;
  }

  async function initReplyHeader() {
    let info;
    try {
      info = await api.runtime.sendMessage({ type: "compose-info" });
    } catch (error) {
      console.error("[outlook-mode] compose info failed", error);
      return;
    }
    if (!info || !info.header) {
      return;
    }
    if (!info.settings.replyHeader && !info.settings.removeQuoteBar) {
      return;
    }
    const run = () => (info.type === "forward" ? transformForward(info) : transformReply(info));
    if (run()) {
      return;
    }
    // The quoted text may be inserted a moment after the editor loads.
    const observer = new MutationObserver(() => {
      if (run()) {
        observer.disconnect();
      }
    });
    observer.observe(doc.body || doc.documentElement, { childList: true, subtree: true });
    setTimeout(() => observer.disconnect(), 5000);
  }

  api.runtime.sendMessage({ type: "compose-hello" }).catch(() => {});
  initReplyHeader();
})();
