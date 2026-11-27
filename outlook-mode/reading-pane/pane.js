"use strict";

// Outlook-style reading pane for the Thunderbird Conversations view, and the
// compose dock for new messages.
//
// Loaded by experiments/outlookPane into Conversations' page, or into
// dock.html (the dock at the bottom of the 3-pane tab, bridge.dock is set).
// The experiment provides window.outlookModeBridge for everything privileged
// (reply data, address lookup, drafts, sending).
//
// - Newest message on top (pane.css).
// - Reply / Reply all / Forward open an inline composer at the top of the
//   conversation instead of a compose window: Conversations' buttons,
//   Ctrl+R / Ctrl+Shift+R / Ctrl+L on a message, and Thunderbird's own
//   commands (routed here by the experiment).
// - In the dock: one tab per new message, the same composer.
//
// This page runs with chrome privileges, so message HTML is never put into
// it: the quoted original is added to the outgoing message by the experiment.

(function () {
  const bridge = window.outlookModeBridge;
  if (!bridge || window.outlookModePane) {
    return;
  }

  const doc = document;
  const DOCK = Boolean(bridge.dock);
  const REPLY_TYPES = ["reply", "replyAll", "forward"];
  const TYPE_LABELS = { new: "New message", reply: "Reply", replyAll: "Reply all", forward: "Forward" };
  // Ctrl/Cmd combinations the composer keeps for text editing; any other one
  // is kept from Thunderbird's shortcuts (Ctrl+D would delete the message).
  const EDIT_KEYS = new Set([
    "a", "c", "v", "x", "z", "y", "b", "i", "u", "k", "enter", "tab", "home", "end",
    "arrowleft", "arrowright", "arrowup", "arrowdown", "backspace", "delete",
    "pageup", "pagedown", "control", "shift", "meta", "alt", "+", "-", "=", "0",
  ]);

  let config = bridge.config();
  let composer = null;
  let openSeq = 0;
  let syncTimer = 0;
  let toastTimer = 0;
  const cleanups = [];

  /* ---------------------------------------------------------------- */
  /* Helpers                                                           */
  /* ---------------------------------------------------------------- */

  function el(tag, props = {}, ...children) {
    const node = doc.createElement(tag);
    for (const [key, value] of Object.entries(props)) {
      if (key === "class") {
        node.className = value;
      } else if (key === "text") {
        node.textContent = value;
      } else if (key.startsWith("on")) {
        node.addEventListener(key.slice(2), value);
      } else if (value === true) {
        node.setAttribute(key, "");
      } else if (value !== false && value != null) {
        node.setAttribute(key, value);
      }
    }
    node.append(...children.filter((child) => child != null));
    return node;
  }

  function listen(target, type, handler, options) {
    target.addEventListener(type, handler, options);
    cleanups.push(() => target.removeEventListener(type, handler, options));
  }

  function errorText(error) {
    return String(error && error.message ? error.message : error);
  }

  function formatSize(bytes) {
    if (!bytes) {
      return "";
    }
    if (bytes < 1024 * 1024) {
      return `${Math.max(1, Math.round(bytes / 1024))} KB`;
    }
    return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
  }

  /* ---------------------------------------------------------------- */
  /* Images                                                            */
  /* ---------------------------------------------------------------- */

  // Conversations' page shows only chrome: and blob: images (its CSP), so
  // data: images (signature, inserted pictures) are shown through blob: URLs,
  // with the data kept in data-ol-src; the experiment puts it back when the
  // message is sent.
  const MAX_PICTURE_BYTES = 5 * 1024 * 1024;
  const blobUrls = new Map(); // data: URL -> blob: URL

  function dataUrlToBlob(dataUrl) {
    const comma = dataUrl.indexOf(",");
    const head = dataUrl.slice(0, comma);
    const body = dataUrl.slice(comma + 1);
    const type = (head.match(/^data:([^;,]+)/) || [])[1] || "application/octet-stream";
    const binary = /;base64/i.test(head) ? atob(body) : decodeURIComponent(body);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i);
    }
    return new Blob([bytes], { type });
  }

  function liveImages(root) {
    for (const img of root.querySelectorAll("img")) {
      const data = img.getAttribute("data-ol-src") || img.getAttribute("src") || "";
      if (!data.startsWith("data:")) {
        continue;
      }
      let url = blobUrls.get(data);
      if (!url) {
        try {
          url = URL.createObjectURL(dataUrlToBlob(data));
        } catch (error) {
          continue; // not valid data
        }
        blobUrls.set(data, url);
      }
      if (img.getAttribute("src") !== url) {
        img.setAttribute("data-ol-src", data);
        img.setAttribute("src", url);
      }
    }
  }

  // Back to data: URLs, for a compose window (Pop out).
  function deadImages(root) {
    for (const img of root.querySelectorAll("img[data-ol-src]")) {
      img.setAttribute("src", img.getAttribute("data-ol-src"));
      img.removeAttribute("data-ol-src");
    }
  }

  function readDataUrl(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error);
      reader.readAsDataURL(file);
    });
  }

  // Raster formats only (no SVG, which can hold script), recognised by
  // their first bytes rather than by the file name or the type it claims.
  const PICTURE_SIGNATURES = [
    ["image/png", [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]],
    ["image/jpeg", [0xff, 0xd8, 0xff]],
    ["image/gif", [0x47, 0x49, 0x46, 0x38]],
    ["image/bmp", [0x42, 0x4d]],
    ["image/webp", [0x52, 0x49, 0x46, 0x46, null, null, null, null, 0x57, 0x45, 0x42, 0x50]],
  ];

  async function pictureType(file) {
    const head = new Uint8Array(await file.slice(0, 12).arrayBuffer());
    const found = PICTURE_SIGNATURES.find(([, bytes]) => bytes.every((b, i) => b === null || head[i] === b));
    return found ? found[0] : null;
  }

  // Pictures from the 🖼 button or the clipboard go into the text at the caret.
  async function insertPictures(c, files) {
    for (const file of files) {
      const type = await pictureType(file);
      if (!type) {
        if (file.type) {
          toast(`${file.name || "That file"} is not a PNG, JPEG, GIF, BMP or WebP picture; attach it instead.`, true);
        }
        continue;
      }
      if (file.size > MAX_PICTURE_BYTES) {
        toast(`${file.name || "The picture"} is larger than 5 MB; attach it instead.`, true);
        continue;
      }
      // The data: URL gets the detected type, not the one the file claims.
      const data = await readDataUrl(file);
      const src = `data:${type};base64,${data.slice(data.indexOf(",") + 1)}`;
      const img = el("img", { src, alt: file.name || "", style: "max-width: 100%;" });
      c.editor.focus();
      doc.execCommand("insertHTML", false, img.outerHTML); // attribute values are escaped by outerHTML
    }
    liveImages(c.editor);
    markDirty(c);
  }

  function toast(text, isError = false) {
    let node = doc.getElementById("ol-toast");
    if (!node) {
      node = el("div", { id: "ol-toast" });
      doc.body.appendChild(node);
    }
    node.textContent = text;
    node.classList.toggle("ol-error", isError);
    node.classList.add("ol-visible");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => node.classList.remove("ol-visible"), isError ? 6000 : 2500);
  }

  // Conversations' message ids, oldest first.
  function conversationIds() {
    return messageItems()
      .map(convIdOfLi)
      .filter((id) => id > 0);
  }

  function convIdOfLi(li) {
    const node = li && li.querySelector("message-header-options[msgid]");
    return node ? Number(node.getAttribute("msgid")) : 0;
  }

  /* ---------------------------------------------------------------- */
  /* Messages in date order                                            */
  /* ---------------------------------------------------------------- */

  // Conversations sorts a conversation by date when it shows it, but a
  // message that comes in while it is shown (new mail, or a reply in Sent
  // Items that Thunderbird's index only now knows) is added at the end of
  // its list, whatever its date. With newest first, such an older message
  // showed above newer ones. So the messages are ordered by their dates
  // here: messageItems() is the list oldest first, and applyOrder() gives
  // each message its place on screen (CSS order; the list is a flex box,
  // pane.css), marking the newest one (data-ol-newest). Conversations' own
  // nodes are not moved.
  const messageDates = new Map(); // conversation id -> date (ms)

  function messageDate(li) {
    const id = convIdOfLi(li);
    if (!id) {
      return null;
    }
    if (!messageDates.has(id)) {
      const date = typeof bridge.dateForConvId === "function" ? bridge.dateForConvId(id) : null;
      if (!Number.isFinite(date)) {
        return null; // asked again next time
      }
      messageDates.set(id, date);
    }
    return messageDates.get(id);
  }

  // The messages, oldest first; in the page's order while a date is unknown.
  function messageItems() {
    const items = [...doc.querySelectorAll("div#messageList > ul > li.message")];
    const dated = items.map((li, index) => ({ li, index, date: messageDate(li) }));
    if (dated.some((item) => item.date === null)) {
      return items;
    }
    return dated.sort((a, b) => a.date - b.date || a.index - b.index).map((item) => item.li);
  }

  function applyOrder() {
    const items = messageItems();
    items.forEach((li, index) => {
      if (li.style.order !== String(index)) {
        li.style.order = String(index);
      }
      li.toggleAttribute("data-ol-newest", index === items.length - 1);
    });
  }

  function clearOrder() {
    for (const li of doc.querySelectorAll("li.message")) {
      li.style.removeProperty("order");
      li.removeAttribute("data-ol-newest");
    }
  }

  /* ---------------------------------------------------------------- */
  /* Configuration                                                     */
  /* ---------------------------------------------------------------- */

  const stylesheet = el("link", {
    rel: "stylesheet",
    href: `chrome://outlookmode/content/pane.css?${config.loadId}`,
  });
  doc.head.appendChild(stylesheet);

  function applyConfig() {
    if (DOCK) {
      return; // open new messages stay when the setting is turned off
    }
    const root = doc.documentElement;
    root.toggleAttribute("data-ol-newest-first", Boolean(config.newestFirst));
    root.toggleAttribute("data-ol-latest-only", latestOnly());
    applyOrder();
    expandMessages(conversationIds());
    updateEarlier(conversationIds());
    root.toggleAttribute("data-ol-inline-reply", Boolean(config.inlineReply));
    root.toggleAttribute("data-ol-pin-button", Boolean(config.pinButton));
    // White like Outlook, or dark while Conversations' ◐ is on.
    root.toggleAttribute("data-ol-light", Boolean(config.lightPane) && !config.darkReader);
    root.toggleAttribute("data-ol-dark", Boolean(config.lightPane) && Boolean(config.darkReader));
    root.toggleAttribute("data-ol-outlook-header", Boolean(config.outlookHeader));
    root.toggleAttribute("data-ol-attachments-top", Boolean(config.attachmentsTop));
    updatePinButtons();
    updateSaveButton();
    updateHeaders();
    updateMessageFrames();
    updateQuoteFolds();
    if (!config.inlineReply && composer) {
      closeComposer({ keepDraft: true });
    }
  }

  /* ---------------------------------------------------------------- */
  /* Every message open                                                */
  /* ---------------------------------------------------------------- */

  // Like Outlook, every mail of the conversation is shown on its own and
  // open, with its own attachments (Conversations lists them in an open
  // message), instead of the newest one quoting the rest. Conversations'
  // own expand-all button is a toggle (it closes everything while one
  // message is open), so each closed message is opened through its store
  // with the plain "messages/msgExpand" action: only what the page shows
  // changes, and nothing is marked read. Without that store, a click on the
  // message's header does the same as the user's click.
  // Conversations closes earlier messages again whenever it shows the
  // conversation anew (selected again, new mail): those are opened again.
  // A message the user closes (a click on its header, the "o" key on it, a
  // click on the conversation's collapse-all button) stays closed until
  // another conversation is shown.
  const USER_CLOSE_MS = 1500;
  let openIds = new Set(); // seen open
  let userClosedIds = new Set();
  let closingIds = new Set(); // the messages the user's last such input was on
  let closingAt = -Infinity;
  let ownClick = false; // a header click of this file, not the user's

  function latestOnly() {
    return Boolean(config.latestOnly) && !config.expandAll;
  }

  if (!DOCK) {
    const noteInput = (event) => {
      if (ownClick) {
        return;
      }
      const path = event.composedPath();
      let ids = [];
      if (event.type === "click" && path.some((node) => node.localName === "conversation-header")) {
        ids = conversationIds();
      } else if (event.type === "click" ? path.some((node) => node.classList?.contains("messageHeader")) : event.key === "o" && !event.ctrlKey && !event.metaKey && !event.altKey) {
        const id = convIdOfLi(path.find((node) => node.matches?.("li.message")));
        ids = id ? [id] : [];
      }
      if (ids.length) {
        closingIds = new Set(ids);
        closingAt = performance.now();
      }
    };
    listen(doc, "click", noteInput, true);
    listen(doc, "keydown", noteInput, true);
  }

  // Each open message makes Conversations load it into its frame. A mail of
  // a Microsoft Graph account that is not stored on this computer yet is
  // downloaded first, and loading another message meanwhile ends that
  // download ("get message: NS_BASE_STREAM_CLOSED"): the frame then shows
  // an empty mail. So messages are opened one at a time, newest first (the
  // top of the pane), Conversations' own loads included; one still loading
  // after LOAD_WAIT_MS no longer holds the rest (a slow Graph download
  // usually ends by itself before that, with the mail or with an error;
  // starting another load earlier would cut it off). A message that loaded
  // empty (Conversations opens several unread ones at once by itself, and
  // fetches others in the background) is loaded again, the way Conversations
  // reloads a message for "show remote content": at once, then after
  // RELOAD_DELAYS_MS, so a slow server has time to recover. Meanwhile the
  // message says it is loading; one that still could not be loaded says so
  // and is loaded again when clicked.
  const MAX_LOADING = 1;
  const LOAD_WAIT_MS = 60000;
  const RELOAD_DELAYS_MS = [0, 3000, 10000];
  const SHOW_LOADING_MS = 1000;
  const MAX_EXPAND_TRIES = 3;
  let loadStarted = new Map(); // id -> when it was seen loading or asked to load here
  let reloads = new Map(); // id -> { count, doc (the empty document), at }
  let expandTries = new Map(); // id -> times asked to open
  let expandTimer = 0;

  function frameLoaded(li) {
    const frame = li.querySelector("iframe");
    if (!frame) {
      return true; // nothing is loading yet
    }
    try {
      const frameDoc = frame.contentDocument;
      const reload = reloads.get(convIdOfLi(li));
      return Boolean(
        frameDoc &&
          frameDoc.URL !== "about:blank" &&
          frameDoc.readyState === "complete" &&
          (!reload || reload.doc !== frameDoc) // the reload has replaced it
      );
    } catch (error) {
      return false;
    }
  }

  // A loaded message with nothing in it (a mail that is only attachments
  // has its list of them).
  function frameEmpty(li) {
    if (li.querySelector("conv-attachment-list")) {
      return false;
    }
    try {
      const body = li.querySelector("iframe")?.contentDocument?.body;
      return Boolean(body) && !body.textContent.trim() && !body.querySelector("img, table, svg, video, object, embed, iframe");
    } catch (error) {
      return false;
    }
  }

  function reloadMessage(dispatch, id) {
    // Two renders: Conversations loads the message again when it goes from
    // "has remote content" back to not.
    dispatch({ type: "messages/setHasRemoteContent", payload: { id, hasRemoteContent: true } });
    setTimeout(() => dispatch({ type: "messages/setHasRemoteContent", payload: { id, hasRemoteContent: false } }), 50);
  }

  function expandMessages(ids) {
    clearTimeout(expandTimer);
    if (!ids.some((id) => openIds.has(id) || userClosedIds.has(id) || loadStarted.has(id))) {
      openIds = new Set(); // another conversation
      userClosedIds = new Set();
      loadStarted = new Map();
      reloads = new Map();
      expandTries = new Map();
    }
    const now = performance.now();
    const recent = now - closingAt < USER_CLOSE_MS;
    const dispatch = customElements.get("conversation-header")?.dispatch;
    const todo = []; // [id, header, "expand" | "reload"], oldest first
    let loading = 0;
    let waiting = false; // a reload waits for its delay
    for (const li of messageItems()) {
      const id = convIdOfLi(li);
      const header = li.querySelector(".messageHeader");
      if (!id || !header) {
        continue;
      }
      const open = header.classList.contains("expanded");
      if (open) {
        openIds.add(id);
        userClosedIds.delete(id);
      } else if (openIds.delete(id) && recent && closingIds.has(id)) {
        userClosedIds.add(id);
      }
      let state = "";
      if (!frameLoaded(li)) {
        if (open || loadStarted.has(id)) {
          if (!loadStarted.has(id)) {
            loadStarted.set(id, now);
          }
          if (now - loadStarted.get(id) < LOAD_WAIT_MS) {
            loading++;
            // At once for a message loaded again (it already failed once).
            state = open && (reloads.has(id) || now - loadStarted.get(id) >= SHOW_LOADING_MS) ? "loading" : "";
          }
        }
      } else if (open && frameEmpty(li) && typeof dispatch === "function") {
        const reload = reloads.get(id) || { count: 0, at: 0 };
        if (reload.count >= RELOAD_DELAYS_MS.length) {
          state = "failed";
        } else {
          state = "loading";
          if (now - reload.at >= RELOAD_DELAYS_MS[reload.count]) {
            todo.push([id, header, "reload"]);
          } else {
            waiting = true;
          }
        }
      }
      if (state) {
        li.setAttribute("data-ol-load", state);
      } else {
        li.removeAttribute("data-ol-load");
      }
      if (config.expandAll && !open && !userClosedIds.has(id) && (expandTries.get(id) || 0) < MAX_EXPAND_TRIES) {
        todo.push([id, header, "expand"]);
      }
    }
    for (const [id, header, what] of todo.reverse()) {
      if (loading >= MAX_LOADING) {
        break;
      }
      loading++;
      loadStarted.set(id, now);
      if (what === "reload") {
        const frame = header.closest("li")?.querySelector("iframe");
        reloads.set(id, { count: (reloads.get(id)?.count || 0) + 1, doc: frame?.contentDocument, at: now });
        reloadMessage(dispatch, id);
        continue;
      }
      openIds.add(id); // asked; it shows as open on the next update
      expandTries.set(id, (expandTries.get(id) || 0) + 1);
      if (typeof dispatch === "function") {
        dispatch({ type: "messages/msgExpand", payload: { id, expand: true } });
      } else {
        ownClick = true;
        try {
          header.click();
        } finally {
          ownClick = false;
        }
      }
    }
    // A frame's load changes nothing the page observer sees: look again
    // while something loads or waits.
    if (loading || todo.length || waiting) {
      expandTimer = setTimeout(() => expandMessages(conversationIds()), 200);
    }
  }

  // "Could not be loaded": a click on that line loads the message again.
  if (!DOCK) {
    listen(doc, "click", (event) => {
      const box = event.target.closest?.('[class^="iframewrap"]');
      const li = box?.closest('li.message[data-ol-load="failed"]');
      const id = convIdOfLi(li);
      if (!id) {
        return;
      }
      const reload = reloads.get(id);
      reloads.set(id, { count: 0, doc: reload?.doc, at: 0 });
      li.setAttribute("data-ol-load", "loading");
      expandMessages(conversationIds());
    });
  }

  /* ---------------------------------------------------------------- */
  /* Opened at the newest message                                      */
  /* ---------------------------------------------------------------- */

  // Conversations scrolls to the message it focuses when it shows a
  // conversation (the one selected in the list, or the first unread one),
  // and again once that message has loaded. With the newest message on top,
  // that is somewhere down the conversation, often the last mail opened
  // before. Those are its only window.scrollTo calls, so with newest first
  // they go to the top of the conversation instead: the newest message, as
  // in Outlook. Once the user scrolls, later calls of the same opening
  // (SCROLL_BURST_MS apart at most: frames that finish loading) leave the
  // page where the user put it. Another conversation shown in the page
  // (syncConversation) starts at the top too.
  const SCROLL_BURST_MS = 5000;
  const pageScrollTo = window.scrollTo;
  let lastAutoScroll = -Infinity;
  let userScrolled = false;
  let shownIds = [];

  function showNewest() {
    userScrolled = false;
    lastAutoScroll = performance.now();
    pageScrollTo.call(window, 0, 0);
  }

  if (!DOCK && typeof pageScrollTo === "function") {
    const ownScrollTo = function (...args) {
      if (!config.newestFirst) {
        return pageScrollTo.apply(this, args);
      }
      const now = performance.now();
      if (now - lastAutoScroll > SCROLL_BURST_MS) {
        userScrolled = false; // another opening
      }
      lastAutoScroll = now;
      if (!userScrolled) {
        pageScrollTo.call(window, 0, 0);
      }
      return undefined;
    };
    window.scrollTo = ownScrollTo;
    cleanups.push(() => {
      if (window.scrollTo === ownScrollTo) {
        window.scrollTo = pageScrollTo;
      }
    });
    const SCROLL_KEYS = new Set(["ArrowUp", "ArrowDown", "PageUp", "PageDown", "Home", "End", " "]);
    const noteUserScroll = (event) => {
      if (event.type !== "keydown" || SCROLL_KEYS.has(event.key)) {
        userScrolled = true;
      }
    };
    for (const type of ["wheel", "keydown", "mousedown", "touchstart"]) {
      listen(doc, type, noteUserScroll, { capture: true, passive: true });
    }
  }

  /* ---------------------------------------------------------------- */
  /* Earlier messages                                                  */
  /* ---------------------------------------------------------------- */

  // Like Outlook, show the newest message, which quotes the rest. pane.css
  // hides the earlier collapsed messages (expanded ones, e.g. unread or
  // selected in the thread pane, stay); this link brings them back.
  let earlierButton = null;
  let earlierShown = false;
  let earlierIds = [];

  function updateEarlier(ids) {
    if (!ids.some((id) => earlierIds.includes(id))) {
      earlierShown = false; // another conversation
    }
    earlierIds = ids;
    doc.documentElement.toggleAttribute("data-ol-show-earlier", earlierShown);
    const items = messageItems().slice(0, -1);
    const hidden = items.filter((li) => !li.querySelector("message-footer")).length;
    const list = doc.getElementById("messageList");
    if (!latestOnly() || !hidden || !list) {
      earlierButton?.remove();
      return;
    }
    if (!earlierButton) {
      earlierButton = el("button", {
        id: "ol-earlier",
        onclick: () => {
          earlierShown = !earlierShown;
          updateEarlier(conversationIds());
        },
      });
    }
    earlierButton.textContent = earlierShown
      ? "Hide earlier messages"
      : `Show ${hidden} earlier message${hidden === 1 ? "" : "s"}`;
    if (earlierButton.previousElementSibling !== list) {
      list.after(earlierButton);
    }
  }

  /* ---------------------------------------------------------------- */
  /* Conversations' ◐ (dark message reader)                            */
  /* ---------------------------------------------------------------- */

  // Conversations marks a message's frame "darkReaderEnabled" only when the
  // message loads, so ◐ changes nothing in the messages already shown. With
  // the pane following ◐ (pane.css), do the same to them at once.
  function updateMessageFrames() {
    if (!config.lightPane) {
      return;
    }
    for (const frame of doc.querySelectorAll("li.message iframe")) {
      try {
        frame.contentDocument?.documentElement?.classList.toggle("darkReaderEnabled", Boolean(config.darkReader));
      } catch (error) {
        // not loaded yet: Conversations sets it when it loads
      }
    }
  }

  /* ---------------------------------------------------------------- */
  /* Pictures fit the reading pane                                     */
  /* ---------------------------------------------------------------- */

  // As in Outlook, a picture wider than the pane is shown at the pane's
  // width (aspect ratio kept), whatever the pane's width is. Pictures with a
  // set height and no width keep it. The style goes into each message's
  // frame when it loads, before Conversations measures the frame's height;
  // when the pane's width changes later, the height is measured again.
  const FIT_STYLE_ID = "ol-fit-pictures";
  const FIT_CSS =
    "img { max-width: 100% !important; box-sizing: border-box; }" +
    " img[width], img[style*='width'] { height: auto !important; }";
  const fittedFrames = new WeakSet();
  const frameWidths = new WeakMap();
  const frameResize = new ResizeObserver((entries) => {
    for (const { target } of entries) {
      const width = Math.round(target.getBoundingClientRect().width);
      if (width && frameWidths.get(target) !== width) {
        frameWidths.set(target, width);
        requestAnimationFrame(() => refitFrame(target));
      }
    }
  });

  function fitPictures(frame) {
    try {
      const frameDoc = frame.contentDocument;
      if (!frameDoc || !frameDoc.head || frameDoc.getElementById(FIT_STYLE_ID)) {
        return false;
      }
      const style = frameDoc.createElement("style");
      style.id = FIT_STYLE_ID;
      style.textContent = FIT_CSS;
      frameDoc.head.append(style);
      return true;
    } catch (error) {
      return false; // not loaded yet: done on its load event
    }
  }

  // Mail laid out wider than the pane (newsletters: a table or box of a
  // fixed 600-800 px, often with text over a background picture, which
  // fitting the pictures cannot shrink) is shown smaller as a whole, at the
  // pane's width, as Outlook does. Mail that fits is left at its size. Run
  // again whenever the pane's width changes.
  const FIT_MIN_ZOOM = 0.4;
  const FIT_SLACK = 8; // px of overflow left alone (a long word, a border)

  function fitWidth(frameDoc) {
    const root = frameDoc.documentElement;
    if (root.dataset.olZoom) {
      root.style.removeProperty("zoom");
      delete root.dataset.olZoom;
    }
    const available = root.clientWidth;
    const needed = Math.max(root.scrollWidth, frameDoc.body ? frameDoc.body.scrollWidth : 0);
    if (available < 100 || needed <= available + FIT_SLACK) {
      return;
    }
    const zoom = Math.max(FIT_MIN_ZOOM, Math.floor((available / needed) * 1000) / 1000);
    root.style.setProperty("zoom", String(zoom));
    root.dataset.olZoom = String(zoom);
  }

  // Same measure as Conversations (content height + 3 px), from a zero
  // height so the frame can also get shorter. Done in one go, so nothing
  // is painted in between, and the pane keeps its scroll position.
  function refitFrame(frame) {
    try {
      const frameDoc = frame.contentDocument;
      if (!frameDoc || !frameDoc.documentElement || !frame.style.height || frameDoc.readyState !== "complete") {
        return;
      }
      // A closed message's frame is not shown: measured, it would be 0 high,
      // and Conversations does not measure it again when it opens loaded.
      if (!frame.getClientRects().length || frameDoc.URL === "about:blank") {
        return;
      }
      // Where the pane (and any scrolling box around the frame) is scrolled
      // to, so the reader stays at the same place.
      const scrolled = [];
      for (let node = frame.parentElement; node; node = node.parentElement) {
        if (node.scrollTop > 0) {
          scrolled.push([node, node.scrollTop]);
        }
      }
      if (doc.scrollingElement && !scrolled.some(([node]) => node === doc.scrollingElement)) {
        scrolled.push([doc.scrollingElement, doc.scrollingElement.scrollTop]);
      }
      fitWidth(frameDoc);
      const before = frame.style.height;
      frame.style.height = "0px";
      void frame.offsetHeight; // lay the pane out first, so the frame's own view is 0 high
      const height = frameDoc.documentElement.scrollHeight + 3;
      frame.style.height = Math.abs(parseFloat(before) - height) > 2 ? `${height}px` : before;
      for (const [node, top] of scrolled) {
        node.scrollTop = top;
      }
    } catch (error) {
      // the frame went away
    }
  }

  function updatePictureFit() {
    for (const frame of doc.querySelectorAll("li.message iframe")) {
      if (!fittedFrames.has(frame)) {
        fittedFrames.add(frame);
        frame.addEventListener(
          "load",
          () => {
            fitPictures(frame);
            // After Conversations has set the frame's height: the mail is
            // fitted to the width (fitWidth), then measured again; once
            // more a little later, for layouts that settle late.
            for (const delay of [50, 600]) {
              setTimeout(() => requestAnimationFrame(() => refitFrame(frame)), delay);
            }
          },
          { capture: true }
        );
        frameResize.observe(frame);
      }
      if (fitPictures(frame)) {
        requestAnimationFrame(() => refitFrame(frame));
      }
    }
  }

  /* ---------------------------------------------------------------- */
  /* Quoted earlier mail folded                                        */
  /* ---------------------------------------------------------------- */

  // With every message open, the earlier mail a reply quotes below its text
  // is shown above as a message of its own. Conversations folds only a long
  // <blockquote> (Thunderbird, Gmail), not Outlook's From / Sent block, so
  // every Outlook reply repeats the whole history. Here the quote of each
  // message but the oldest (it may quote mail that is not in the
  // conversation; Conversations leaves it too) goes behind
  // "- show quoted text -". The message's own nodes are only moved into
  // hidden boxes in its own document: nothing of the message is copied, and
  // a message that is nothing but the quote (a forward without text) stays
  // as it is.
  const QUOTE_BOX_ATTR = "data-ol-quoted";
  // Outlook's first line of the quoted header ("From:" in the mail's
  // language).
  const FROM_LABEL = /^\s*(From|Từ|De|Von|Da|Van|Od|От|Från|Fra|Kimden|发件人|發件人|寄件者|差出人|보낸 사람)\s*[:：]/;
  // The line plain-text Outlook mail puts above that header.
  const TEXT_QUOTE_LINE = /^[ \t]*_{20,}[ \t]*$/m;
  const folds = new WeakMap(); // message document -> { boxes, toggle, shown } or null (nothing to fold)
  const foldWatched = new WeakSet(); // frames with a load listener
  const foldPending = new WeakSet(); // frames waiting for Conversations

  function foldActive() {
    return Boolean(config.foldQuotes) && !latestOnly();
  }

  // The underscore line of a plain-text Outlook quote, as its own text node.
  function textQuoteStart(frameDoc) {
    const walker = frameDoc.createTreeWalker(frameDoc.body, 4 /* NodeFilter.SHOW_TEXT */);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const match = TEXT_QUOTE_LINE.exec(node.data);
      if (!match) {
        continue;
      }
      let after = node.data.slice(match.index + match[0].length);
      const ahead = frameDoc.createTreeWalker(frameDoc.body, 4);
      ahead.currentNode = node;
      for (let next = ahead.nextNode(); next && after.length < 200; next = ahead.nextNode()) {
        after += next.data;
      }
      if (FROM_LABEL.test(after)) {
        return match.index ? node.splitText(match.index) : node;
      }
    }
    return null;
  }

  // Where the quoted mail starts: [node, toEnd]. toEnd: everything from
  // there on is the quote (Outlook puts the original below its header);
  // otherwise only that box is, so answers written between quotes stay.
  function quoteStart(frameDoc) {
    const body = frameDoc.body;
    const found = [];
    for (const sel of ["#appendonsend", "#divRplyFwdMsg"]) {
      const node = body.querySelector(sel);
      if (node) {
        found.push([node, true]);
      }
    }
    const header = [...body.querySelectorAll("div[style*='border-top']")].find((node) => FROM_LABEL.test(node.textContent));
    if (header) {
      found.push([header, true]);
    }
    // New Outlook and Outlook on the web, without #divRplyFwdMsg: a line
    // (<hr>) right above the header block ("From:" / "差出人:" ...).
    const rule = [...body.querySelectorAll("hr")].find((node) => {
      const next = node.nextElementSibling;
      return next && FROM_LABEL.test(next.textContent);
    });
    if (rule) {
      found.push([rule, true]);
    }
    for (const sel of [".gmail_quote", ".yahoo_quoted", ".moz-cite-prefix", ".moz-forward-container", "blockquote"]) {
      const node = body.querySelector(sel);
      if (node) {
        found.push([node, false]);
      }
    }
    if (!found.some(([, toEnd]) => toEnd)) {
      const line = textQuoteStart(frameDoc);
      if (line) {
        found.push([line, true]);
      }
    }
    // The first in the message (a box comes before what it holds).
    found.sort(([a], [b]) => (a === b ? 0 : a.compareDocumentPosition(b) & 4 /* FOLLOWING */ ? -1 : 1));
    return found[0] || null;
  }

  function foldQuote(frame, frameDoc) {
    const start = quoteStart(frameDoc);
    if (!start) {
      folds.set(frameDoc, null);
      return;
    }
    const [node, toEnd] = start;
    // A blockquote Conversations folded itself stays its own.
    if (!toEnd && frameDoc.querySelector(".showhidequote")) {
      folds.set(frameDoc, null);
      return;
    }
    const before = frameDoc.createRange();
    before.setStart(frameDoc.body, 0);
    before.setEndBefore(node);
    if (!before.toString().trim() && !before.cloneContents().querySelector("img")) {
      folds.set(frameDoc, null); // nothing but the quote
      return;
    }
    const boxes = [];
    const wrap = (first, last) => {
      const box = frameDoc.createElement("div");
      box.setAttribute(QUOTE_BOX_ATTR, "");
      first.parentNode.insertBefore(box, first);
      for (let next = first; next; ) {
        const following = next === last ? null : next.nextSibling;
        box.append(next);
        next = following;
      }
      boxes.push(box);
    };
    if (toEnd) {
      // From the start to the end of its box, then what follows each box
      // around it, up to the body.
      let first = node;
      for (let parent = node.parentNode; parent; parent = parent.parentNode) {
        if (first) {
          wrap(first, parent.lastChild);
        }
        if (parent === frameDoc.body) {
          break;
        }
        first = parent.nextSibling;
      }
    } else {
      const last = node.matches(".moz-cite-prefix") && node.nextElementSibling ? node.nextElementSibling : node;
      wrap(node, last);
    }
    const toggle = frameDoc.createElement("button");
    toggle.type = "button";
    toggle.className = "ol-quote-toggle";
    toggle.style.cssText =
      "cursor: pointer; user-select: none; background-color: inherit; border: none; padding: 0; margin: 4px 0; color: orange; font-size: small;";
    boxes[0].before(toggle);
    const fold = { boxes, toggle, shown: false };
    toggle.addEventListener("click", (event) => {
      event.preventDefault();
      fold.shown = !fold.shown;
      showFold(fold, frame);
    });
    folds.set(frameDoc, fold);
    showFold(fold, frame);
  }

  function showFold(fold, frame) {
    const active = foldActive();
    for (const box of fold.boxes) {
      box.style.display = fold.shown || !active ? "" : "none";
    }
    fold.toggle.hidden = !active;
    fold.toggle.textContent = fold.shown ? "- hide quoted text -" : "- show quoted text -";
    requestAnimationFrame(() => refitFrame(frame));
  }

  // Conversations sets a message up when it loads (its own quote folding,
  // fonts, then the frame's height); its style sheet comes last.
  function conversationsDone(frameDoc) {
    return [...frameDoc.querySelectorAll("style")].some((style) => style.textContent.includes("span.moz-txt-formfeed"));
  }

  function scheduleFold(frame, tries = 0) {
    let frameDoc;
    try {
      frameDoc = frame.contentDocument;
    } catch (error) {
      return;
    }
    if (!frameDoc || !frame.isConnected) {
      return;
    }
    if (folds.has(frameDoc)) {
      const fold = folds.get(frameDoc);
      if (fold) {
        showFold(fold, frame);
      }
      return;
    }
    if (!foldActive() || (foldPending.has(frame) && !tries)) {
      return;
    }
    const ready = frameDoc.body && frameDoc.URL !== "about:blank" && frameDoc.readyState === "complete";
    if (ready && (conversationsDone(frameDoc) || tries >= 25)) {
      foldPending.delete(frame);
      try {
        foldQuote(frame, frameDoc);
      } catch (error) {
        folds.set(frameDoc, null);
        bridge.reportError("fold quote", errorText(error));
      }
      return;
    }
    if (tries >= 25) {
      foldPending.delete(frame); // tried again when the message loads
      return;
    }
    foldPending.add(frame);
    setTimeout(() => scheduleFold(frame, tries + 1), 200);
  }

  function updateQuoteFolds() {
    const items = messageItems();
    items.forEach((li, index) => {
      const frame = li.querySelector("iframe");
      if (!frame || index === 0) {
        return; // the oldest message keeps its quote
      }
      if (!foldWatched.has(frame)) {
        foldWatched.add(frame);
        frame.addEventListener("load", () => scheduleFold(frame), { capture: true });
      }
      scheduleFold(frame);
    });
  }

  function unfoldAll() {
    for (const frame of doc.querySelectorAll("li.message iframe")) {
      try {
        const fold = folds.get(frame.contentDocument);
        if (fold) {
          fold.boxes.forEach((box) => (box.style.display = ""));
          fold.toggle.remove();
        }
      } catch (error) {
        // the frame went away
      }
    }
  }

  /* ---------------------------------------------------------------- */
  /* Outlook-style message header                                      */
  /* ---------------------------------------------------------------- */

  // Conversations writes "to A, B, and C" (To, Cc and Bcc together, cut
  // short) in an expanded message's header. Outlook has a To line and a Cc
  // line, and the full date; a click shows everyone with their addresses.
  // The lines go right after the header (pane.css hides the "to …" part).
  const DATE_FORMAT = new Intl.DateTimeFormat(undefined, {
    weekday: "short",
    year: "numeric",
    month: "numeric",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  });

  function recipientLine(label, people) {
    if (!people.length) {
      return null;
    }
    const names = people.map((person, index) =>
      el(
        "span",
        { class: "ol-rcpt", title: person.email },
        el("span", { class: "ol-rcpt-name", text: person.name || person.email }),
        person.name ? el("span", { class: "ol-rcpt-email", text: ` <${person.email}>` }) : null,
        index < people.length - 1 ? "; " : ""
      )
    );
    return el(
      "div",
      { class: "ol-rcpt-line" },
      el("span", { class: "ol-rcpt-label", text: label }),
      el("span", { class: "ol-rcpt-list" }, ...names)
    );
  }

  function buildRecipients(info, convId) {
    const total = info.to.length + info.cc.length + info.bcc.length;
    const toggle = el("button", {
      type: "button",
      class: "ol-rcpt-toggle",
      title: "Show everyone, with addresses",
      text: `▾ ${total}`,
    });
    const box = el(
      "div",
      { class: "ol-recipients", "data-conv-id": String(convId) },
      el(
        "div",
        { class: "ol-rcpt-lines" },
        recipientLine("To", info.to),
        recipientLine("Cc", info.cc),
        recipientLine("Bcc", info.bcc)
      ),
      el(
        "div",
        { class: "ol-rcpt-side" },
        el("span", { class: "ol-rcpt-date", text: info.time ? DATE_FORMAT.format(new Date(info.time)) : "" }),
        total > 1 ? toggle : null
      )
    );
    const flip = () => {
      const open = box.classList.toggle("ol-open");
      toggle.textContent = open ? "▴" : `▾ ${total}`;
      toggle.title = open ? "Show less" : "Show everyone, with addresses";
    };
    toggle.addEventListener("click", flip);
    // Clicking the names opens them too, unless text is being selected.
    box.querySelector(".ol-rcpt-lines").addEventListener("click", () => {
      if (!box.classList.contains("ol-open") && String(window.getSelection()).trim() === "") {
        flip();
      }
    });
    return box;
  }

  function updateHeaders() {
    for (const li of doc.querySelectorAll("li.message")) {
      const header = li.querySelector(".messageHeader");
      let box = li.querySelector(".ol-recipients");
      const convId = convIdOfLi(li);
      if (!config.outlookHeader || !header || !header.classList.contains("expanded") || !convId) {
        box?.remove();
        continue;
      }
      if (box && box.dataset.convId === String(convId) && box.previousElementSibling === header) {
        continue;
      }
      const info = bridge.messageDetails(convId);
      box?.remove();
      if (info) {
        header.after(buildRecipients(info, convId));
      }
    }
  }

  /* ---------------------------------------------------------------- */
  /* Save the conversation as .eml files                               */
  /* ---------------------------------------------------------------- */

  // A button in Conversations' toolbar of the conversation (next to Archive
  // and Delete) saves the whole conversation, your replies in Sent Items
  // included, as one .eml file ("<subject> - <today>.eml") whose text is the
  // conversation, newest first. The experiment checks that the messages
  // belong to the conversation shown, and builds the file.
  let saving = false;

  function saveToolbar() {
    const header = doc.querySelector("conversation-header");
    return header?.shadowRoot?.querySelector("conv-actions-buttons")?.shadowRoot || null;
  }

  function updateSaveButton() {
    const bar = saveToolbar();
    if (!bar || bar.querySelector(".ol-save") || typeof bridge.saveConversation !== "function") {
      return;
    }
    const icon = doc.createElement("svg-icon");
    icon.setAttribute("aria-hidden", "true");
    icon.setAttribute("hash", "save_alt");
    const button = el("button", {
      class: "button-flat actions-button ol-save",
      title: "Save the conversation as one .eml file",
      onclick: saveConversation,
    }, icon);
    button.setAttribute("aria-label", "Save the conversation as one .eml file");
    bar.insertBefore(button, bar.querySelector(".archive") || null);
  }

  async function saveConversation(event) {
    event?.preventDefault();
    const ids = conversationIds();
    if (saving || !ids.length) {
      return;
    }
    saving = true;
    toast("Saving the conversation…");
    try {
      const result = await bridge.saveConversation(ids);
      if (result && result.saved) {
        toast(`Saved the conversation (${result.saved} message${result.saved === 1 ? "" : "s"})${result.file ? ` to ${result.file}` : ""}.`);
      } else {
        doc.getElementById("ol-toast")?.classList.remove("ol-visible"); // cancelled
      }
    } catch (error) {
      toast(`Could not save the conversation: ${errorText(error)}`, true);
    } finally {
      saving = false;
    }
  }

  /* ---------------------------------------------------------------- */
  /* Pin button (Conversations' star)                                  */
  /* ---------------------------------------------------------------- */

  // Conversations' star flags the message; Outlook Mode pins instead. The
  // button is React's, so only an attribute and the title are changed on it;
  // pane.css swaps the icon.
  function updatePinButtons() {
    for (const li of doc.querySelectorAll("li.message")) {
      const button = li.querySelector(".messageHeader button.star");
      const convId = convIdOfLi(li);
      if (!button || !convId) {
        continue;
      }
      if (!config.pinButton) {
        button.removeAttribute("data-ol-pinned");
        continue;
      }
      const pinned = Boolean(bridge.isPinned(convId));
      button.toggleAttribute("data-ol-pinned", pinned);
      button.title = pinned ? "Unpin" : "Pin";
    }
  }

  function onStarClick(event) {
    if (!config.pinButton || event.button !== 0) {
      return false;
    }
    const button = event.composedPath().find((node) => node.localName === "button" && node.classList.contains("star"));
    const convId = button && convIdOfLi(button.closest("li.message"));
    if (!convId) {
      return false;
    }
    event.preventDefault();
    event.stopImmediatePropagation(); // no flag, and the header does not toggle
    if (!bridge.togglePin(convId)) {
      toast("Could not pin this message.", true);
      return true;
    }
    // The tag is set by the background a moment later.
    setTimeout(updatePinButtons, 400);
    setTimeout(updatePinButtons, 1500);
    return true;
  }

  /* ---------------------------------------------------------------- */
  /* Taking over Reply / Reply all / Forward                           */
  /* ---------------------------------------------------------------- */

  // Conversations' buttons (header, its "more" menu, footer) are
  // <action-button type="reply|replyAll|forward" msgid="N"> in shadow roots.
  function onClick(event) {
    if (onStarClick(event) || !config.inlineReply || event.button !== 0) {
      return;
    }
    for (const node of event.composedPath()) {
      if (node.localName !== "action-button") {
        continue;
      }
      const type = node.getAttribute("type");
      const convId = Number(node.getAttribute("msgid"));
      if (!REPLY_TYPES.includes(type) || !convId) {
        return;
      }
      event.preventDefault();
      event.stopImmediatePropagation();
      node.dispatchEvent(new CustomEvent("button-clicked")); // closes the "more" menu
      openByConvId(convId, type);
      return;
    }
  }

  function onKeyDown(event) {
    if (!config.inlineReply || !(event.ctrlKey || event.metaKey) || event.altKey) {
      return;
    }
    if (composer && composer.root.contains(event.target)) {
      return;
    }
    const key = event.key.toLowerCase();
    const type =
      key === "r" ? (event.shiftKey ? "replyAll" : "reply") : key === "l" && !event.shiftKey ? "forward" : null;
    const li = type && event.target.closest ? event.target.closest("li.message") : null;
    const convId = convIdOfLi(li);
    if (!convId) {
      return; // Thunderbird's own command runs (and is routed here)
    }
    event.preventDefault();
    event.stopImmediatePropagation();
    openByConvId(convId, type);
  }

  function openByConvId(convId, type) {
    const uri = bridge.uriForConvId(convId);
    if (!uri) {
      toast("This message is no longer available.", true);
      return;
    }
    openComposer(uri, type, convId, { focus: true });
  }

  // Called by the experiment for Thunderbird's commands. Returns true when
  // the reply opens here.
  function openFor(uri, type, options = {}) {
    if (!config.inlineReply || !REPLY_TYPES.includes(type)) {
      return false;
    }
    const ids = conversationIds();
    if (!ids.length) {
      return false;
    }
    if (options.newest) {
      const newestId = ids[ids.length - 1];
      const newestUri = bridge.uriForConvId(newestId);
      if (newestUri) {
        openComposer(newestUri, type, newestId, { focus: true });
        return true;
      }
    }
    const convId = bridge.convIdForUri(uri);
    if (!convId) {
      return false;
    }
    openComposer(uri, type, convId, { focus: true });
    return true;
  }

  /* ---------------------------------------------------------------- */
  /* Composer                                                          */
  /* ---------------------------------------------------------------- */

  async function openComposer(uri, type, convId, { focus = true } = {}) {
    if (composer && composer.uri === uri && (composer.type === type || composer.sending)) {
      if (focus) {
        focusEditor(composer);
      }
      return;
    }
    if (composer) {
      closeComposer({ keepDraft: true });
    }
    const seq = ++openSeq;
    const pending = bridge.pendingSend(uri);
    const saved = pending ? pending.draft : bridge.getDraft(uri);
    if (pending) {
      type = saved.type;
    }
    let data;
    try {
      data = await bridge.prepare(uri, type);
    } catch (error) {
      if (seq === openSeq) {
        toast(`Could not start the reply: ${errorText(error)}`, true);
      }
      return;
    }
    if (seq !== openSeq) {
      return; // another reply was opened meanwhile
    }
    const sameType = Boolean(saved && saved.type === type);
    const c = buildComposer(data, {
      uri,
      convId,
      type,
      identityKey: (saved && saved.identityKey) || data.identityKey,
      to: sameType ? saved.to : data.to,
      cc: sameType ? saved.cc : data.cc,
      bcc: (saved && saved.bcc) || "",
      subject: sameType ? saved.subject : data.subject,
      html: saved ? saved.html : null,
      attachments: sameType ? saved.attachments || [] : data.attachments,
      files: (saved && saved.files) || [],
    });
    composer = c;
    if (DOCK) {
      dockBody.replaceChildren(c.root);
      renderDockBar();
    } else {
      const list = doc.getElementById("messageList");
      list.parentNode.insertBefore(c.root, list);
    }
    fitAddresses(c);
    if (saved && saved.dirty) {
      c.dirty = true;
      setStatus(c, "Draft");
    }
    if (saved && !sameType) {
      markDirty(c);
    }
    if (pending) {
      setSending(c, true);
      showCountdown(c, pending.until);
    }
    if (focus) {
      if (!DOCK) {
        c.root.scrollIntoView({ block: "start" });
      }
      if (type === "new" && !c.to.value.trim()) {
        c.to.focus();
      } else {
        focusEditor(c);
      }
    }
  }

  function buildComposer(data, state) {
    const c = {
      ...state,
      data,
      dirty: false,
      sending: false,
      saveTimer: 0,
      countdownTimer: 0,
    };

    c.kind = el("span", { class: "ol-kind", text: TYPE_LABELS[state.type] });
    c.status = el("span", { class: "ol-status" });

    c.from = el("select", { class: "ol-input", onchange: () => (swapSignature(c), markDirty(c)) });
    for (const identity of data.identities) {
      const option = el("option", { value: identity.key, text: identity.label });
      option.selected = identity.key === state.identityKey;
      c.from.appendChild(option);
    }

    c.to = addressInput(c, state.to);
    c.cc = addressInput(c, state.cc);
    c.bcc = addressInput(c, state.bcc);
    c.subject = el("input", {
      class: "ol-input",
      type: "text",
      oninput: () => {
        markDirty(c);
        if (DOCK) {
          renderDockBar(); // the tab shows the subject
        }
      },
    });
    c.subject.value = state.subject || "";

    const rows = {
      from: row("From", c.from, data.identities.length > 1),
      to: row("To", c.to, true),
      cc: row("Cc", c.cc, Boolean(state.cc) || state.type === "replyAll"),
      bcc: row("Bcc", c.bcc, Boolean(state.bcc)),
      subject: row("Subject", c.subject, state.type === "forward" || state.type === "new"),
    };
    const reveal = (name, input) => {
      const target = rows[name.toLowerCase()];
      const button = el("button", {
        class: `ol-link${target.classList.contains("ol-hidden") ? "" : " ol-hidden"}`,
        text: name,
        title: `Show ${name}`,
        onclick: () => {
          target.classList.remove("ol-hidden");
          button.classList.add("ol-hidden");
          fitAddress(input);
          input.focus();
        },
      });
      return button;
    };
    rows.to.append(
      el("span", { class: "ol-reveal" }, reveal("Cc", c.cc), reveal("Bcc", c.bcc), reveal("Subject", c.subject))
    );

    c.editor = el("div", {
      class: "ol-editor",
      contenteditable: "true",
      role: "textbox",
      "aria-multiline": "true",
      "aria-label": "Message body",
      spellcheck: "true",
    });
    c.editor.style.fontSize = data.fontSize || "11pt";
    if (data.font) {
      c.editor.style.fontFamily = data.font;
    }
    if (state.html != null) {
      c.editor.innerHTML = bridge.sanitize(state.html); // this composer's own earlier content, sanitized
      liveImages(c.editor);
    } else {
      c.editor.append(el("div", {}, el("br")), el("div", {}, el("br")));
      swapSignature(c);
    }
    c.editor.addEventListener("input", (event) => {
      autoList(c, event);
      applyFontSize(c);
      liveImages(c.editor);
      markDirty(c);
      updateMention(c);
    });
    c.editor.addEventListener("paste", (event) => {
      if (c.plainPaste) {
        // Ctrl+Shift+V: only the text is taken (inserted by onComposerKey).
        event.preventDefault();
        c.plainPaste.text = event.clipboardData ? event.clipboardData.getData("text/plain") : "";
        return;
      }
      const files = [...(event.clipboardData?.files || [])].filter((file) => /^image\//.test(file.type));
      if (files.length) {
        event.preventDefault();
        insertPictures(c, files);
      }
    });
    c.editor.addEventListener("keyup", (event) => {
      if (["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) {
        updateMention(c);
      }
    });
    c.editor.addEventListener("blur", () => setTimeout(() => mentionState && mentionState.c === c && hideMention(), 150));

    c.pictureInput = el("input", {
      type: "file",
      accept: "image/*",
      multiple: true,
      class: "ol-hidden",
      onchange: () => {
        insertPictures(c, [...c.pictureInput.files]);
        c.pictureInput.value = "";
      },
    });
    c.fileInput = el("input", {
      type: "file",
      multiple: true,
      class: "ol-hidden",
      onchange: () => {
        addFiles(c, c.fileInput.files);
        c.fileInput.value = "";
      },
    });
    c.attachmentList = el("div", { class: "ol-attachments" });

    const original = data.original || {};
    const originalLines = [
      ["From", original.from],
      ["Sent", original.sent],
      ["To", original.to],
      ["Cc", original.cc],
      ["Subject", original.subject],
    ]
      .filter(([, value]) => value)
      .map(([label, value]) => el("div", {}, el("b", { text: `${label}: ` }), value));
    const originalBox = data.original && el(
      "details",
      { class: "ol-original" },
      el("summary", {
        text:
          state.type === "forward"
            ? "The forwarded message is included below your text"
            : "The original message is quoted below your reply",
      }),
      ...originalLines
    );

    c.sendButton = el("button", { class: "ol-primary", text: "Send", title: "Send (Ctrl+Enter)", onclick: () => send(c) });
    c.discardButton = el("button", { text: "Discard", onclick: () => discard(c) });
    c.popOutButton = el("button", {
      text: "Pop out",
      title: "Continue in a compose window",
      onclick: () => popOut(c),
    });
    c.actions = el("div", { class: "ol-actions" }, c.sendButton, c.discardButton, c.popOutButton, c.status);
    c.countdownText = el("span");
    c.countdown = el(
      "div",
      { class: "ol-countdown ol-hidden" },
      c.countdownText,
      el("button", { text: "Undo", onclick: () => undoSend(c) }),
      el("button", { text: "Send now", onclick: () => sendNow(c) })
    );

    c.root = el(
      "section",
      { class: "ol-composer", "aria-label": TYPE_LABELS[state.type] },
      el("div", { class: "ol-head" }, c.kind),
      rows.from,
      rows.to,
      rows.cc,
      rows.bcc,
      rows.subject,
      toolbar(c),
      c.editor,
      c.attachmentList,
      c.fileInput,
      c.pictureInput,
      originalBox,
      c.actions,
      c.countdown
    );
    c.root.addEventListener("keydown", (event) => onComposerKey(c, event));
    c.root.addEventListener("dragover", (event) => {
      if (event.dataTransfer && [...event.dataTransfer.types].includes("Files")) {
        event.preventDefault();
        event.stopPropagation();
      }
    });
    c.root.addEventListener("drop", (event) => {
      if (event.dataTransfer && event.dataTransfer.files.length) {
        event.preventDefault();
        event.stopPropagation();
        addFiles(c, event.dataTransfer.files);
      }
    });
    renderAttachments(c);
    return c;
  }

  function row(label, input, visible) {
    return el(
      "div",
      { class: `ol-row${visible ? "" : " ol-hidden"}` },
      el("label", { class: "ol-label", text: label }),
      input
    );
  }

  /* ---------------------------------------------------------------- */
  /* Formatting toolbar (Outlook's Format Text tab)                    */
  /* ---------------------------------------------------------------- */

  const ICONS = "chrome://messenger/skin/icons/";
  const FONTS = [
    "Aptos", "Arial", "Calibri", "Cambria", "Courier New", "Georgia", "Segoe UI",
    "Tahoma", "Times New Roman", "Trebuchet MS", "Verdana",
  ];
  const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];
  const BLOCK_STYLES = [
    ["div", "Normal"],
    ["h1", "Heading 1"],
    ["h2", "Heading 2"],
    ["h3", "Heading 3"],
    ["blockquote", "Quote"],
    ["pre", "Code"],
  ];
  // Outlook's standard colours, then a row of greys and darker tones.
  const FONT_COLORS = [
    "#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0",
    "#000000", "#404040", "#7F7F7F", "#A6A6A6", "#D9D9D9", "#FFFFFF", "#1F4E79", "#2E75B6", "#C55A11", "#548235",
  ];
  const HIGHLIGHT_COLORS = [
    "#FFFF00", "#00FF00", "#00FFFF", "#FF00FF", "#0000FF", "#FF0000", "#000080", "#008080",
    "#008000", "#800080", "#800000", "#808000", "#808080", "#C0C0C0", "#000000",
  ];
  const EMOJIS = (
    "😀 😃 😄 😁 😊 🙂 😉 😍 😘 😎 🤔 😅 😂 😢 😮 🙏 👍 👎 👏 🙌 💪 👌 ✌️ 🤝 👋 " +
    "🎉 🎂 🎁 ❤️ 💯 ✅ ❌ ⚠️ ❓ ⭐ 🔥 🚀 📌 📎 📅 ⏰ 📞 💡 ☕ 🍻"
  ).split(" ");
  // Buttons that show whether the text at the caret has their format.
  const STATE_COMMANDS = new Set([
    "bold", "italic", "underline", "strikeThrough", "subscript", "superscript",
    "insertUnorderedList", "insertOrderedList", "justifyLeft", "justifyCenter", "justifyRight", "justifyFull",
  ]);
  // Outlook's keys in the message body (Ctrl+...).
  const FORMAT_KEYS = {
    b: "bold", i: "italic", u: "underline", k: "link",
    e: "justifyCenter", l: "justifyLeft", r: "justifyRight", j: "justifyFull",
    m: "indent", "shift+m": "outdent", "shift+l": "insertUnorderedList",
    "]": "grow", "[": "shrink", "shift+>": "grow", "shift+<": "shrink",
    "=": "subscript", "shift++": "superscript", " ": "removeFormat",
    "shift+f": "fontMenu", "shift+p": "sizeMenu",
  };

  function toolbar(c) {
    c.stateButtons = [];
    c.lastColor = "#FF0000";
    c.lastHighlight = "#FFFF00";
    const button = (command, title, { icon = "", glyph = "", menu = false } = {}) => {
      const node = el("button", {
        type: "button",
        class: `ol-tool ol-tool-${command}${icon ? " ol-icon" : ""}${menu ? " ol-menu" : ""}`,
        title,
        "data-command": command,
        style: icon ? `--ol-icon: url("${icon.includes(":") ? icon : ICONS + icon}")` : null,
        text: glyph,
      });
      // Keep the selection in the editor.
      node.addEventListener("mousedown", (event) => event.preventDefault());
      node.addEventListener("click", () => runTool(c, command, node));
      if (STATE_COMMANDS.has(command)) {
        node.setAttribute("aria-pressed", "false");
        c.stateButtons.push(node);
      }
      return node;
    };
    const split = (main, arrow) => el("span", { class: "ol-split" }, main, arrow);
    const sep = () => el("span", { class: "ol-sep" });
    const select = (name, title, options, onChange) => {
      const node = el("select", { class: `ol-tool-select ol-tool-${name}`, title });
      for (const [value, label] of options) {
        node.append(el("option", { value, text: label }));
      }
      node.addEventListener("change", () => {
        onChange(node.value);
        updateToolState(c);
      });
      // Back to the text with Escape or after a choice.
      node.addEventListener("keydown", (event) => {
        if (event.key === "Escape") {
          restoreSelection(c);
        }
      });
      return node;
    };

    c.fontSelect = select("font", "Font (Ctrl+Shift+F)", [["", "Font"], ...FONTS.map((f) => [f, f])], (value) => {
      if (value) {
        exec(c, "fontName", value, true);
      }
    });
    c.sizeSelect = select(
      "size",
      "Font size (Ctrl+Shift+P)",
      [["", "Size"], ...FONT_SIZES.map((size) => [String(size), String(size)])],
      (value) => value && setFontSize(c, Number(value))
    );
    c.styleSelect = select("style", "Styles", BLOCK_STYLES, (value) => exec(c, "formatBlock", value));

    c.colorButton = button("color", "Font color", { glyph: "A" });
    c.highlightButton = button("highlight", "Text highlight color", { glyph: "ab" });
    c.colorButton.style.setProperty("--ol-swatch", c.lastColor);
    c.highlightButton.style.setProperty("--ol-swatch", c.lastHighlight);

    return el(
      "div",
      { class: "ol-toolbar", role: "toolbar", "aria-label": "Formatting" },
      button("undo", "Undo (Ctrl+Z)", { glyph: "↶" }),
      button("redo", "Redo (Ctrl+Y)", { glyph: "↷" }),
      sep(),
      c.fontSelect,
      c.sizeSelect,
      button("grow", "Increase font size (Ctrl+])", { glyph: "A+" }),
      button("shrink", "Decrease font size (Ctrl+[)", { glyph: "A−" }),
      sep(),
      button("bold", "Bold (Ctrl+B)", { icon: "bold.svg" }),
      button("italic", "Italic (Ctrl+I)", { icon: "italics.svg" }),
      button("underline", "Underline (Ctrl+U)", { icon: "underline.svg" }),
      button("strikeThrough", "Strikethrough", { glyph: "S" }),
      button("subscript", "Subscript (Ctrl+=)", { glyph: "x₂" }),
      button("superscript", "Superscript (Ctrl+Shift++)", { glyph: "x²" }),
      sep(),
      split(c.highlightButton, button("highlightMenu", "Highlight colors", { glyph: "▾", menu: true })),
      split(c.colorButton, button("colorMenu", "Font colors", { glyph: "▾", menu: true })),
      button("removeFormat", "Clear all formatting (Ctrl+Space)", { icon: "remove-text-styling.svg" }),
      sep(),
      button("insertUnorderedList", "Bullets (Ctrl+Shift+L)", { icon: "bullet-list.svg" }),
      button("insertOrderedList", "Numbering", { icon: "number-list.svg" }),
      button("outdent", "Decrease indent (Ctrl+Shift+M)", { icon: "outdent.svg" }),
      button("indent", "Increase indent (Ctrl+M)", { icon: "indent.svg" }),
      sep(),
      button("justifyLeft", "Align left (Ctrl+L)", { icon: "left-align.svg" }),
      button("justifyCenter", "Center (Ctrl+E)", { icon: "center-align.svg" }),
      button("justifyRight", "Align right (Ctrl+R)", { icon: "right-align.svg" }),
      button("justifyFull", "Justify (Ctrl+J)", { icon: "justify.svg" }),
      sep(),
      c.styleSelect,
      sep(),
      button("link", "Insert link (Ctrl+K)", { icon: "chrome://messenger/skin/icons/new/compact/link.svg" }),
      button("picture", "Insert picture (or paste one)", { icon: "image.svg" }),
      button("tableMenu", "Insert table", { icon: "table.svg", menu: true }),
      button("insertHorizontalRule", "Horizontal line", { icon: "hline.svg" }),
      button("emojiMenu", "Emoji", { icon: "smiley.svg", menu: true }),
      button("signatureMenu", "Signature", { glyph: "✍ Signature ▾", menu: true }),
      sep(),
      button("attach", "Attach files (or drop them here)", { icon: "attach.svg" })
    );
  }

  // The editor's selection, kept while a list or a menu has the focus.
  function saveSelection(c) {
    const selection = window.getSelection();
    if (selection && selection.rangeCount && c.editor.contains(selection.getRangeAt(0).commonAncestorContainer)) {
      c.savedRange = selection.getRangeAt(0).cloneRange();
    }
  }

  function restoreSelection(c) {
    c.editor.focus();
    if (c.savedRange) {
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(c.savedRange);
    }
  }

  // styleWithCSS: colours and fonts as inline styles, which mail programs
  // keep; the rest as plain tags (<b>, <i>, <ul>).
  function exec(c, command, value = null, css = false) {
    restoreSelection(c);
    doc.execCommand("styleWithCSS", false, css);
    doc.execCommand(command, false, value);
    doc.execCommand("styleWithCSS", false, false);
    markDirty(c);
    updateToolState(c);
  }

  // Sizes in points, like Outlook: execCommand only knows sizes 1-7, so the
  // text is marked with size 7 and the marks become font-size spans.
  function setFontSize(c, points) {
    c.pendingFontSize = points;
    restoreSelection(c);
    doc.execCommand("styleWithCSS", false, false);
    doc.execCommand("fontSize", false, "7");
    applyFontSize(c);
    markDirty(c);
    updateToolState(c);
  }

  // Also runs on input: with nothing selected, the size applies to what is
  // typed next.
  function applyFontSize(c) {
    const marks = [...c.editor.querySelectorAll('font[size="7"]')];
    if (!marks.length || !c.pendingFontSize) {
      return;
    }
    const spans = marks.map((font) => {
      const span = el("span", { style: `font-size: ${c.pendingFontSize}pt;` });
      for (const inner of font.querySelectorAll("[style]")) {
        inner.style.fontSize = "";
      }
      span.append(...font.childNodes);
      font.replaceWith(span);
      return span;
    });
    const range = doc.createRange();
    range.setStart(spans[0], 0);
    range.setEnd(spans[spans.length - 1], spans[spans.length - 1].childNodes.length);
    if (!range.collapsed) {
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
    }
  }

  function caretElement(c) {
    const selection = window.getSelection();
    let node = selection && selection.rangeCount ? selection.getRangeAt(0).startContainer : null;
    if (node && node.nodeType === Node.TEXT_NODE) {
      node = node.parentElement;
    }
    return node && c.editor.contains(node) ? node : c.editor;
  }

  function currentPoints(c) {
    const px = parseFloat(window.getComputedStyle(caretElement(c)).fontSize) || 14.667;
    return Math.round(px * 0.75 * 2) / 2;
  }

  function stepFontSize(c, direction) {
    const now = currentPoints(c);
    const next =
      direction > 0
        ? FONT_SIZES.find((size) => size > now) || FONT_SIZES[FONT_SIZES.length - 1]
        : [...FONT_SIZES].reverse().find((size) => size < now) || FONT_SIZES[0];
    setFontSize(c, next);
  }

  function queryState(command) {
    try {
      return doc.queryCommandState(command);
    } catch (error) {
      return false;
    }
  }

  // Pressed buttons and the font, size and style at the caret.
  function updateToolState(c) {
    if (!c.stateButtons || c !== composer) {
      return;
    }
    for (const node of c.stateButtons) {
      node.setAttribute("aria-pressed", String(queryState(node.dataset.command)));
    }
    const style = window.getComputedStyle(caretElement(c));
    const family = (style.fontFamily || "").split(",")[0].replace(/["']/g, "").trim();
    c.fontSelect.options[0].textContent = family || "Font";
    c.fontSelect.value = FONTS.includes(family) ? family : "";
    const points = currentPoints(c);
    c.sizeSelect.options[0].textContent = String(points);
    c.sizeSelect.value = FONT_SIZES.includes(points) ? String(points) : "";
    let block = "";
    try {
      block = doc.queryCommandValue("formatBlock");
    } catch (error) {
      // not supported here
    }
    c.styleSelect.value = BLOCK_STYLES.some(([tag]) => tag === block) ? block : "div";
  }

  /* Menus: colours, table size, emoji ------------------------------- */

  let formatMenu = null; // { c, box, onOutside }

  function closeFormatMenu() {
    if (formatMenu) {
      doc.removeEventListener("mousedown", formatMenu.onOutside, true);
      formatMenu.box.remove();
      formatMenu = null;
    }
  }

  function openFormatMenu(c, anchor, content) {
    const wasOpen = formatMenu && formatMenu.anchor === anchor;
    closeFormatMenu();
    if (wasOpen) {
      return; // a second click closes it
    }
    const box = el("div", { class: "ol-popup" }, ...content);
    box.addEventListener("mousedown", (event) => event.preventDefault()); // keep the selection
    c.root.appendChild(box);
    const rootRect = c.root.getBoundingClientRect();
    const rect = anchor.getBoundingClientRect();
    box.style.top = `${rect.bottom - rootRect.top + 2}px`;
    box.style.left = `${Math.max(0, Math.min(rect.left - rootRect.left, rootRect.width - box.offsetWidth))}px`;
    const onOutside = (event) => {
      if (!box.contains(event.target) && event.target !== anchor) {
        closeFormatMenu();
      }
    };
    doc.addEventListener("mousedown", onOutside, true);
    formatMenu = { c, box, anchor, onOutside };
  }

  function swatches(colors, onPick, extra) {
    return el(
      "div",
      { class: "ol-swatches" },
      ...colors.map((color) =>
        el("button", {
          type: "button",
          class: "ol-swatch",
          title: color,
          style: `background: ${color}`,
          onclick: () => {
            closeFormatMenu();
            onPick(color);
          },
        })
      ),
      extra
    );
  }

  function pickColor(c, color) {
    c.lastColor = color;
    c.colorButton.style.setProperty("--ol-swatch", color);
    exec(c, "foreColor", color, true);
  }

  function pickHighlight(c, color) {
    if (color) {
      c.lastHighlight = color;
      c.highlightButton.style.setProperty("--ol-swatch", color);
    }
    exec(c, "hiliteColor", color || "transparent", true);
  }

  function tableMenu(c, anchor) {
    const ROWS = 8;
    const COLS = 10;
    const label = el("div", { class: "ol-popup-label", text: "Insert table" });
    const cells = [];
    const grid = el("div", { class: "ol-table-grid", style: `grid-template-columns: repeat(${COLS}, 16px)` });
    for (let row = 1; row <= ROWS; row++) {
      for (let col = 1; col <= COLS; col++) {
        const cell = el("button", { type: "button", class: "ol-table-cell", "aria-label": `${col} × ${row}` });
        cell.addEventListener("mouseenter", () => {
          label.textContent = `${col} × ${row} Table`;
          for (const other of cells) {
            other.node.classList.toggle("ol-active", other.row <= row && other.col <= col);
          }
        });
        cell.addEventListener("click", () => {
          closeFormatMenu();
          insertTable(c, row, col);
        });
        cells.push({ node: cell, row, col });
        grid.append(cell);
      }
    }
    openFormatMenu(c, anchor, [label, grid]);
  }

  function insertTable(c, rows, cols) {
    const cell = '<td style="border: 1px solid #A6A6A6; padding: 4px 6px; vertical-align: top;"><br></td>';
    const row = `<tr>${cell.repeat(cols)}</tr>`;
    exec(
      c,
      "insertHTML",
      `<table style="border-collapse: collapse; width: 100%;"><tbody>${row.repeat(rows)}</tbody></table><div><br></div>`
    );
  }

  function emojiMenu(c, anchor) {
    openFormatMenu(c, anchor, [
      el(
        "div",
        { class: "ol-emoji-grid" },
        ...EMOJIS.map((emoji) =>
          el("button", {
            type: "button",
            class: "ol-emoji",
            text: emoji,
            onclick: () => {
              closeFormatMenu();
              exec(c, "insertText", emoji);
            },
          })
        )
      ),
    ]);
  }

  function insertLink(c) {
    restoreSelection(c);
    const current = caretElement(c).closest("a");
    const url = window.prompt(
      current ? "Link address (empty to remove the link):" : "Link address:",
      current ? current.getAttribute("href") : "https://"
    );
    if (url === null) {
      return;
    }
    restoreSelection(c);
    let target = url.trim();
    if (!target && current) {
      const range = doc.createRange();
      range.selectNode(current);
      window.getSelection().removeAllRanges();
      window.getSelection().addRange(range);
      doc.execCommand("unlink", false, null);
    } else if (target) {
      if (/^[^\s@:/]+@[^\s@]+\.[^\s@]+$/.test(target)) {
        target = `mailto:${target}`; // an e-mail address
      } else if (!/^[a-z][\w+.-]*:/i.test(target)) {
        target = `https://${target}`;
      }
      if (!/^(https?:|mailto:|ftp:|tel:)/i.test(target)) {
        return;
      }
      if (window.getSelection().isCollapsed) {
        doc.execCommand("insertHTML", false, el("a", { href: target, text: url.trim() }).outerHTML);
      } else {
        doc.execCommand("createLink", false, target);
      }
    }
    markDirty(c);
  }

  function runTool(c, command, anchor) {
    if (c.sending) {
      return;
    }
    switch (command) {
      case "attach":
        c.fileInput.click();
        return;
      case "picture":
        c.pictureInput.click();
        return;
      case "link":
        insertLink(c);
        return;
      case "grow":
      case "shrink":
        stepFontSize(c, command === "grow" ? 1 : -1);
        return;
      case "color":
        pickColor(c, c.lastColor);
        return;
      case "highlight":
        pickHighlight(c, c.lastHighlight);
        return;
      case "colorMenu":
        openFormatMenu(c, anchor, [swatches(FONT_COLORS, (color) => pickColor(c, color))]);
        return;
      case "highlightMenu":
        openFormatMenu(c, anchor, [
          swatches(
            HIGHLIGHT_COLORS,
            (color) => pickHighlight(c, color),
            el("button", {
              type: "button",
              class: "ol-swatch-none",
              text: "No color",
              onclick: () => {
                closeFormatMenu();
                pickHighlight(c, "");
              },
            })
          ),
        ]);
        return;
      case "tableMenu":
        tableMenu(c, anchor);
        return;
      case "emojiMenu":
        emojiMenu(c, anchor);
        return;
      case "signatureMenu":
        openFormatMenu(c, anchor, [
          el("button", {
            type: "button",
            class: "ol-menu-item",
            text: "Insert signature",
            onclick: () => {
              closeFormatMenu();
              swapSignature(c);
              markDirty(c);
            },
          }),
          el("button", {
            type: "button",
            class: "ol-menu-item",
            text: "Signatures… (Outlook Mode preferences)",
            onclick: () => {
              closeFormatMenu();
              bridge.openSettings("signature");
            },
          }),
        ]);
        return;
      case "fontMenu":
        c.fontSelect.focus();
        return;
      case "sizeMenu":
        c.sizeSelect.focus();
        return;
      default:
        exec(c, command);
    }
  }

  function onComposerKey(c, event) {
    const accel = event.ctrlKey || event.metaKey;
    const key = event.key.toLowerCase();
    if (mentionKey(event) || suggestKey(event)) {
      return;
    }
    if (c.editor.contains(event.target) && moveCaret(c, event)) {
      event.stopPropagation();
      return;
    }
    if (event.key === "Tab" && !accel && !event.altKey && c.editor.contains(event.target)) {
      event.preventDefault();
      event.stopPropagation();
      tabInText(c, event.shiftKey);
      return;
    }
    if (accel && key === "enter") {
      event.preventDefault();
      send(c);
    } else if (DOCK && accel && key === "n" && !event.altKey) {
      event.preventDefault();
      openDockMessage(bridge.dock.create());
    } else if (accel && event.shiftKey && !event.altKey && key === "v" && c.editor.contains(event.target)) {
      // Paste as text only, in the font where it goes (Outlook's "Keep Text
      // Only"). The paste handler takes the text; it is inserted after the
      // paste command, as commands cannot run inside one another.
      event.preventDefault();
      const paste = (c.plainPaste = { text: "" });
      try {
        c.editor.ownerDocument.execCommand("paste");
      } finally {
        c.plainPaste = null;
      }
      if (paste.text) {
        c.editor.ownerDocument.execCommand("insertText", false, paste.text);
      }
    } else if (key === "escape" && formatMenu) {
      event.preventDefault();
      closeFormatMenu();
    } else if (accel && !event.altKey && c.editor.contains(event.target)) {
      const command = FORMAT_KEYS[`${event.shiftKey ? "shift+" : ""}${key}`];
      if (command) {
        event.preventDefault();
        saveSelection(c);
        runTool(c, command, null);
      }
    }
    if (accel && !EDIT_KEYS.has(key)) {
      event.preventDefault();
    }
    // Typing here must not reach Conversations' or Thunderbird's shortcuts.
    event.stopPropagation();
  }

  /* ---------------------------------------------------------------- */
  /* Tab, and lists started by typing, as in Outlook                   */
  /* ---------------------------------------------------------------- */

  // Without this, Tab left the text for the next control of the page.
  // - In a bulleted or numbered list: the item goes one level in (a list
  //   inside the list); Shift+Tab one level out, and out of the list at the
  //   first level.
  // - Elsewhere, with the caret at the start of a paragraph or several
  //   paragraphs selected: the paragraphs are indented (Shift+Tab: back).
  // - Otherwise a tab is typed: four spaces that do not collapse, which every
  //   mail program shows the same way.
  const TAB_TEXT = "\u00a0\u00a0\u00a0\u00a0";
  const BLOCKS = "div, p, li, h1, h2, h3, h4, h5, h6, blockquote, pre, td, th";

  function blockOf(c, node) {
    const element = node && (node.nodeType === 1 ? node : node.parentElement);
    const block = element && element.closest(BLOCKS);
    return block && block !== c.editor && c.editor.contains(block) ? block : c.editor;
  }

  // The text of the caret's paragraph before the caret.
  function textBeforeCaret(c, range) {
    const before = doc.createRange();
    before.setStart(blockOf(c, range.startContainer), 0);
    before.setEnd(range.startContainer, range.startOffset);
    return before;
  }

  function tabInText(c, back) {
    const selection = window.getSelection();
    if (!selection.rangeCount || !c.editor.contains(selection.getRangeAt(0).commonAncestorContainer)) {
      return;
    }
    const range = selection.getRangeAt(0);
    saveSelection(c);
    const start = blockOf(c, range.startContainer);
    const end = blockOf(c, range.endContainer);
    const item = start.closest("li");
    if (item && c.editor.contains(item)) {
      exec(c, back ? "outdent" : "indent");
    } else if (back) {
      exec(c, "outdent", null, true);
    } else if (start !== end || !textBeforeCaret(c, range).toString().replace(/[\s\u00a0]/g, "")) {
      exec(c, "indent", null, true);
    } else {
      restoreSelection(c);
      doc.execCommand("insertText", false, TAB_TEXT);
      markDirty(c);
    }
  }

  // "- ", "* " or "• " typed at the start of a paragraph starts a bulleted
  // list, "1. " or "1) " a numbered one (Outlook's AutoFormat); Ctrl+Z
  // takes it back.
  function autoList(c, event) {
    if (event.inputType !== "insertText" || event.data !== " ") {
      return;
    }
    const selection = window.getSelection();
    if (!selection.rangeCount || !selection.isCollapsed) {
      return;
    }
    const range = selection.getRangeAt(0);
    if (!c.editor.contains(range.startContainer) || blockOf(c, range.startContainer).closest("li")) {
      return;
    }
    const before = textBeforeCaret(c, range);
    const match = /^[\s\u00a0]*([-*\u2022]|1[.)])[\s\u00a0]$/.exec(before.toString());
    if (!match) {
      return;
    }
    selection.removeAllRanges();
    selection.addRange(before);
    doc.execCommand("delete", false, null);
    doc.execCommand(match[1].startsWith("1") ? "insertOrderedList" : "insertUnorderedList", false, null);
    saveSelection(c);
    markDirty(c);
    updateToolState(c);
  }

  // Arrow keys, Home / End and Page Up / Down in the text. The pane is a
  // page inside Thunderbird's own window: there, Gecko moves the caret for
  // these keys as for page navigation and then gives the focus to another
  // element of the page (the dock's tab button), so the next key typed
  // lands outside the text. The caret is moved here instead
  // (Selection.modify), with Shift extending the selection and Ctrl going by
  // word / to the start or end of the text, and kept in view.
  const CARET_KEYS = {
    ArrowLeft: ["left", "character", "word"],
    ArrowRight: ["right", "character", "word"],
    ArrowUp: ["backward", "line", "line"],
    ArrowDown: ["forward", "line", "line"],
    Home: ["backward", "lineboundary", "text"],
    End: ["forward", "lineboundary", "text"],
    PageUp: ["backward", "page", "page"],
    PageDown: ["forward", "page", "page"],
  };

  function moveCaret(c, event) {
    const how = CARET_KEYS[event.key];
    if (!how || event.altKey || event.isComposing) {
      return false;
    }
    const selection = window.getSelection();
    if (!selection.rangeCount || !c.editor.contains(selection.focusNode) || typeof selection.modify !== "function") {
      return false;
    }
    const accel = event.ctrlKey || event.metaKey;
    const alter = event.shiftKey ? "extend" : "move";
    const [direction, granularity, accelGranularity] = how;
    const unit = accel ? accelGranularity : granularity;
    event.preventDefault();
    if (unit === "text") {
      // Ctrl+Home / Ctrl+End: start or end of the whole text, inside its
      // first or last line (not between the lines' blocks).
      const [node, offset] = textEdge(c.editor, direction === "forward");
      if (alter === "extend") {
        selection.extend(node, offset);
      } else {
        selection.collapse(node, offset);
      }
    } else if (unit === "page") {
      const lineHeight = parseFloat(getComputedStyle(c.editor).lineHeight) || 18;
      const lines = Math.max(1, Math.floor((window.innerHeight * 0.8) / lineHeight) - 1);
      for (let i = 0; i < lines; i++) {
        selection.modify(alter, direction, "line");
      }
    } else {
      selection.modify(alter, direction, unit);
    }
    if (doc.activeElement !== c.editor) {
      c.editor.focus({ preventScroll: true });
    }
    revealCaret(c);
    return true;
  }

  // [node, offset] of the very start or end of the text, inside its first
  // or last line.
  function textEdge(root, end) {
    const LEAVES = new Set(["BR", "IMG", "HR", "INPUT"]);
    let node = root;
    for (;;) {
      const child = end ? node.lastChild : node.firstChild;
      if (!child) {
        return [node, end ? node.childNodes.length : 0];
      }
      if (child.nodeType === Node.TEXT_NODE) {
        return [child, end ? child.length : 0];
      }
      if (child.nodeType !== Node.ELEMENT_NODE || LEAVES.has(child.nodeName)) {
        // Before a final <br> (the line's own end), or at the node's edge.
        const index = Array.prototype.indexOf.call(node.childNodes, child);
        return [node, end ? (child.nodeName === "BR" ? index : index + 1) : 0];
      }
      node = child;
    }
  }

  // Scrolls the text's scrolling box so the caret is in view.
  function revealCaret(c) {
    const selection = window.getSelection();
    if (!selection.rangeCount) {
      return;
    }
    const range = selection.getRangeAt(0).cloneRange();
    range.collapse(selection.focusNode === range.startContainer && selection.focusOffset === range.startOffset);
    let rect = range.getBoundingClientRect();
    if (!rect.height) {
      const node = selection.focusNode;
      const element = node && (node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement);
      rect = element ? element.getBoundingClientRect() : rect;
    }
    for (let box = c.editor; box; box = box.parentElement) {
      const style = getComputedStyle(box);
      if (/(auto|scroll)/.test(style.overflowY) && box.scrollHeight > box.clientHeight) {
        const view = box.getBoundingClientRect();
        if (rect.top < view.top) {
          box.scrollTop -= view.top - rect.top + 4;
        } else if (rect.bottom > view.bottom) {
          box.scrollTop += rect.bottom - view.bottom + 4;
        }
        return;
      }
    }
  }

  function focusEditor(c) {
    c.editor.focus();
    const range = doc.createRange();
    range.setStart(c.editor.firstChild || c.editor, 0);
    range.collapse(true);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
  }

  function swapSignature(c) {
    const identity = c.data.identities.find((i) => i.key === c.from.value);
    let signature = c.editor.querySelector(".moz-signature");
    if (!identity || !identity.signature) {
      if (signature) {
        signature.remove();
      }
      return;
    }
    if (!signature) {
      signature = el("div", { class: "moz-signature" });
      c.editor.appendChild(signature);
    }
    signature.innerHTML = bridge.sanitize(identity.signature); // sanitized when stored, and again here
    liveImages(signature);
  }

  function markDirty(c) {
    if (c !== composer) {
      return;
    }
    c.dirty = true;
    setStatus(c, "Draft");
    clearTimeout(c.saveTimer);
    c.saveTimer = setTimeout(() => saveDraft(c), 800);
  }

  function snapshot(c) {
    return {
      type: c.type,
      identityKey: c.from.value,
      to: c.to.value,
      cc: c.cc.value,
      bcc: c.bcc.value,
      subject: c.subject.value,
      html: c.editor.innerHTML,
      attachments: c.attachments,
      files: c.files,
      dirty: c.dirty,
    };
  }

  function saveDraft(c) {
    clearTimeout(c.saveTimer);
    if (c.dirty && !c.sending) {
      bridge.saveDraft(c.uri, snapshot(c));
    }
  }

  function setStatus(c, text, isError = false) {
    c.status.textContent = text;
    c.status.classList.toggle("ol-error", isError);
  }

  function setSending(c, on) {
    c.sending = on;
    c.editor.contentEditable = on ? "false" : "true";
    for (const node of c.root.querySelectorAll("input, select, .ol-actions button, .ol-tool, .ol-link, .ol-remove")) {
      node.disabled = on;
    }
  }

  function closeComposer({ keepDraft }) {
    const c = composer;
    if (!c) {
      return;
    }
    if (keepDraft) {
      saveDraft(c);
    }
    clearTimeout(c.saveTimer);
    clearInterval(c.countdownTimer);
    hideSuggest();
    hideMention();
    closeFormatMenu();
    c.root.remove();
    composer = null;
  }

  /* ---------------------------------------------------------------- */
  /* Attachments                                                       */
  /* ---------------------------------------------------------------- */

  function addFiles(c, fileList) {
    for (const file of fileList) {
      // The file's path stays in the experiment; drafts keep only a token.
      const id = bridge.addFile(file);
      if (id && !c.files.some((f) => f.id === id)) {
        c.files.push({ id, name: file.name, size: file.size, type: file.type || "" });
      }
    }
    renderAttachments(c);
    markDirty(c);
  }

  function renderAttachments(c) {
    c.attachmentList.replaceChildren();
    const chip = (name, size, onRemove) =>
      el(
        "span",
        { class: "ol-chip" },
        `📎 ${name}${size ? ` (${formatSize(size)})` : ""}`,
        el("button", { class: "ol-remove", text: "×", title: `Remove ${name}`, onclick: onRemove })
      );
    c.attachments.forEach((a, index) =>
      c.attachmentList.appendChild(
        chip(a.name, a.size, () => {
          c.attachments = c.attachments.filter((_, i) => i !== index);
          renderAttachments(c);
          markDirty(c);
        })
      )
    );
    c.files.forEach((f, index) =>
      c.attachmentList.appendChild(
        chip(f.name, f.size, () => {
          c.files = c.files.filter((_, i) => i !== index);
          renderAttachments(c);
          markDirty(c);
        })
      )
    );
    c.attachmentList.classList.toggle("ol-hidden", !c.attachments.length && !c.files.length);
  }

  /* ---------------------------------------------------------------- */
  /* Address suggestions                                               */
  /* ---------------------------------------------------------------- */

  let suggestState = null; // { input, box, items, active }

  // A textarea that wraps and grows with the recipients; past its CSS
  // max-height it scrolls, so a long To or Cc list stays readable.
  function addressInput(c, value) {
    const input = el("textarea", {
      class: "ol-input ol-address",
      rows: "1",
      spellcheck: "false",
      autocomplete: "off",
    });
    input.value = (value || "").replace(/\s*[\r\n]+\s*/g, " ");
    input.addEventListener("input", () => {
      // Pasted lists arrive one address per line; keep them on one logical line.
      if (/[\r\n]/.test(input.value)) {
        const caret = input.selectionStart;
        const before = input.value.slice(0, caret).replace(/\s*[\r\n]+\s*/g, ", ");
        input.value = before + input.value.slice(caret).replace(/\s*[\r\n]+\s*/g, ", ");
        input.setSelectionRange(before.length, before.length);
      }
      fitAddress(input);
      markDirty(c);
      suggest(input);
    });
    input.addEventListener("keydown", (event) => {
      // No line breaks in an address list; the composer still handles
      // Enter for suggestions and Ctrl+Enter for Send.
      if (event.key === "Enter") {
        event.preventDefault();
      }
    });
    input.addEventListener("focus", () => fitAddress(input));
    input.addEventListener("blur", () => setTimeout(() => suggestState && suggestState.input === input && hideSuggest(), 150));
    return input;
  }

  function fitAddress(input) {
    if (input.localName !== "textarea" || !input.isConnected) {
      return;
    }
    input.style.height = "auto";
    if (input.scrollHeight) {
      const style = window.getComputedStyle(input);
      const borders = parseFloat(style.borderTopWidth) + parseFloat(style.borderBottomWidth) || 0;
      input.style.height = `${input.scrollHeight + borders}px`;
    }
  }

  function fitAddresses(c) {
    for (const input of [c.to, c.cc, c.bcc]) {
      fitAddress(input);
    }
  }

  function currentToken(input) {
    const before = input.value.slice(0, input.selectionStart);
    const start = Math.max(before.lastIndexOf(","), before.lastIndexOf(";")) + 1;
    return { start, text: before.slice(start).trim() };
  }

  function suggest(input) {
    const { text } = currentToken(input);
    const items = text.length >= 1 ? bridge.searchAddresses(text) : [];
    if (!items.length) {
      hideSuggest();
      return;
    }
    if (!suggestState || suggestState.input !== input) {
      hideSuggest();
      const box = el("div", { class: "ol-suggest", role: "listbox" });
      composer.root.appendChild(box);
      suggestState = { input, box, items: [], active: 0 };
    }
    suggestState.items = items;
    suggestState.active = 0;
    renderSuggest();
    const rootRect = composer.root.getBoundingClientRect();
    const rect = input.getBoundingClientRect();
    suggestState.box.style.left = `${rect.left - rootRect.left}px`;
    suggestState.box.style.top = `${rect.bottom - rootRect.top + 2}px`;
    suggestState.box.style.width = `${rect.width}px`;
  }

  function renderSuggest() {
    const { box, items, active } = suggestState;
    box.replaceChildren(
      ...items.map((item, index) => {
        const option = el("div", {
          class: `ol-suggest-item${index === active ? " ol-active" : ""}`,
          role: "option",
          text: item,
        });
        option.addEventListener("mousedown", (event) => {
          event.preventDefault();
          chooseSuggestion(index);
        });
        return option;
      })
    );
  }

  function chooseSuggestion(index) {
    const { input, items } = suggestState;
    const { start } = currentToken(input);
    const head = input.value.slice(0, start).replace(/\s*$/, "");
    const tail = input.value.slice(input.selectionStart).replace(/^[^,;]*[,;]?\s*/, "");
    const value = `${head ? `${head} ` : ""}${items[index]}, `;
    input.value = value + tail;
    input.setSelectionRange(value.length, value.length);
    hideSuggest();
    if (composer) {
      markDirty(composer);
    }
  }

  // Returns true when the key was used by the suggestion list.
  function suggestKey(event) {
    if (!suggestState || event.target !== suggestState.input) {
      return false;
    }
    const count = suggestState.items.length;
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      suggestState.active = (suggestState.active + (event.key === "ArrowDown" ? 1 : count - 1)) % count;
      renderSuggest();
    } else if (event.key === "Enter" || event.key === "Tab") {
      chooseSuggestion(suggestState.active);
    } else if (event.key === "Escape") {
      hideSuggest();
    } else {
      return false;
    }
    event.preventDefault();
    event.stopPropagation();
    return true;
  }

  function hideSuggest() {
    if (suggestState) {
      suggestState.box.remove();
      suggestState = null;
    }
  }

  /* ---------------------------------------------------------------- */
  /* @mentions                                                         */
  /* ---------------------------------------------------------------- */

  // As in the compose window (compose/compose.js): "@name" offers contacts
  // from the address books; the chosen one becomes a mailto link "@Name" in
  // the body and is added to To.
  // "@" at the start of a line or after a space or bracket, then the query.
  const MENTION_TRIGGER = /(^|[\s(\[{>"'])@([^\s@]{1,40})$/;
  let mentionState = null; // { c, box, items, active }

  function mentionContext(c) {
    const selection = window.getSelection();
    if (!selection || !selection.rangeCount || !selection.isCollapsed) {
      return null;
    }
    const node = selection.anchorNode;
    const offset = selection.anchorOffset;
    if (!node || node.nodeType !== Node.TEXT_NODE || !c.editor.contains(node) || node.parentElement.closest("a")) {
      return null;
    }
    const match = node.data.slice(0, offset).match(MENTION_TRIGGER);
    return match ? { node, start: offset - match[2].length - 1, end: offset, query: match[2] } : null;
  }

  function mentionRange(context) {
    const range = doc.createRange();
    range.setStart(context.node, context.start);
    range.setEnd(context.node, context.end);
    return range;
  }

  function updateMention(c) {
    const context = c === composer && !c.sending ? mentionContext(c) : null;
    const items = context ? bridge.searchContacts(context.query) : [];
    if (!items.length) {
      hideMention();
      return;
    }
    if (!mentionState || mentionState.c !== c) {
      hideMention();
      const box = el("div", { class: "ol-suggest ol-mention", role: "listbox" });
      // mousedown + preventDefault keeps the caret in the editor.
      box.addEventListener("mousedown", (event) => {
        event.preventDefault();
        const row = event.target.closest("[data-index]");
        if (row) {
          chooseMention(Number(row.dataset.index));
        }
      });
      c.root.appendChild(box);
      mentionState = { c, box, items: [], active: 0 };
    }
    mentionState.items = items;
    mentionState.active = 0;
    renderMention();
    // Under the "@name", or above it when there is no room below.
    const range = mentionRange(context);
    const rects = range.getClientRects ? range.getClientRects() : [];
    const rect = rects.length ? rects[rects.length - 1] : context.node.parentElement.getBoundingClientRect();
    const rootRect = c.root.getBoundingClientRect();
    const box = mentionState.box;
    let top = rect.bottom - rootRect.top + 4;
    if (rect.bottom + box.offsetHeight + 4 > window.innerHeight && rect.top - box.offsetHeight - 4 > 0) {
      top = rect.top - rootRect.top - box.offsetHeight - 4;
    }
    box.style.top = `${top}px`;
    box.style.left = `${Math.max(0, Math.min(rect.left - rootRect.left, rootRect.width - box.offsetWidth))}px`;
  }

  function renderMention() {
    const { box, items, active } = mentionState;
    box.replaceChildren(
      ...items.map((item, index) =>
        el(
          "div",
          {
            class: `ol-suggest-item ol-mention-item${index === active ? " ol-active" : ""}`,
            role: "option",
            "data-index": String(index),
          },
          el("b", { text: item.name }),
          el("span", { class: "ol-mention-email", text: item.email })
        )
      )
    );
  }

  const MENTION_EMAIL = /^[^\s@<>(),;:"'\\]+@[^\s@<>(),;:"'\\]+\.[^\s@<>(),;:"'\\]+$/;

  function chooseMention(index) {
    const { c, items } = mentionState;
    const item = items[index];
    const context = mentionContext(c);
    hideMention();
    if (!item || !context) {
      return;
    }
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(mentionRange(context));
    // Only a plain address goes into the link; el() sets the address and the
    // name as an attribute and as text, so outerHTML escapes both.
    if (!MENTION_EMAIL.test(String(item.email || ""))) {
      return;
    }
    const link = el("a", { href: `mailto:${item.email}`, text: `@${String(item.label || item.email)}` });
    doc.execCommand("insertHTML", false, `${link.outerHTML}&nbsp;`);
    addMentionRecipient(c, item);
    markDirty(c);
  }

  // Adds the mentioned person to To unless To, Cc or Bcc has them already.
  function addMentionRecipient(c, item) {
    const email = item.email.toLowerCase().replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const present = new RegExp(`(^|[\\s<,;"])${email}($|[\\s>,;"])`, "i");
    if ([c.to, c.cc, c.bcc].some((input) => present.test(input.value))) {
      return;
    }
    const current = c.to.value.replace(/[\s,;]+$/, "");
    c.to.value = `${current ? `${current}, ` : ""}${item.address}, `;
    fitAddress(c.to);
  }

  // Returns true when the key was used by the mention list.
  function mentionKey(event) {
    if (!mentionState || !mentionState.c.editor.contains(event.target)) {
      return false;
    }
    const count = mentionState.items.length;
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      mentionState.active = (mentionState.active + (event.key === "ArrowDown" ? 1 : count - 1)) % count;
      renderMention();
    } else if ((event.key === "Enter" || event.key === "Tab") && !event.ctrlKey && !event.metaKey) {
      chooseMention(mentionState.active);
    } else if (event.key === "Escape") {
      hideMention();
    } else {
      return false;
    }
    event.preventDefault();
    event.stopPropagation();
    return true;
  }

  function hideMention() {
    if (mentionState) {
      mentionState.box.remove();
      mentionState = null;
    }
  }

  /* ---------------------------------------------------------------- */
  /* Send, Undo Send, discard, pop out                                 */
  /* ---------------------------------------------------------------- */

  async function send(c) {
    if (c !== composer || c.sending) {
      return;
    }
    const draft = snapshot(c);
    if (!draft.to.trim() && !draft.cc.trim() && !draft.bcc.trim()) {
      setStatus(c, "Add at least one recipient.", true);
      c.to.focus();
      return;
    }
    if (c.type === "new" && !draft.subject.trim() && !window.confirm("Send this message without a subject?")) {
      c.subject.focus();
      return;
    }
    clearTimeout(c.saveTimer);
    setSending(c, true);
    setStatus(c, "Sending…");
    try {
      const result = await bridge.send(c.uri, draft);
      if (c !== composer) {
        return;
      }
      if (result.pending) {
        showCountdown(c, result.until);
      } else {
        finishSent(c);
      }
    } catch (error) {
      if (c === composer) {
        clearCountdown(c);
        setSending(c, false);
        setStatus(c, errorText(error), true);
      }
    }
  }

  function showCountdown(c, until) {
    const tick = () => {
      const left = Math.ceil((until - Date.now()) / 1000);
      c.countdownText.textContent = left > 0 ? `Sending in ${left}s…` : "Sending…";
    };
    tick();
    clearInterval(c.countdownTimer);
    c.countdownTimer = setInterval(tick, 500);
    c.actions.classList.add("ol-hidden");
    c.countdown.classList.remove("ol-hidden");
    for (const button of c.countdown.querySelectorAll("button")) {
      button.disabled = false;
    }
    if (DOCK) {
      renderDockBar(); // the tab shows it is sending
    }
  }

  function clearCountdown(c) {
    clearInterval(c.countdownTimer);
    c.countdown.classList.add("ol-hidden");
    c.actions.classList.remove("ol-hidden");
    if (DOCK) {
      renderDockBar();
    }
  }

  function undoSend(c) {
    bridge.cancelSend(c.uri);
    clearCountdown(c);
    setSending(c, false);
    setStatus(c, "Not sent. You can keep editing.");
    focusEditor(c);
  }

  function sendNow(c) {
    for (const button of c.countdown.querySelectorAll("button")) {
      button.disabled = true;
    }
    c.countdownText.textContent = "Sending…";
    bridge.sendNow(c.uri).catch(() => {}); // the result comes through onSendResult
  }

  function finishSent(c) {
    if (c !== composer) {
      return;
    }
    composer.dirty = false;
    closeComposer({ keepDraft: false });
    if (DOCK) {
      dockAfterClose(c.uri, "Message sent.");
    } else {
      toast("Sent.");
    }
  }

  // From the experiment, when a send finishes. Returns true when shown here.
  function onSendResult(uri, result) {
    const c = composer;
    if (!c || c.uri !== uri) {
      return DOCK ? dockSendResult(uri, result) : false;
    }
    if (result.ok) {
      finishSent(c);
    } else {
      clearCountdown(c);
      setSending(c, false);
      setStatus(c, `Not sent: ${result.error}`, true);
    }
    return true;
  }

  function discard(c) {
    if (c.sending || (c.dirty && !window.confirm("Discard this draft?"))) {
      return;
    }
    bridge.discardDraft(c.uri);
    closeComposer({ keepDraft: false });
    if (DOCK) {
      dockAfterClose(c.uri);
    }
  }

  function popOut(c) {
    if (
      c.files.length &&
      !window.confirm("Files attached here have to be attached again in the compose window. Continue?")
    ) {
      return;
    }
    const draft = snapshot(c);
    const body = c.editor.cloneNode(true);
    for (const signature of body.querySelectorAll(".moz-signature")) {
      signature.remove(); // the compose window adds its own
    }
    deadImages(body);
    draft.html = body.innerHTML;
    if (!bridge.popOut(c.uri, draft)) {
      setStatus(c, "Could not open the compose window.", true);
      return;
    }
    closeComposer({ keepDraft: false });
    if (DOCK) {
      dockAfterClose(c.uri);
    }
  }

  /* ---------------------------------------------------------------- */
  /* Compose dock (new messages)                                       */
  /* ---------------------------------------------------------------- */

  // dock.html: a grip to resize, a bar with one tab per new message, and the
  // composer of the active tab. The experiment sizes the dock.
  const dockGrip = DOCK ? doc.getElementById("ol-dock-grip") : null;
  const dockBar = DOCK ? doc.getElementById("ol-dock-bar") : null;
  const dockBody = DOCK ? doc.getElementById("ol-dock-body") : null;
  let dockNoticeTimer = 0;

  function renderDockBar(notice = "") {
    const state = bridge.dock.state();
    const tabs = bridge.dock.messages().map((m) => {
      const active = Boolean(composer && composer.uri === m.uri);
      const title = (active ? composer.subject.value : m.subject).trim() || "New message";
      const tab = el(
        "div",
        { class: `ol-dock-tab${active ? " ol-active" : ""}`, role: "tab", title, "aria-selected": String(active) },
        el("span", { class: "ol-dock-title", text: `${m.sending ? "⏳ " : ""}${title}` }),
        el("button", {
          class: "ol-dock-close",
          text: "×",
          title: "Close",
          onclick: (event) => {
            event.stopPropagation();
            closeDockMessage(m.uri);
          },
        })
      );
      tab.addEventListener("click", () => openDockMessage(m.uri));
      return tab;
    });
    const button = (text, title, onclick) => el("button", { class: "ol-dock-button", text, title, onclick });
    dockBar.replaceChildren(
      el(
        "div",
        { class: "ol-dock-tabs", role: "tablist" },
        ...tabs,
        notice ? el("span", { class: "ol-dock-notice", text: notice }) : null
      ),
      button("+", "New message (Ctrl+N)", () => openDockMessage(bridge.dock.create())),
      button(state.minimized ? "▴" : "▾", state.minimized ? "Restore" : "Minimize", () => {
        bridge.dock.setMinimized(!state.minimized);
        renderDockBar();
      }),
      button(state.maximized ? "❐" : "□", state.maximized ? "Restore size" : "Maximize", () => {
        bridge.dock.setMaximized(!state.maximized);
        renderDockBar();
      })
    );
  }

  function openDockMessage(uri) {
    clearTimeout(dockNoticeTimer);
    if (bridge.dock.state().minimized) {
      bridge.dock.setMinimized(false);
    }
    openComposer(uri, "new", 0, { focus: true });
    renderDockBar();
  }

  // The experiment added a message, or the page just started.
  function dockChanged() {
    const focusUri = bridge.dock.takeFocus();
    const last = bridge.dock.messages().at(-1);
    if (focusUri) {
      openDockMessage(focusUri);
    } else if (!composer && last) {
      openDockMessage(last.uri);
    } else {
      renderDockBar();
    }
  }

  // After a message was sent, discarded or popped out: show the next one, or
  // tuck the dock away (after a short notice).
  function dockAfterClose(uri, notice = "") {
    bridge.dock.remove(uri);
    const last = bridge.dock.messages().at(-1);
    if (last) {
      openDockMessage(last.uri);
      if (notice) {
        toast(notice);
      }
      return;
    }
    if (!notice) {
      renderDockBar();
      bridge.dock.hide();
      return;
    }
    bridge.dock.setMinimized(true);
    renderDockBar(`✓ ${notice}`);
    clearTimeout(dockNoticeTimer);
    dockNoticeTimer = setTimeout(() => {
      if (!bridge.dock.messages().length) {
        bridge.dock.hide();
      }
    }, 4000);
  }

  // A message sent from a tab that is not shown (Undo Send ran out meanwhile).
  function dockSendResult(uri, result) {
    if (!uri.startsWith("new:")) {
      return false; // an inline reply, shown by its conversation page
    }
    if (result.ok) {
      if (!composer && !bridge.dock.messages().length) {
        dockAfterClose(uri, "Message sent.");
      } else {
        renderDockBar();
        toast("Sent.");
      }
    } else {
      renderDockBar();
      toast(`Not sent (${result.subject || "no subject"}): ${result.error}`, true);
    }
    return true;
  }

  function closeDockMessage(uri) {
    if (composer && composer.uri === uri) {
      discard(composer);
      return;
    }
    if (bridge.pendingSend(uri)) {
      return; // sending
    }
    const draft = bridge.getDraft(uri);
    if (draft && draft.dirty && !window.confirm("Discard this draft?")) {
      return;
    }
    bridge.discardDraft(uri);
    bridge.dock.remove(uri);
    renderDockBar();
    if (!bridge.dock.messages().length) {
      bridge.dock.hide();
    }
  }

  if (DOCK) {
    dockGrip.addEventListener("pointerdown", (event) => {
      if (event.button !== 0) {
        return;
      }
      event.preventDefault();
      dockGrip.setPointerCapture(event.pointerId);
      const startY = event.screenY;
      const startHeight = bridge.dock.height();
      const move = (e) => bridge.dock.setHeight(startHeight - (e.screenY - startY));
      dockGrip.addEventListener("pointermove", move);
      dockGrip.addEventListener(
        "pointerup",
        () => {
          dockGrip.removeEventListener("pointermove", move);
          renderDockBar();
        },
        { once: true }
      );
    });
    dockBar.addEventListener("dblclick", (event) => {
      if (event.target === dockBar || event.target.classList.contains("ol-dock-tabs")) {
        bridge.dock.setMinimized(!bridge.dock.state().minimized);
        renderDockBar();
      }
    });
  }

  /* ---------------------------------------------------------------- */
  /* Following the conversation                                        */
  /* ---------------------------------------------------------------- */

  // Conversations reuses this page for every conversation: close the
  // composer (its draft is kept) when another conversation is shown, and
  // bring back a draft or a countdown when its conversation comes back.
  function syncConversation() {
    applyOrder();
    const ids = conversationIds();
    if (!DOCK && ids.length && !ids.some((id) => shownIds.includes(id))) {
      if (config.newestFirst) {
        showNewest(); // another conversation
      }
    }
    if (ids.length) {
      shownIds = ids;
    }
    expandMessages(ids);
    updateEarlier(ids);
    updatePinButtons();
    updateSaveButton();
    updateHeaders();
    updateMessageFrames();
    updatePictureFit();
    updateQuoteFolds();
    if (!ids.length || !config.inlineReply) {
      return;
    }
    if (composer) {
      if (ids.includes(composer.convId)) {
        return;
      }
      closeComposer({ keepDraft: true });
    }
    for (const id of [...ids].reverse()) {
      const uri = bridge.uriForConvId(id);
      if (!uri) {
        continue;
      }
      const pending = bridge.pendingSend(uri);
      const draft = pending ? pending.draft : bridge.getDraft(uri);
      if (draft && (pending || draft.dirty)) {
        openComposer(uri, draft.type, id, { focus: false });
        return;
      }
    }
  }

  function scheduleSync() {
    clearTimeout(syncTimer);
    syncTimer = setTimeout(() => {
      try {
        syncConversation();
      } catch (error) {
        bridge.reportError("sync", errorText(error));
      }
    }, 150);
  }

  const observer = new MutationObserver(scheduleSync);
  const list = doc.getElementById("messageList");
  if (list) {
    observer.observe(list, { childList: true, subtree: true });
  }

  /* ---------------------------------------------------------------- */
  /* Start and stop                                                    */
  /* ---------------------------------------------------------------- */

  function destroy() {
    observer.disconnect();
    frameResize.disconnect();
    clearTimeout(expandTimer);
    unfoldAll();
    clearTimeout(syncTimer);
    clearTimeout(toastTimer);
    clearTimeout(dockNoticeTimer);
    closeComposer({ keepDraft: true });
    for (const cleanup of cleanups) {
      cleanup();
    }
    stylesheet.remove();
    doc.getElementById("ol-toast")?.remove();
    for (const url of blobUrls.values()) {
      URL.revokeObjectURL(url);
    }
    blobUrls.clear();
    earlierButton?.remove();
    saveToolbar()?.querySelector(".ol-save")?.remove();
    clearOrder();
    for (const li of doc.querySelectorAll("li.message[data-ol-load]")) {
      li.removeAttribute("data-ol-load");
    }
    for (const box of doc.querySelectorAll(".ol-recipients")) {
      box.remove();
    }
    doc.documentElement.removeAttribute("data-ol-light");
    doc.documentElement.removeAttribute("data-ol-dark");
    doc.documentElement.removeAttribute("data-ol-outlook-header");
    doc.documentElement.removeAttribute("data-ol-attachments-top");
    doc.documentElement.removeAttribute("data-ol-newest-first");
    doc.documentElement.removeAttribute("data-ol-latest-only");
    doc.documentElement.removeAttribute("data-ol-show-earlier");
    doc.documentElement.removeAttribute("data-ol-inline-reply");
  }

  // The toolbar follows the caret; menus and lists act on the kept selection.
  listen(doc, "selectionchange", () => {
    const c = composer;
    if (c && c.editor && doc.activeElement === c.editor) {
      saveSelection(c);
      updateToolState(c);
    }
  });

  if (!DOCK) {
    listen(window, "click", onClick, true);
    listen(window, "keydown", onKeyDown, true);
  }
  // A wider or narrower pane rewraps the recipients.
  listen(window, "resize", () => composer && fitAddresses(composer));

  window.outlookModePane = {
    openFor,
    onSendResult,
    dockChanged,
    configChanged() {
      config = bridge.config();
      applyConfig();
    },
    destroy,
  };

  applyConfig();
  scheduleSync();
})();
