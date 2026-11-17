"use strict";

// Outlook-style reading pane (experiments/outlookPane + reading-pane/): in the
// Thunderbird Conversations view, newest mail on top, and Reply / Reply all /
// Forward written inline in the pane instead of a compose window; new
// messages in a dock at the bottom of the main window. This file keeps the
// experiment configured, opens the compose window when the user pops a
// message out, and reports failed sends that no pane showed.

let readingPaneStatus = { error: "" };

async function configureReadingPane() {
  const settings = await getSettings();
  try {
    const status = await messenger.outlookPane.configure({
      enabled: Boolean(
        settings.readingPaneNewestFirst ||
          settings.readingPaneExpandAll ||
          settings.readingPaneFoldQuotes ||
          settings.readingPaneLatestOnly ||
          settings.inlineReply ||
          settings.inlineNewMessage
      ),
      newestFirst: Boolean(settings.readingPaneNewestFirst),
      expandAll: Boolean(settings.readingPaneExpandAll),
      foldQuotes: Boolean(settings.readingPaneFoldQuotes),
      attachmentsTop: settings.readingPaneAttachmentsTop !== false,
      latestOnly: Boolean(settings.readingPaneLatestOnly),
      inlineReply: Boolean(settings.inlineReply),
      inlineNew: Boolean(settings.inlineNewMessage),
      pinButton: Boolean(settings.pinButton),
      outlookHeader: settings.readingPaneOutlookHeader !== false,
      lightPane: settings.readingPaneLight !== false,
      pinKeyword: PIN_TAG_KEY,
      undoSendSeconds: Number(settings.undoSendSeconds) || 0,
      replyHeader: settings.replyHeader !== false,
      removeQuoteBar: settings.removeQuoteBar !== false,
    });
    readingPaneStatus = { ...status, error: "" };
  } catch (error) {
    readingPaneStatus = { error: String(error && error.message ? error.message : error) };
    logError("reading pane", error);
  }
}

// Put the text written in the pane above the quote the compose window made.
function prependToBody(body, html) {
  const match = body.match(/<body[^>]*>/i);
  if (!match) {
    return html + body;
  }
  const at = match.index + match[0].length;
  return body.slice(0, at) + html + body.slice(at);
}

function htmlToText(html) {
  const doc = new DOMParser().parseFromString(html, "text/html");
  for (const br of doc.querySelectorAll("br")) {
    br.replaceWith("\n");
  }
  for (const block of doc.querySelectorAll("div, p, li")) {
    block.append("\n");
  }
  return doc.body.textContent.replace(/\n{3,}/g, "\n\n").trim();
}

const POP_OUT_TYPES = ["new", "reply", "replyAll", "forward"];

// The fields the user filled in the pane, checked: a known type, one of the
// user's own identities, well-formed addresses only, a subject of normal
// length. The compose window then shows them for the user to check and send.
async function popOutFields(details) {
  if (!details || !POP_OUT_TYPES.includes(details.type)) {
    return null;
  }
  if (details.type !== "new" && !Number.isInteger(details.messageId)) {
    return null;
  }
  const identities = (await messenger.accounts.list(false)).flatMap((account) => account.identities || []);
  const identityId = identities.some((identity) => identity.id === details.identityId) ? details.identityId : undefined;
  const addresses = (list) =>
    (Array.isArray(list) ? list : [])
      .filter((address) => typeof address === "string" && address.length <= 500 && EMAIL_PATTERN.test(emailOf(address)))
      .slice(0, 500);
  return {
    identityId,
    to: addresses(details.to),
    cc: addresses(details.cc),
    bcc: addresses(details.bcc),
    subject: typeof details.subject === "string" ? details.subject.slice(0, 998) : "",
  };
}

// A second check after Thunderbird's sanitizer, here where the HTML goes into
// the mail: active elements, event handler attributes and script/data URLs
// are removed (pictures as data:image stay).
const ACTIVE_ELEMENTS = "script, iframe, frame, frameset, object, embed, applet, form, input, button, textarea, select, meta, base, link, svg, math, template";
const URL_ATTRIBUTES = /^(href|src|action|formaction|xlink:href|background|poster|srcset|ping)$/i;

// CSS with a script URL or old script hooks; CSS escapes are refused too,
// since they can hide those words.
function activeCss(css) {
  const text = css.replace(/[\u0000-\u0020\u007f]+/g, "").toLowerCase();
  return text.includes("\\") || /javascript:|vbscript:|expression\(|-moz-binding|behavior:/.test(text);
}

function inertHtml(html) {
  const doc = new DOMParser().parseFromString(`<!DOCTYPE html><body>${html}</body>`, "text/html");
  doc.querySelectorAll(ACTIVE_ELEMENTS).forEach((element) => element.remove());
  doc.querySelectorAll("style").forEach((element) => activeCss(element.textContent) && element.remove());
  for (const element of doc.body.querySelectorAll("*")) {
    for (const { name, value } of [...element.attributes]) {
      const url = value.replace(/[\u0000-\u0020\u007f]+/g, "").toLowerCase();
      if (
        /^on/i.test(name) ||
        name.toLowerCase() === "srcdoc" ||
        (name.toLowerCase() === "style" && activeCss(value)) ||
        (URL_ATTRIBUTES.test(name) && /^(javascript|vbscript|data):/.test(url) && !(name.toLowerCase() === "src" && url.startsWith("data:image/")))
      ) {
        element.removeAttribute(name);
      }
    }
  }
  return doc.body.innerHTML;
}

async function popOutReply(details) {
  const fields = await popOutFields(details);
  if (!fields) {
    logError("pop out reply", new Error("invalid pop-out request"));
    return;
  }
  const tab =
    details.type === "new"
      ? await messenger.compose.beginNew(fields)
      : details.type === "forward"
      ? await messenger.compose.beginForward(details.messageId, "forwardInline", fields)
      : await messenger.compose.beginReply(
          details.messageId,
          details.type === "replyAll" ? "replyToAll" : "replyToSender",
          fields
        );
  if (!details.html || !htmlToText(details.html)) {
    return;
  }
  const current = await messenger.compose.getComposeDetails(tab.id);
  if (current.isPlainText) {
    await messenger.compose.setComposeDetails(tab.id, {
      plainTextBody: `${htmlToText(details.html)}\n\n${current.plainTextBody || ""}`,
    });
  } else {
    await messenger.compose.setComposeDetails(tab.id, {
      body: prependToBody(current.body || "", inertHtml(await messenger.outlookPane.sanitizeHtml(details.html))),
    });
  }
}

messenger.outlookPane.onPopOut.addListener((details) => {
  popOutReply(details).catch((error) => logError("pop out reply", error));
});

messenger.outlookPane.onOpenSettings.addListener((section) => {
  openSettings(section).catch((error) => logError("open settings", error));
});

// Conversations' star, shown as a pin.
messenger.outlookPane.onPin.addListener(async (messageId) => {
  try {
    await togglePin([await messenger.messages.get(messageId)]);
    const [tab] = await messenger.tabs.query({ active: true, currentWindow: true });
    await resortAfterPin(tab).catch(() => {});
  } catch (error) {
    logError("pin from the reading pane", error);
  }
});

messenger.outlookPane.onSendResult.addListener((result) => {
  if (!result.ok) {
    logError("inline reply", result.error);
    if (!result.shown) {
      notify("Message not sent", `${result.subject || "(no subject)"}: ${result.error}`);
    }
  }
});

messenger.storage.onChanged.addListener((changes) => {
  if (changes.settings) {
    configureReadingPane();
  }
});

handlers["reading-pane-status"] = fromOwnPage(async () => {
  try {
    readingPaneStatus = { ...(await messenger.outlookPane.status()), error: readingPaneStatus.error };
  } catch (error) {
    // keep the last status
  }
  return readingPaneStatus;
});

configureReadingPane();
