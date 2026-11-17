"use strict";

// Preferences → Signature. Two ways to get an Outlook signature, with its
// tables and pictures:
// - the .htm file classic Outlook keeps (experiment outlookPane.importSignature);
// - a mail you sent from Outlook: new Outlook and Outlook on the web keep
//   signatures in the cloud, so a sent mail is where the original HTML and
//   pictures are. The pictures are cid: parts of that mail.

const QUOTE_START = [
  "#divRplyFwdMsg",
  "#appendonsend",
  "#x_divRplyFwdMsg",
  "blockquote",
  ".moz-cite-prefix",
  ".moz-forward-container",
  "div[style*='border-top:solid #E1E1E1']",
  "div[style*='border-top: solid #E1E1E1']",
  "div[style*='border-top:solid #B5C4DF']",
].join(", ");

function findPart(part, test) {
  if (test(part)) {
    return part;
  }
  for (const child of part.parts || []) {
    const found = findPart(child, test);
    if (found) {
      return found;
    }
  }
  return null;
}

function contentIds(part, map = new Map()) {
  const id = ((part.headers || {})["content-id"] || [])[0];
  if (id && part.partName) {
    map.set(id.trim().replace(/^<|>$/g, "").toLowerCase(), part.partName);
  }
  for (const child of part.parts || []) {
    contentIds(child, map);
  }
  return map;
}

function before(node, stop) {
  return !stop || Boolean(node.compareDocumentPosition(stop) & Node.DOCUMENT_POSITION_FOLLOWING);
}

// The elements of the signature in a mail written in Outlook.
function signatureElements(doc) {
  // New Outlook and Outlook on the web mark it.
  const marked = doc.querySelector("#Signature, #x_Signature, [id^='Signature'], #signature");
  if (marked) {
    return [marked];
  }
  // Classic Outlook does not: take the last table with a picture in the
  // mail's own text (before the quoted mail), with a short closing line
  // ("Sincerely,") right above it.
  const stop = doc.querySelector(QUOTE_START);
  const tables = [...doc.querySelectorAll("table")].filter(
    (table) => table.querySelector("img") && before(table, stop) && !table.parentElement.closest("table")
  );
  const table = tables[tables.length - 1];
  if (!table) {
    return [];
  }
  const elements = [table];
  let previous = table.previousElementSibling;
  while (previous && !previous.textContent.trim() && !previous.querySelector("img")) {
    previous = previous.previousElementSibling; // empty paragraphs
  }
  const text = previous ? previous.textContent.trim() : "";
  if (previous && text && text.length <= 40 && !previous.querySelector("img, table")) {
    elements.unshift(previous);
  }
  return elements;
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

async function selectedMailMessage() {
  const mailTabs = await messenger.mailTabs.query({});
  const active = mailTabs.find((tab) => tab.active) || mailTabs[0];
  if (!active) {
    throw userError("Open the mail tab and select a mail first.");
  }
  const selected = await collectList(await messenger.mailTabs.getSelectedMessages(active.id), 1);
  if (!selected.length) {
    throw userError("In the message list, select a mail you sent from Outlook that has your signature.");
  }
  return selected[0];
}

async function signatureFromSelectedMail(identityId) {
  const message = await selectedMailMessage();
  // Only a mail the user sent: a mail from someone else could carry crafted
  // HTML. The From line can be forged, so the mail must also be in a Sent
  // folder (where received mail does not go); the HTML is still sanitized.
  const folder = message.folder || {};
  const inSent = (folder.specialUse || []).includes("sent") || folder.type === "sent";
  if (!inSent || !(await myAddresses()).has(emailOf(message.author))) {
    throw userError(
      `"${message.subject}" was not sent by you. Select a mail you sent from Outlook (in Sent Items) that has your signature.`
    );
  }
  const full = await messenger.messages.getFull(message.id);
  const htmlPart = findPart(full, (part) => /^text\/html/i.test(part.contentType || "") && part.body);
  if (!htmlPart) {
    throw userError(`"${message.subject}" has no HTML text, so it has no Outlook signature.`);
  }
  const doc = new DOMParser().parseFromString(htmlPart.body, "text/html");
  const elements = signatureElements(doc);
  if (!elements.length) {
    throw userError(
      `No signature found in "${message.subject}". Select a mail you wrote in Outlook, with your signature (a table with your logo).`
    );
  }
  const cids = contentIds(full);
  const box = doc.createElement("div");
  box.append(...elements.map((element) => element.cloneNode(true)));
  for (const img of box.querySelectorAll("img")) {
    const src = (img.getAttribute("src") || "").trim();
    if (/^cid:/i.test(src)) {
      const partName = cids.get(src.slice(4).toLowerCase());
      if (partName) {
        img.setAttribute("src", await fileToDataUrl(await messenger.messages.getAttachmentFile(message.id, partName)));
      } else {
        img.remove();
      }
    } else if (!/^(data|https?):/i.test(src)) {
      img.remove();
    }
  }
  // The mail's HTML is untrusted: sanitize it here (scripts, event handlers,
  // javascript: URLs, frames, forms dropped) before it is stored.
  const clean = await messenger.outlookPane.sanitizeHtml(box.innerHTML);
  const result = await messenger.outlookPane.setSignature(identityId, clean);
  return { ...result, subject: message.subject || "" };
}

handlers["import-signature"] = fromOwnPage((message) => messenger.outlookPane.importSignature(message.identityId));
handlers["signature-from-mail"] = fromOwnPage((message) => signatureFromSelectedMail(message.identityId));
handlers["get-signature"] = fromOwnPage((message) => messenger.outlookPane.getSignature(message.identityId));
