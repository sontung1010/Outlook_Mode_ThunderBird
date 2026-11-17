"use strict";

// @mention: contact search for the compose script, and adding the chosen
// person to the To field.

const MAX_MENTION_RESULTS = 8;
let contactCache = null;
let contactStats = { books: 0, contacts: 0, loadedAt: "", error: "" };

function parseVCard(vCard) {
  const result = { fn: "", emails: [] };
  if (!vCard) {
    return result;
  }
  const lines = vCard.replace(/\r?\n[ \t]/g, "").split(/\r?\n/);
  for (const line of lines) {
    const colon = line.indexOf(":");
    if (colon < 0) {
      continue;
    }
    const key = line.slice(0, colon).split(";")[0].toUpperCase();
    const value = line
      .slice(colon + 1)
      .replace(/\\n/gi, " ")
      .replace(/\\([,;\\])/g, "$1")
      .trim();
    if (key === "FN" && !result.fn) {
      result.fn = value;
    } else if (key === "EMAIL" && value) {
      result.emails.push(value);
    }
  }
  return result;
}

function contactInfo(node) {
  const { fn, emails } = parseVCard(node.vCard);
  const props = node.properties || {};
  const name =
    fn || props.DisplayName || [props.FirstName, props.LastName].filter(Boolean).join(" ");
  if (!emails.length && props.PrimaryEmail) {
    emails.push(props.PrimaryEmail);
  }
  return { name, email: emails[0] || "" };
}

async function loadContacts() {
  const books = await messenger.addressBooks.list(true);
  const seen = new Set();
  const contacts = [];
  for (const book of books) {
    for (const node of book.contacts || []) {
      const { name, email } = contactInfo(node);
      if (!email || seen.has(email.toLowerCase())) {
        continue;
      }
      seen.add(email.toLowerCase());
      const fullName = name || email;
      // "Name (Department)" -> "Name" for the text shown in the body.
      const label = fullName.replace(/\s*\([^)]*\)\s*$/, "") || fullName;
      contacts.push({
        name: fullName,
        email,
        label,
        haystack: fold(`${fullName} ${email}`),
        labelFolded: fold(label),
      });
    }
  }
  contacts.sort((a, b) => a.name.localeCompare(b.name, "vi"));
  contactStats = {
    books: books.length,
    contacts: contacts.length,
    loadedAt: new Date().toISOString(),
    error: "",
  };
  return contacts;
}

async function getContacts() {
  if (!contactCache) {
    contactCache = loadContacts().catch((error) => {
      contactCache = null;
      contactStats.error = String(error && error.message ? error.message : error);
      throw error;
    });
  }
  return contactCache;
}

async function searchContacts(query) {
  const words = fold(query).split(/\s+/).filter(Boolean);
  if (!words.length) {
    return [];
  }
  const contacts = await getContacts();
  const matches = contacts.filter((c) => words.every((w) => c.haystack.includes(w)));
  // Contacts whose name has a word starting with the first typed word come first.
  const rank = (c) =>
    c.labelFolded.split(/\s+/).some((part) => part.startsWith(words[0])) ? 0 : 1;
  matches.sort((a, b) => rank(a) - rank(b));
  return matches
    .slice(0, MAX_MENTION_RESULTS)
    .map(({ name, email, label }) => ({ name, email, label }));
}

// Adds a @mentioned person to To. Only someone from the address books (as
// the mention popup offers), with the address and name the address book
// has, so the page cannot add other recipients (an "address" such as
// "a@x.com>, b@y.com" would add two).
async function addRecipient(tabId, email) {
  const wanted = typeof email === "string" ? email.trim().toLowerCase() : "";
  if (!EMAIL_PATTERN.test(wanted)) {
    return { added: false };
  }
  const contact = (await getContacts()).find((c) => c.email.toLowerCase() === wanted);
  if (!contact) {
    return { added: false };
  }
  const details = await messenger.compose.getComposeDetails(tabId);
  const all = [...(details.to || []), ...(details.cc || []), ...(details.bcc || [])];
  if (all.some((r) => emailOf(r) === wanted)) {
    return { added: false };
  }
  // Quotes, angle brackets, backslashes and control characters cannot end
  // the name part early.
  const name = String(contact.name || "")
    .replace(/["<>\\\u0000-\u001f\u007f]/g, "")
    .trim()
    .slice(0, 200);
  const to = [...(details.to || []), name ? `"${name}" <${contact.email}>` : contact.email];
  await messenger.compose.setComposeDetails(tabId, { to });
  return { added: true };
}

function invalidateContacts() {
  contactCache = null;
}

for (const event of [
  messenger.contacts?.onCreated,
  messenger.contacts?.onUpdated,
  messenger.contacts?.onDeleted,
  messenger.addressBooks?.onCreated,
  messenger.addressBooks?.onUpdated,
  messenger.addressBooks?.onDeleted,
]) {
  event?.addListener(invalidateContacts);
}

handlers["mention-search"] = (message) =>
  searchContacts(message.query || "").catch((error) => {
    logError("mention search", error);
    return [];
  });

handlers["mention-add"] = (message, sender) =>
  sender.tab?.type === "messageCompose" ? addRecipient(sender.tab.id, message.email) : { added: false };
