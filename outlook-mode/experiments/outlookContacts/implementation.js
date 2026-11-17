"use strict";

/* global ExtensionAPI, ExtensionCommon, ChromeUtils, Services, Cc, Ci */

// Experiment: new contacts in the company's format, in the address book
// (about:addressbook).
//
// In many organizations, a person's display name is "Full Name (Company-Department)",
// e.g. "Phạm Minh Châu (Contoso-RND-HN)". The page script
// (contacts/contact-form.js) puts the email field first, fills in the
// display name of a known address, and adds a "Company – department" field
// that suggests the codes already used. This file gives it two read-only
// lookups, both on data already on this computer:
// - findPerson(email): the name of an address, from the address books
//   (Personal, Collected Addresses...) or else from Thunderbird's mail index
//   (Gloda: the names seen in mail headers);
// - knownCodes(): the "Company-Department" codes of the contacts in the
//   address books, with how often each is used and with which mail domains.
// Nothing goes to the network; nothing is written here (the user saves the
// contact with Thunderbird's own Save).

var { MailServices } = ChromeUtils.importESModule("resource:///modules/MailServices.sys.mjs");
var { setTimeout: ctSetTimeout, clearTimeout: ctClearTimeout } = ChromeUtils.importESModule(
  "resource://gre/modules/Timer.sys.mjs"
);

const CT_PACKAGE = "outlookcontacts";
const CT_PAGE_URL = "about:addressbook";
const CT_LOAD_ID = Date.now(); // busts the script cache when the add-on is updated
const CT_MAX_CODES = 500;
const CT_GLODA_TIMEOUT_MS = 3000;

const ctState = {
  enabled: false,
  chromeHandle: null,
  observing: false,
  windows: new Set(), // address book pages running contact-form.js
  lastError: "",
};

function ctNote(where, error) {
  ctState.lastError = `${new Date().toISOString()} ${where}: ${error && error.message ? error.message : error}`;
  console.error(`[outlook-mode] contacts: ${where}`, error);
}

// "Phạm Minh Châu (Contoso-RND-HN)" -> { name, code: "Contoso-RND-HN" }.
// A code is the text in the final brackets with at least one "-".
function ctSplitName(text) {
  const match = String(text || "").match(/^(.*?)\s*\(([^()]+)\)\s*$/);
  const code = match ? match[2].trim() : "";
  if (!match || !/^[\p{L}\p{N}&.]+(?:\s*-\s*[^-()]+)+$/u.test(code)) {
    return { name: String(text || "").trim(), code: "" };
  }
  return { name: match[1].trim(), code: code.replace(/\s*-\s*/g, "-") };
}

function ctEmail(text) {
  const email = String(text || "").trim().toLowerCase();
  return /^[^@\s"'<>()[\],;:\\]+@[^@\s"'<>()[\],;:\\]+\.[a-z]{2,}$/i.test(email) ? email : "";
}

function ctBooks() {
  return MailServices.ab.directories.filter(
    (book) => book.dirType != Ci.nsIAbManager.LDAP_DIRECTORY_TYPE
  );
}

function ctCards(book) {
  try {
    return book.childCards.filter((card) => !card.isMailList);
  } catch (error) {
    return [];
  }
}

// The code of a contact: from its display name, else from Company and
// Department.
function ctCodeOf(card) {
  const fromName = ctSplitName(card.displayName).code;
  if (fromName) {
    return fromName;
  }
  try {
    const org = card.supportsVCard ? card.vCardProperties.getFirstValue("org") : null;
    const parts = (Array.isArray(org) ? org : [org]).map((v) => String(v || "").trim()).filter(Boolean);
    return parts.length > 1 ? parts.join("-") : "";
  } catch (error) {
    return "";
  }
}

function ctKnownCodes() {
  const codes = new Map(); // lower-case code -> { code, count, domains: Map }
  for (const book of ctBooks()) {
    for (const card of ctCards(book)) {
      const code = ctCodeOf(card);
      if (!code) {
        continue;
      }
      const key = code.toLowerCase();
      const entry = codes.get(key) || { code, count: 0, domains: new Map() };
      entry.count++;
      const domain = String(card.primaryEmail || "").toLowerCase().split("@")[1];
      if (domain) {
        entry.domains.set(domain, (entry.domains.get(domain) || 0) + 1);
      }
      codes.set(key, entry);
    }
  }
  return [...codes.values()]
    .sort((a, b) => b.count - a.count || a.code.localeCompare(b.code))
    .slice(0, CT_MAX_CODES)
    .map((entry) => ({ code: entry.code, count: entry.count, domains: Object.fromEntries(entry.domains) }));
}

// The name Thunderbird's mail index knows for an address ("" if none).
function ctGlodaName(email) {
  let Gloda;
  let GlodaConstants;
  try {
    ({ Gloda } = ChromeUtils.importESModule("resource:///modules/gloda/GlodaPublic.sys.mjs"));
    ({ GlodaConstants } = ChromeUtils.importESModule("resource:///modules/gloda/GlodaConstants.sys.mjs"));
  } catch (error) {
    return Promise.resolve("");
  }
  return new Promise((resolve) => {
    const timer = ctSetTimeout(() => resolve(""), CT_GLODA_TIMEOUT_MS);
    try {
      const query = Gloda.newQuery(GlodaConstants.NOUN_IDENTITY);
      query.kind("email");
      query.value(email);
      query.getCollection({
        onItemsAdded() {},
        onItemsModified() {},
        onItemsRemoved() {},
        onQueryCompleted(collection) {
          ctClearTimeout(timer);
          const names = collection.items
            .map((identity) => String(identity.contact?.name || "").trim())
            .filter((name) => name && name.toLowerCase() != email);
          resolve(names[0] || "");
        },
      });
    } catch (error) {
      ctNote("mail index", error);
      resolve("");
    }
  });
}

// { name, source: "book" | "mail", book? } or null.
async function ctFindPerson(rawEmail) {
  const email = ctEmail(rawEmail);
  if (!email) {
    return null;
  }
  for (const book of ctBooks()) {
    let card = null;
    try {
      card = book.cardForEmailAddress(email);
    } catch (error) {
      continue;
    }
    if (card && card.displayName && card.displayName.toLowerCase() != email) {
      return { name: card.displayName, source: "book", book: book.dirName };
    }
  }
  const name = await ctGlodaName(email);
  return name ? { name, source: "mail" } : null;
}

/* ------------------------------------------------------------------ */
/* Address book pages                                                  */
/* ------------------------------------------------------------------ */

// The page and this file are both chrome code (system principal): the page
// uses these results directly.
function ctBridge() {
  return {
    enabled: () => ctState.enabled,
    findPerson: (email) => ctFindPerson(email),
    knownCodes: () => ctKnownCodes(),
    splitName: (text) => ctSplitName(text),
  };
}

function ctIsPage(win) {
  try {
    return String(win?.document?.documentURI || "").startsWith(CT_PAGE_URL);
  } catch (error) {
    return false;
  }
}

function ctInject(win) {
  if (!ctState.enabled || !ctIsPage(win) || win.outlookModeContactForm) {
    return;
  }
  try {
    win.outlookContactsBridge = ctBridge();
    Services.scriptloader.loadSubScript(`chrome://${CT_PACKAGE}/content/contact-form.js?${CT_LOAD_ID}`, win);
    if (!win.outlookModeContactForm) {
      return;
    }
    ctState.windows.add(win);
    win.addEventListener("unload", () => ctState.windows.delete(win), { once: true });
  } catch (error) {
    ctNote("inject", error);
  }
}

function ctWhenLoaded(win) {
  if (win.document.readyState == "complete") {
    ctInject(win);
  } else {
    win.addEventListener("load", () => ctInject(win), { once: true });
  }
}

const ctObserver = {
  observe(subject) {
    const win = subject?.defaultView;
    if (win && ctIsPage(win)) {
      ctWhenLoaded(win);
    }
  },
};

function ctScanOpen() {
  for (const win of Services.wm.getEnumerator("mail:3pane")) {
    for (const tab of win.gTabmail?.tabInfo || []) {
      const page = tab.browser?.contentWindow;
      if (page && ctIsPage(page)) {
        ctWhenLoaded(page);
      }
    }
  }
}

function ctRemoveAll() {
  for (const win of ctState.windows) {
    try {
      win.outlookModeContactForm?.destroy();
      delete win.outlookModeContactForm;
      delete win.outlookContactsBridge;
    } catch (error) {
      // the page is going away
    }
  }
  ctState.windows.clear();
}

function ctStart(extension) {
  if (!ctState.chromeHandle) {
    const aomStartup = Cc["@mozilla.org/addons/addon-manager-startup;1"].getService(Ci.amIAddonManagerStartup);
    const manifestURI = Services.io.newURI("manifest.json", null, extension.rootURI);
    ctState.chromeHandle = aomStartup.registerChrome(manifestURI, [["content", CT_PACKAGE, "contacts/"]]);
  }
  if (!ctState.observing) {
    Services.obs.addObserver(ctObserver, "document-element-inserted");
    ctState.observing = true;
  }
  ctScanOpen();
}

function ctStop() {
  if (ctState.observing) {
    Services.obs.removeObserver(ctObserver, "document-element-inserted");
    ctState.observing = false;
  }
  ctRemoveAll();
}

this.outlookContacts = class extends ExtensionCommon.ExtensionAPI {
  onShutdown(isAppShutdown) {
    if (isAppShutdown) {
      return;
    }
    ctStop();
    if (ctState.chromeHandle) {
      ctState.chromeHandle.destruct();
      ctState.chromeHandle = null;
    }
    Services.obs.notifyObservers(null, "startupcache-invalidate");
  }

  getAPI(context) {
    return {
      outlookContacts: {
        async configure(options) {
          ctState.enabled = Boolean(options.enabled);
          if (ctState.enabled) {
            ctStart(context.extension);
          } else {
            ctStop();
          }
          return { enabled: ctState.enabled };
        },
        async status() {
          return {
            enabled: ctState.enabled,
            pages: ctState.windows.size,
            codes: ctState.enabled ? ctKnownCodes().length : 0,
            lastError: ctState.lastError,
          };
        },
      },
    };
  }
};
