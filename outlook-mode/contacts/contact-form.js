"use strict";

/* global detailsPane */

// Runs in the address book (about:addressbook), loaded by the
// outlookContacts experiment, which sets window.outlookContactsBridge.
//
// Contacts in the company's format: the display name is
// "Full Name (Company-Department)", e.g. "Phạm Minh Châu (Contoso-RND-HN)",
// and Company / Department hold "Contoso" / "RND-HN".
// - The Email Addresses field comes first; a new contact starts there.
// - Pasting "Phạm Minh Châu (Contoso-RND-HN) <chau.pm@contoso.example>"
//   (as copied from a To/From line) fills the address, the display name,
//   Company and Department.
// - Typing an address of a new contact looks up its name in the address
//   books (Collected Addresses...) and in the mail index, and fills an
//   empty display name with it.
// - A "Company – department" field under the name suggests the codes
//   already used by contacts (those of the same mail domain first). Picking
//   or typing one sets Company, Department and the "(…)" of the display
//   name; typing a display name with a "(…)" code sets them too.
// Nothing is saved here: the user clicks Save as usual.

(() => {
  const bridge = window.outlookContactsBridge;
  if (!bridge || window.outlookModeContactForm) {
    return;
  }

  const STATUS_ID = "olContactStatus";
  const STYLE_ID = "olContactStyle";
  const CODE_FIELD_ID = "olCodeField";
  const CODE_INPUT_ID = "olCodeInput";
  const DEBOUNCE_MS = 500;
  const MAX_CHIPS = 8;
  const EMAIL_INPUT = '#vcard-email input[type="email"]';
  const ADDRESS = /[^@\s"'<>()[\],;:\\]+@[^@\s"'<>()[\],;:\\]+\.[a-z]{2,}/i;

  let form = null;
  let edit = null;
  let status = null;
  let codeField = null;
  let codeInput = null;
  let chips = null;
  let codes = []; // [{ code, count, domains }]
  let bodyObserver = null;
  let editObserver = null;
  let timer = 0;
  let session = 0;
  let editing = false;
  let done = new Set(); // addresses looked up in this form
  let autoName = ""; // the display name this script filled in
  let destroyed = false;

  /* ---------------------------------------------------------------- */
  /* Text helpers (also used by the tests)                             */
  /* ---------------------------------------------------------------- */

  function addressIn(text) {
    const match = String(text || "").match(ADDRESS);
    return match ? match[0].toLowerCase() : "";
  }

  // "Contoso - RND-HN " -> "Contoso-RND-HN"
  function cleanCode(text) {
    return String(text || "")
      .trim()
      .replace(/^\(|\)$/g, "")
      .replace(/\s*-\s*/g, "-")
      .replace(/\s+/g, " ")
      .trim();
  }

  // Company names known for this form: the first part of the known codes,
  // and the name of the address's mail domain ("contoso" of
  // contoso.example).
  function knownCompanies() {
    const names = new Set(codes.map((entry) => splitCode(entry.code)[0].toLowerCase()).filter(Boolean));
    const domain = (currentAddress().split("@")[1] || "").split(".")[0];
    if (domain) {
      names.add(domain.toLowerCase());
    }
    return names;
  }

  // The text in a final "(…)": a full code ("Contoso-RND-HN", with at
  // least one "-"), or a known company alone ("Contoso", as some mail
  // shows people: "Vũ Anh (Contoso)"). "" for anything else.
  function codeIn(text, companies) {
    const code = cleanCode(text);
    if (/^[\p{L}\p{N}&.]+(?:-[^-()]+)+$/u.test(code)) {
      return code;
    }
    return /^[\p{L}\p{N}&.]+$/u.test(code) && companies.has(code.toLowerCase()) ? code : "";
  }

  // "Phạm Minh Châu (Contoso-RND-HN)" -> { name, code }. A company
  // alone in brackets before the code is dropped too:
  // "Vũ Anh (Contoso) (Contoso-RND&AI-HN)" -> "Vũ Anh" + code.
  function splitName(text, companies = knownCompanies()) {
    const tail = /^(.*?)\s*\(([^()]+)\)\s*$/;
    const match = String(text || "").match(tail);
    const code = match ? codeIn(match[2], companies) : "";
    if (!code) {
      return { name: String(text || "").trim(), code: "" };
    }
    let name = match[1].trim();
    const company = splitCode(code)[0].toLowerCase();
    for (let inner = name.match(tail); inner; inner = name.match(tail)) {
      const word = cleanCode(inner[2]);
      if (word.toLowerCase() !== company && word.toLowerCase() !== code.toLowerCase()) {
        break;
      }
      name = inner[1].trim();
    }
    return { name, code };
  }

  function joinName(name, code) {
    return name ? (code ? `${name} (${code})` : name) : "";
  }

  // "Contoso-RND-HN" -> ["Contoso", "RND-HN"]
  function splitCode(code) {
    const [company, ...rest] = cleanCode(code).split("-");
    return [company || "", rest.join("-")];
  }

  // 'Name (Code) <a@b.c>' or '"Name" <a@b.c>' -> { name, email }; null when
  // the text is not of that form.
  function parseMailbox(text) {
    const match = String(text || "").match(/^\s*"?([^"<>]*?)"?\s*<\s*([^<>\s]+@[^<>\s]+)\s*>\s*$/);
    if (!match || !addressIn(match[2])) {
      return null;
    }
    return { name: match[1].trim(), email: addressIn(match[2]) };
  }

  // Known codes, best first for an address: its mail domain's codes, then
  // the most used; filtered by what is typed.
  function suggest(list, typed, email) {
    const domain = String(email || "").split("@")[1] || "";
    const words = cleanCode(typed).toLowerCase().split(/[-\s]+/).filter(Boolean);
    return list
      .filter((entry) => words.every((word) => entry.code.toLowerCase().includes(word)))
      .map((entry) => ({ entry, score: (entry.domains?.[domain] || 0) * 1000 + entry.count }))
      .sort((a, b) => b.score - a.score || a.entry.code.localeCompare(b.entry.code))
      .map(({ entry }) => entry);
  }

  /* ---------------------------------------------------------------- */
  /* Page elements                                                     */
  /* ---------------------------------------------------------------- */

  function ensureStyle() {
    if (document.getElementById(STYLE_ID)) {
      return;
    }
    const style = document.createElement("style");
    style.id = STYLE_ID;
    style.textContent = `
#${STATUS_ID} {
  margin: 4px 0 0;
  font-size: 0.9em;
  color: light-dark(#424242, #c8c8c8);
}
#${STATUS_ID}[data-state="found"] { color: light-dark(#0e700e, #6ccb5f); }
#${STATUS_ID}[data-state="warn"] { color: light-dark(#a4262c, #ff99a4); }
vcard-edit > #addr-book-edit-email[data-ol-first],
vcard-edit > #${CODE_FIELD_ID} { grid-column: 1 / -1; }
#${CODE_FIELD_ID} legend { font-weight: 600; margin-block-end: 6px; }
#${CODE_INPUT_ID} { inline-size: 100%; box-sizing: border-box; }
#${CODE_FIELD_ID} .ol-code-chips {
  display: flex;
  flex-wrap: wrap;
  gap: 4px 6px;
  margin-top: 6px;
}
#${CODE_FIELD_ID} .ol-code-chip {
  font: inherit;
  font-size: 0.9em;
  min-height: 0;
  margin: 0;
  padding: 2px 10px;
  border-radius: 12px;
  border: 1px solid color-mix(in srgb, currentColor 30%, transparent);
  background: transparent;
  color: inherit;
  cursor: pointer;
}
#${CODE_FIELD_ID} .ol-code-chip:hover { background: color-mix(in srgb, currentColor 10%, transparent); }
#${CODE_FIELD_ID} .ol-code-chip[aria-pressed="true"] {
  border-color: light-dark(#0f6cbd, #479ef5);
  color: light-dark(#0f6cbd, #479ef5);
}
#${CODE_FIELD_ID} .ol-code-hint { font-size: 0.85em; opacity: 0.75; margin-top: 4px; }
`;
    document.head.append(style);
  }

  function statusLine() {
    if (status && status.isConnected) {
      return status;
    }
    const headings = form?.querySelector(".contact-headings");
    if (!headings) {
      return null;
    }
    status = document.createElement("p");
    status.id = STATUS_ID;
    status.setAttribute("role", "status");
    status.setAttribute("aria-live", "polite");
    headings.append(status);
    return status;
  }

  function show(state, text) {
    const line = statusLine();
    if (line) {
      line.hidden = !text;
      line.dataset.state = state;
      line.textContent = text;
    }
  }

  function buildCodeField() {
    codeField = document.createElement("fieldset");
    codeField.id = CODE_FIELD_ID;
    codeField.className = "addr-book-edit-fieldset fieldset-reset";
    const legend = document.createElement("legend");
    legend.textContent = "Company – department";
    codeInput = document.createElement("input");
    codeInput.type = "text";
    codeInput.id = CODE_INPUT_ID;
    codeInput.autocomplete = "off";
    codeInput.spellcheck = false;
    codeInput.placeholder = "e.g. Contoso-RND-HN: type to search the codes already used";
    codeInput.setAttribute("aria-label", "Company – department");
    chips = document.createElement("div");
    chips.className = "ol-code-chips";
    const hint = document.createElement("div");
    hint.className = "ol-code-hint";
    hint.textContent =
      "Sets Company and Department, and the (…) after the display name. Codes of your contacts are suggested, those of the same mail domain first.";
    codeField.append(legend, codeInput, chips, hint);
    codeInput.addEventListener("input", () => renderChips());
    codeInput.addEventListener("change", () => applyCode(codeInput.value));
    codeInput.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        // Enter would save the form: it takes the first suggestion instead.
        event.preventDefault();
        event.stopPropagation();
        const first = chips.querySelector(".ol-code-chip");
        const typed = cleanCode(codeInput.value);
        const exact = codes.some((entry) => entry.code.toLowerCase() === typed.toLowerCase());
        applyCode(!exact && first ? first.dataset.code : typed);
      }
    });
  }

  function renderChips() {
    if (!chips) {
      return;
    }
    chips.textContent = "";
    const current = cleanCode(codeInput.value).toLowerCase();
    // A chosen code shows all suggestions (it is marked); typing filters.
    const chosen = codes.some((entry) => entry.code.toLowerCase() === current);
    for (const entry of suggest(codes, chosen ? "" : codeInput.value, currentAddress()).slice(0, MAX_CHIPS)) {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "ol-code-chip";
      chip.dataset.code = entry.code;
      chip.textContent = entry.code;
      chip.title = `Used by ${entry.count} contact${entry.count === 1 ? "" : "s"}`;
      chip.setAttribute("aria-pressed", String(entry.code.toLowerCase() === current));
      chip.addEventListener("click", () => applyCode(entry.code));
      chips.append(chip);
    }
  }

  // Email Addresses first, "Company – department" right after the name. The
  // editor draws its fields again each time a contact is opened, so this
  // runs after each drawing.
  function place() {
    if (!edit || !editing) {
      return;
    }
    const email = edit.querySelector(":scope > #addr-book-edit-email");
    const name = edit.querySelector(":scope > #addr-book-edit-n");
    if (!email || !name) {
      return;
    }
    email.dataset.olFirst = "";
    if (edit.firstElementChild !== email) {
      edit.insertBefore(email, edit.firstElementChild);
    }
    if (name.nextElementSibling !== codeField) {
      name.after(codeField);
    }
  }

  /* ---------------------------------------------------------------- */
  /* Form fields                                                       */
  /* ---------------------------------------------------------------- */

  function isNewContact() {
    return !detailsPane.currentCard;
  }

  function emailInputs() {
    return [...(form?.querySelectorAll(EMAIL_INPUT) || [])];
  }

  function currentAddress() {
    for (const input of emailInputs()) {
      const address = addressIn(input.value);
      if (address) {
        return address;
      }
    }
    return "";
  }

  function displayInput() {
    return form?.querySelector("#vCardDisplayName") || null;
  }

  // Sets a field as typing would (Thunderbird then marks the form changed
  // and updates the name at the top).
  function setValue(input, value) {
    if (input && input.value !== value) {
      input.value = value;
      input.dispatchEvent(new Event("input", { bubbles: true }));
    }
  }

  function orgInputs(create) {
    let org = form.querySelector("vcard-org");
    if (!org && create) {
      const add = form.querySelector("#vcard-add-org");
      if (!add) {
        return null;
      }
      const focused = document.activeElement;
      add.click(); // Thunderbird adds Job title, Role, Company, Department
      org = form.querySelector("vcard-org");
      focused?.focus?.();
    }
    if (!org) {
      return null;
    }
    return {
      company: org.querySelector('input[name="org"]'),
      unit: org.querySelector('input[name="orgUnit"]'),
    };
  }

  function codeFromForm() {
    const fromName = splitName(displayInput()?.value).code;
    if (fromName) {
      return fromName;
    }
    const org = orgInputs(false);
    const parts = [org?.company?.value, org?.unit?.value].map((v) => String(v || "").trim()).filter(Boolean);
    return parts.length > 1 ? cleanCode(parts.join("-")) : "";
  }

  // Company, Department and the "(…)" of the display name from a code.
  function applyCode(raw) {
    const code = cleanCode(raw);
    codeInput.value = code;
    renderChips();
    const [company, unit] = splitCode(code);
    const org = orgInputs(Boolean(code));
    if (org) {
      setValue(org.company, company);
      setValue(org.unit, unit);
    }
    const display = displayInput();
    const { name } = splitName(display?.value);
    if (display && name) {
      const value = joinName(name, code);
      if (display.value === autoName) {
        autoName = value;
      }
      setValue(display, value);
    }
  }

  // A full name, maybe with its code, into an empty display name (or one
  // this script filled in before).
  function applyName(fullName) {
    const display = displayInput();
    if (!display || (display.value.trim() && display.value !== autoName)) {
      return false;
    }
    const { name, code } = splitName(fullName);
    const value = joinName(name, code || cleanCode(codeInput.value));
    autoName = value;
    setValue(display, value);
    if (code) {
      applyCode(code);
    }
    return true;
  }

  /* ---------------------------------------------------------------- */
  /* Email lookup                                                      */
  /* ---------------------------------------------------------------- */

  async function lookup() {
    const mySession = session;
    // A whole "Name (Code) <address>" pasted into an email field.
    for (const input of emailInputs()) {
      const mailbox = parseMailbox(input.value);
      if (mailbox) {
        setValue(input, mailbox.email);
        if (mailbox.name) {
          applyName(mailbox.name);
          show("found", `Filled from the pasted address: ${mailbox.name}.`);
          done.add(mailbox.email);
          renderChips();
          return;
        }
      }
    }
    const email = currentAddress();
    renderChips(); // the suggestions follow the mail domain
    if (!email || done.has(email)) {
      return;
    }
    done.add(email);
    let person = null;
    try {
      person = await bridge.findPerson(email);
    } catch (error) {
      person = null;
    }
    if (destroyed || !editing || mySession !== session || currentAddress() !== email) {
      return;
    }
    if (!person) {
      show("idle", `${email} is not in your address books or mail yet: type the display name, then pick the company – department below.`);
      return;
    }
    const filled = applyName(person.name);
    const where = person.source === "book" ? `in ${person.book}` : "in your mail";
    if (person.source === "book") {
      show("warn", `${email} is already ${where} as “${person.name}”.${filled ? " The name was filled in." : ""}`);
    } else {
      show("found", `Name found ${where}: “${person.name}”.${filled ? " Check it, then Save." : ""}`);
    }
  }

  function schedule(delay) {
    clearTimeout(timer);
    timer = setTimeout(() => {
      timer = 0;
      lookup().catch((error) => console.error("[outlook-mode] contact form", error));
    }, delay);
  }

  function onInput(event) {
    if (!editing) {
      return;
    }
    const target = event.target;
    if (target.matches?.(EMAIL_INPUT) && isNewContact()) {
      schedule(event.type === "change" ? 0 : DEBOUNCE_MS);
    } else if (target.id === "vCardDisplayName" && event.type === "change") {
      // A display name typed with its code sets Company / Department; one
      // typed without, gets the code chosen below.
      const { name, code } = splitName(target.value);
      if (code) {
        if (code !== cleanCode(codeInput.value)) {
          applyCode(code);
        } else if (joinName(name, code) !== target.value.trim()) {
          setValue(target, joinName(name, code)); // e.g. a company twice in brackets
        }
      } else if (name && cleanCode(codeInput.value)) {
        applyCode(codeInput.value);
      }
    }
  }

  /* ---------------------------------------------------------------- */
  /* Opening and closing a contact                                     */
  /* ---------------------------------------------------------------- */

  function editingChanged() {
    const now = document.body.classList.contains("is-editing");
    if (now === editing) {
      return;
    }
    editing = now;
    session++;
    clearTimeout(timer);
    done = new Set();
    autoName = "";
    if (!editing || !bridge.enabled()) {
      show("idle", "");
      return;
    }
    try {
      codes = bridge.knownCodes() || [];
    } catch (error) {
      codes = [];
    }
    place();
    codeInput.value = codeFromForm();
    renderChips();
    if (isNewContact()) {
      show("idle", "Type or paste the email address first: the name is filled in from your address books and mail.");
      // After Thunderbird puts the focus in the first name field.
      setTimeout(() => {
        const input = emailInputs()[0];
        if (input && editing && !input.value) {
          input.focus();
        }
      }, 0);
      if (currentAddress()) {
        schedule(0); // opened with an address in it ("Add to address book")
      }
    } else {
      show("idle", "");
    }
  }

  function start() {
    form = document.getElementById("editContactForm");
    edit = form?.querySelector("vcard-edit");
    if (!form || !edit || typeof detailsPane === "undefined") {
      return false;
    }
    ensureStyle();
    buildCodeField();
    form.addEventListener("input", onInput, true);
    form.addEventListener("change", onInput, true);
    bodyObserver = new MutationObserver(editingChanged);
    bodyObserver.observe(document.body, { attributes: true, attributeFilter: ["class"] });
    editObserver = new MutationObserver(place);
    editObserver.observe(edit, { childList: true });
    editingChanged();
    return true;
  }

  function destroy() {
    destroyed = true;
    clearTimeout(timer);
    bodyObserver?.disconnect();
    editObserver?.disconnect();
    form?.removeEventListener("input", onInput, true);
    form?.removeEventListener("change", onInput, true);
    codeField?.remove();
    status?.remove();
    document.getElementById(STYLE_ID)?.remove();
    const email = edit?.querySelector(":scope > #addr-book-edit-email");
    const name = edit?.querySelector(":scope > #addr-book-edit-n");
    if (email && name) {
      delete email.dataset.olFirst;
      name.after(email);
    }
    delete window.outlookModeContactForm;
  }

  if (!start()) {
    return;
  }
  window.outlookModeContactForm = {
    destroy,
    _test: { addressIn, cleanCode, splitName, splitCode, parseMailbox, suggest },
  };
})();
