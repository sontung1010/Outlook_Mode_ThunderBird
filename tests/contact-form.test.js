"use strict";
// Run: NODE_PATH=<dir with jsdom>/node_modules node tests/contact-form.test.js
// (jsdom: `npm install jsdom@24` in any temporary folder)
// Runs contacts/contact-form.js in a jsdom copy of the address book's
// contact form (a small fake of Thunderbird's vCard editor) with a fake
// bridge (experiments/outlookContacts), and checks the company format flows.
const fs = require("fs");
const assert = require("assert");
const { JSDOM } = require("jsdom");

const SCRIPT = require("path").join(__dirname, "../outlook-mode/contacts/contact-form.js");
const html = `<!doctype html><html><head></head><body>
<form id="editContactForm"><div class="contact-header"><div class="contact-headings"><h1 id="editContactHeadingName"></h1></div></div>
<vcard-edit></vcard-edit></form></body></html>`;
const dom = new JSDOM(html, { runScripts: "outside-only", pretendToBeVisual: true });
const { window } = dom;
const { document } = window;
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const same = (actual, expected, message) => assert.strictEqual(JSON.stringify(actual), JSON.stringify(expected), message);

/* ---------------- fake Thunderbird form ---------------- */

const edit = document.querySelector("vcard-edit");
const typed = []; // input events seen by "Thunderbird"
document.getElementById("editContactForm").addEventListener("input", (event) => typed.push(event.target.id || event.target.name));

// Draws the fields as Thunderbird's editor does (replacing all children).
function draw({ email = "", display = "", org = null } = {}) {
  edit.textContent = "";
  edit.insertAdjacentHTML(
    "beforeend",
    `<fieldset id="addr-book-edit-n"><input id="firstName"><input id="vCardDisplayName"></fieldset>
     <fieldset id="addr-book-edit-email"><table><tbody id="vcard-email"><tr><td><input type="email"></td></tr></tbody></table></fieldset>
     <fieldset id="addr-book-edit-url"></fieldset>
     <fieldset id="addr-book-edit-org"><button type="button" id="vcard-add-org"></button></fieldset>`
  );
  document.querySelector('#vcard-email input').value = email;
  document.getElementById("vCardDisplayName").value = display;
  document.getElementById("vcard-add-org").addEventListener("click", function () {
    const el = document.createElement("vcard-org");
    el.innerHTML = '<input name="org"><input name="orgUnit">';
    this.before(el);
    el.querySelector("input").focus(); // Thunderbird moves the focus there
    this.hidden = true;
  });
  if (org) {
    document.getElementById("vcard-add-org").click();
    document.querySelector('vcard-org input[name="org"]').value = org[0];
    document.querySelector('vcard-org input[name="orgUnit"]').value = org[1];
  }
}

function open(options = {}, existing = false) {
  window.detailsPane.currentCard = existing ? { UID: "x" } : null;
  draw(options);
  document.body.classList.add("is-editing");
}

async function close() {
  document.body.classList.remove("is-editing");
  await sleep(10);
}

const $ = (sel) => document.querySelector(sel);
const emailInput = () => $('#vcard-email input');
const display = () => $("#vCardDisplayName").value;
const org = () => [$('vcard-org input[name="org"]')?.value, $('vcard-org input[name="orgUnit"]')?.value];
const statusText = () => ($("#olContactStatus") && !$("#olContactStatus").hidden ? $("#olContactStatus").textContent : "");
const chipCodes = () => [...document.querySelectorAll(".ol-code-chip")].map((c) => c.textContent);
const codeInput = () => $("#olCodeInput");

function type(input, value, event = "input") {
  input.value = value;
  input.dispatchEvent(new window.Event(event, { bubbles: true }));
}

/* ---------------- fake bridge ---------------- */

const CODES = [
  { code: "Fabrikam-HR-HN", count: 9, domains: { "fabrikam.example": 9 } },
  { code: "Contoso-RND-HN", count: 5, domains: { "contoso.example": 5 } },
  { code: "Contoso-MKT-HN", count: 3, domains: { "contoso.example": 3 } },
  { code: "Contoso-APP-HN", count: 1, domains: { "contoso.example": 1 } },
];
const PEOPLE = {
  "chau.pm@contoso.example": { name: "Phạm Minh Châu (Contoso-RND-HN)", source: "book", book: "Collected Addresses" },
  "trang.lt@contoso.example": { name: "Lý Thu Trang", source: "mail" },
};
const lookups = [];
window.detailsPane = { currentCard: null };
window.outlookContactsBridge = {
  enabled: () => true,
  knownCodes: () => CODES,
  findPerson: async (email) => {
    lookups.push(email);
    return PEOPLE[email] || null;
  },
};

window.eval(fs.readFileSync(SCRIPT, "utf8"));
assert.ok(window.outlookModeContactForm, "script started");
const t = window.outlookModeContactForm._test;

(async () => {
  /* 1. text helpers */
  same(t.splitName("Phạm Minh Châu (Contoso-RND-HN)"), { name: "Phạm Minh Châu", code: "Contoso-RND-HN" });
  same(t.splitName("Admin (Office)"), { name: "Admin (Office)", code: "" }, "brackets without a code are part of the name");
  same(t.splitName("A ( Contoso - MKT - HN )"), { name: "A", code: "Contoso-MKT-HN" });
  same(t.splitCode("Contoso-RND-HN"), ["Contoso", "RND-HN"]);
  const COMPANIES = new window.Set(["contoso", "fabrikam"]);
  // A known company alone in brackets (as some mail shows people) is a code
  // too, and is not kept twice.
  same(t.splitName("Vũ Anh (Contoso)", COMPANIES), { name: "Vũ Anh", code: "Contoso" });
  same(t.splitName("Vũ Anh (Contoso) (Contoso-RND&AI-HN)", COMPANIES), { name: "Vũ Anh", code: "Contoso-RND&AI-HN" });
  same(t.splitName("Vũ Anh (contoso) (Contoso-MKT-HN)", COMPANIES), { name: "Vũ Anh", code: "Contoso-MKT-HN" });
  same(t.splitName("Lan (Freelancer)", COMPANIES), { name: "Lan (Freelancer)", code: "" }, "an unknown word in brackets stays in the name");
  same(t.parseMailbox("Phạm Minh Châu (Contoso-RND-HN) <chau.pm@contoso.example>"), {
    name: "Phạm Minh Châu (Contoso-RND-HN)",
    email: "chau.pm@contoso.example",
  });
  same(t.parseMailbox('"Lý Trang" <Ha@X.vn>'), { name: "Lý Trang", email: "ha@x.vn" });
  assert.strictEqual(t.parseMailbox("a@contoso.example"), null);
  same(t.suggest(CODES, "", "a@contoso.example").map((e) => e.code), ["Contoso-RND-HN", "Contoso-MKT-HN", "Contoso-APP-HN", "Fabrikam-HR-HN"], "same domain first");
  same(t.suggest(CODES, "mkt", "").map((e) => e.code), ["Contoso-MKT-HN"]);
  same(t.suggest(CODES, "contoso hn", "").length, 3);

  /* 2. layout: email first, the code field after the name, focus in email */
  open();
  await sleep(20);
  same([...edit.children].map((e) => e.id), ["addr-book-edit-email", "addr-book-edit-n", "olCodeField", "addr-book-edit-url", "addr-book-edit-org"]);
  assert.strictEqual(document.activeElement, emailInput(), "a new contact starts in the email field");
  assert.match(statusText(), /email address first/);
  same(chipCodes(), ["Fabrikam-HR-HN", "Contoso-RND-HN", "Contoso-MKT-HN", "Contoso-APP-HN"], "most used first without an address");

  /* 3. a whole "Name (Code) <address>" pasted */
  type(emailInput(), "Phạm Minh Châu (Contoso-RND-HN) <chau.pm@contoso.example>");
  await sleep(600);
  assert.strictEqual(emailInput().value, "chau.pm@contoso.example", "bare address");
  assert.strictEqual(display(), "Phạm Minh Châu (Contoso-RND-HN)");
  same(org(), ["Contoso", "RND-HN"]);
  assert.strictEqual(codeInput().value, "Contoso-RND-HN");
  assert.strictEqual($('.ol-code-chip[aria-pressed="true"]').textContent, "Contoso-RND-HN");
  assert.ok(typed.includes("vCardDisplayName") && typed.includes("org"), "Thunderbird sees the changes as typing");
  assert.match(statusText(), /Filled from the pasted address/);
  same(lookups, [], "nothing to look up");
  await close();

  /* 4. an address already in an address book */
  open();
  await sleep(20);
  type(emailInput(), "chau.pm@contoso.example");
  await sleep(600);
  assert.match(statusText(), /already in Collected Addresses as “Phạm Minh Châu \(Contoso-RND-HN\)”/);
  assert.strictEqual(display(), "Phạm Minh Châu (Contoso-RND-HN)");
  same(org(), ["Contoso", "RND-HN"]);
  same(chipCodes().slice(0, 3), ["Contoso-RND-HN", "Contoso-MKT-HN", "Contoso-APP-HN"], "codes of the address's domain first");
  await close();

  /* 5. a name from the mail, then a code picked */
  open();
  await sleep(20);
  type(emailInput(), "trang.lt@contoso.example", "change");
  await sleep(50);
  assert.strictEqual(display(), "Lý Thu Trang");
  assert.match(statusText(), /Name found in your mail: “Lý Thu Trang”/);
  same(org(), [undefined, undefined], "no code yet");
  [...document.querySelectorAll(".ol-code-chip")].find((c) => c.textContent === "Contoso-MKT-HN").click();
  assert.strictEqual(display(), "Lý Thu Trang (Contoso-MKT-HN)");
  same(org(), ["Contoso", "MKT-HN"]);
  // Another code replaces the first one everywhere.
  [...document.querySelectorAll(".ol-code-chip")].find((c) => c.textContent === "Contoso-RND-HN").click();
  assert.strictEqual(display(), "Lý Thu Trang (Contoso-RND-HN)");
  same(org(), ["Contoso", "RND-HN"]);
  await close();

  /* 5b. a name from the mail with the company alone in brackets */
  PEOPLE["anh.vu@contoso.example"] = { name: "Vũ Anh (Contoso)", source: "mail" };
  open();
  await sleep(20);
  type(emailInput(), "anh.vu@contoso.example", "change");
  await sleep(50);
  assert.strictEqual(display(), "Vũ Anh (Contoso)");
  assert.strictEqual(codeInput().value, "Contoso", "the company is taken as a start of the code");
  same(org(), ["Contoso", ""]);
  same(chipCodes().slice(0, 3), ["Contoso-RND-HN", "Contoso-MKT-HN", "Contoso-APP-HN"], "its codes are suggested");
  [...document.querySelectorAll(".ol-code-chip")].find((c) => c.textContent === "Contoso-RND-HN").click();
  assert.strictEqual(display(), "Vũ Anh (Contoso-RND-HN)", "Contoso not twice");
  same(org(), ["Contoso", "RND-HN"]);
  // A display name saved with the company twice is tidied when edited.
  type($("#vCardDisplayName"), "Vũ Anh (Contoso) (Contoso-RND-HN)", "change");
  assert.strictEqual(display(), "Vũ Anh (Contoso-RND-HN)");
  await close();

  /* 6. unknown address: name typed, code by Enter on a few letters */
  open();
  await sleep(20);
  type(emailInput(), "an.nv@contoso.example", "change");
  await sleep(50);
  assert.match(statusText(), /not in your address books or mail yet/);
  type($("#vCardDisplayName"), "Nguyễn Văn An");
  type(codeInput(), "app");
  same(chipCodes(), ["Contoso-APP-HN"], "suggestions follow the typing");
  let submitted = false;
  document.getElementById("editContactForm").addEventListener("submit", () => (submitted = true), { once: true });
  codeInput().dispatchEvent(new window.KeyboardEvent("keydown", { key: "Enter", bubbles: true, cancelable: true }));
  assert.ok(!submitted, "Enter does not save the form");
  assert.strictEqual(codeInput().value, "Contoso-APP-HN");
  assert.strictEqual(display(), "Nguyễn Văn An (Contoso-APP-HN)");
  same(org(), ["Contoso", "APP-HN"]);
  await close();

  /* 7. a new code typed by hand, and a display name typed with its code */
  open();
  await sleep(20);
  type($("#vCardDisplayName"), "Lê Minh (Northwind-R&D-HCM)", "change");
  same(org(), ["Northwind", "R&D-HCM"], "the display name's code sets Company / Department");
  assert.strictEqual(codeInput().value, "Northwind-R&D-HCM");
  type(codeInput(), "Contoso - NEW - HN", "change");
  assert.strictEqual(display(), "Lê Minh (Contoso-NEW-HN)");
  same(org(), ["Contoso", "NEW-HN"]);
  // A display name typed without a code gets the chosen one.
  type($("#vCardDisplayName"), "Lê Minh Tuấn", "change");
  assert.strictEqual(display(), "Lê Minh Tuấn (Contoso-NEW-HN)");
  await close();

  /* 8. an existing contact: the field shows its code, no lookup */
  const before = lookups.length;
  open({ email: "a@contoso.example", display: "An", org: ["Contoso", "MKT-HN"] }, true);
  await sleep(20);
  assert.strictEqual(codeInput().value, "Contoso-MKT-HN", "code from Company / Department");
  type(emailInput(), "b@contoso.example", "change");
  await sleep(50);
  assert.strictEqual(lookups.length, before, "no lookup for an existing contact");
  assert.strictEqual(statusText(), "");

  /* 9. the editor draws the fields again: the order is kept */
  draw({ email: "x@y.vn" });
  await sleep(10);
  same([...edit.children].map((e) => e.id).slice(0, 3), ["addr-book-edit-email", "addr-book-edit-n", "olCodeField"]);
  await close();

  /* 10. destroy */
  open();
  await sleep(10);
  window.outlookModeContactForm.destroy();
  same([...edit.children].map((e) => e.id).slice(0, 2), ["addr-book-edit-n", "addr-book-edit-email"], "order back");
  assert.ok(!$("#olCodeField") && !$("#olContactStatus") && !$("#olContactStyle"));
  assert.ok(!window.outlookModeContactForm);

  console.log("contact-form: all checks passed");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
