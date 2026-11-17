"use strict";
// Run: NODE_PATH=<dir with jsdom>/node_modules node tests/cleanup.test.js
// Runs cleanup/cleanup.js in a jsdom copy of cleanup.html with a fake
// background, and checks the quick-clean cards, the filters, the
// protections and what the buttons send.
const fs = require("fs");
const path = require("path");
const assert = require("assert");
const { JSDOM } = require("jsdom");

const ROOT = path.join(__dirname, "../outlook-mode/cleanup");
const dom = new JSDOM(fs.readFileSync(path.join(ROOT, "cleanup.html"), "utf8"), { runScripts: "outside-only", pretendToBeVisual: true });
const { window } = dom;
const { document } = window;
window.HTMLElement.prototype.scrollIntoView = function () {};
let confirms = 0;
window.confirm = () => (confirms++, true);

const DAY = 864e5;
const now = Date.now();
const READ = 1;
const FLAGGED = 2;
const PINNED = 4;
const ATTACHMENT = 8;
const MB = 1024 * 1024;
// [id, folder, date, size, author, subject, flags]
const rows = [
  [1, 0, now - 2 * DAY, 20e3, "Ann <ann@x.com>", "Recent", READ],
  [2, 0, now - 100 * DAY, 30e3, "Ann <ann@x.com>", "Old 100 days", READ],
  [3, 0, now - 400 * DAY, 40e3, "Bob <bob@x.com>", "Old 400 days", READ],
  [4, 0, now - 400 * DAY, 8 * MB, "Bob <bob@x.com>", "Big and old", READ | ATTACHMENT],
  [5, 0, now - 400 * DAY, 10e3, "Bob <bob@x.com>", "Old but pinned", READ | PINNED],
  [6, 0, now - 400 * DAY, 10e3, "Bob <bob@x.com>", "Old but flagged", READ | FLAGGED],
  [7, 0, now - 60 * DAY, 15e3, "Jira <jira@company.com>", "[JIRA] Issue updated", 0],
  [8, 0, now - 10 * DAY, 15e3, "SharePoint Online <no-reply@sharepoint.com>", "Shared with you", READ],
  [9, 1, now - 500 * DAY, 25e3, "Me <me@x.com>", "Sent long ago", READ],
  [10, 2, now - 5 * DAY, 50e3, "Spam <spam@bad.com>", "Win", 0],
  [11, 3, now - 1 * DAY, 60e3, "Ann <ann@x.com>", "Deleted already", READ],
  // Bitbucket puts the person who did something as the sender name.
  [12, 0, now - 20 * DAY, 12e3, "Đặng Quốc Bảo <notifications-noreply@bitbucket.org>", "[Bitbucket] Comment", READ],
  [13, 0, now - 21 * DAY, 13e3, "Ngô Gia Huy <pullrequests-reply@bitbucket.org>", "[Bitbucket] Pull request", READ],
];
const applied = [];
let serverUsage = { used: 1.96 * 1024 ** 3, limit: 2 * 1024 ** 3, source: "server" };
const stored = {};
const progressSeen = [];
let pageListener = null;
window.messenger = {
  storage: {
    local: {
      async get(key) {
        return key in stored ? { [key]: structuredClone(stored[key]) } : {};
      },
      async set(items) {
        Object.assign(stored, structuredClone(items));
      },
    },
  },
  runtime: {
    onMessage: {
      addListener(listener) {
        pageListener = listener;
      },
    },
    async sendMessage(message) {
      switch (message.type) {
        case "cleanup-accounts":
          return [{ id: "account1", name: "Work", type: "graph" }];
        case "cleanup-scan":
          return {
            folders: [
              { id: "f0", name: "Inbox", path: "/Inbox", kind: "inbox" },
              { id: "f1", name: "Sent", path: "/Sent", kind: "sent" },
              { id: "f2", name: "Junk", path: "/Junk", kind: "junk" },
              { id: "f3", name: "Deleted Items", path: "/Trash", kind: "trash" },
            ],
            rows: structuredClone(rows),
            usage: serverUsage,
          };
        case "cleanup-apply":
          applied.push(message);
          // Half way: what the action bar says.
          pageListener({ type: "cleanup-apply-progress", handled: Math.floor(message.ids.length / 2), total: message.ids.length });
          progressSeen.push([document.querySelector("#status").textContent, document.querySelector("#act-progress").hidden, document.querySelector("#act-progress-fill").style.width]);
          return { ok: true, doneIds: message.ids, skipped: 0, notDone: 0, failed: [] };
      }
      throw new Error(`unexpected ${message.type}`);
    },
  },
};

const wait = (ms = 20) => new Promise((resolve) => setTimeout(resolve, ms));
const $ = (sel) => document.querySelector(sel);
const mailRows = () => [...document.querySelectorAll("#results tr")].slice(1).filter((tr) => !tr.classList.contains("month"));
const resultIds = () => mailRows().map((tr) => tr.children[3].textContent);
const card = (re) => [...document.querySelectorAll(".preset")].find((c) => re.test(c.textContent));

(async () => {
  window.eval(fs.readFileSync(path.join(ROOT, "cleanup.js"), "utf8"));
  await wait(50);

  // 1. Usage against the quota; Deleted Items counted apart.
  assert.match($("#usage-text").textContent, /1\.96 GB of 2\.00 GB used \(98%\)/);
  assert.match($("#usage-text").textContent, /Deleted Items: 1 mail/);
  assert.ok($("#usage-fill").classList.contains("full"), "almost full shows red");
  assert.ok($("#limit-box").hidden, "no need to type the limit when the server gives it");

  // 2. Default: everything but Junk and Deleted Items, pinned and flagged left out.
  assert.match($("#summary").textContent, /^9 mails/);

  // 2b. Automatic mail: first card, its own senders table, shown by service.
  assert.match(document.querySelector(".preset").textContent, /^🤖 Notifications & no-reply — start here/);
  const autoSenders = [...document.querySelectorAll("#auto-senders tr")].slice(1).map((tr) => tr.children[0].textContent);
  assert.ok(autoSenders.includes("🤖 Bitbucket — notifications-noreply@bitbucket.org +1 more"), autoSenders.join(" | "));
  assert.doesNotMatch($("#senders").textContent, /Đặng Quốc Bảo|Ngô Gia Huy|jira|sharepoint/i, "not listed as people");

  // 3. Older than 1 year: old Inbox and Sent mail, never the pinned or flagged one.
  [...document.querySelectorAll(".chip")].find((c) => c.textContent.startsWith("1 year")).click();
  await wait();
  assert.deepEqual(resultIds().sort(), ["Big and old", "Old 400 days", "Sent long ago"].map((s) => (s === "Big and old" ? "📎 Big and old" : s)).sort());
  assert.match([...document.querySelectorAll(".chip")].find((c) => c.textContent.startsWith("3 months")).textContent, /3 months: 4 ·/);

  // 4. Large mail; notifications & newsletters; unread for over a month.
  card(/Large mail/).click();
  await wait();
  assert.deepEqual(resultIds(), ["📎 Big and old"]);
  card(/Notifications/).click();
  await wait();
  assert.deepEqual(resultIds().sort(), ["Shared with you", "[JIRA] Issue updated", "[Bitbucket] Comment", "[Bitbucket] Pull request"].sort());
  const fromOf = (subject) => mailRows().find((tr) => tr.children[3].textContent === subject).children[2].textContent;
  assert.equal(fromOf("[Bitbucket] Comment"), "🤖 Bitbucket · Đặng Quốc Bảo");
  assert.equal(fromOf("[Bitbucket] Pull request"), "🤖 Bitbucket · Ngô Gia Huy");
  // Review of the Bitbucket row: all its addresses.
  [...document.querySelectorAll("#auto-senders tr")].find((tr) => /Bitbucket/.test(tr.textContent)).querySelector("button").click();
  await wait();
  assert.equal($("#from").value, "bitbucket.org");
  assert.equal(resultIds().length, 2);
  card(/Unread for over a month/).click();
  await wait();
  assert.deepEqual(resultIds(), ["[JIRA] Issue updated"]);

  // 5. Flagged mail can be included on purpose; pinned mail never.
  assert.equal($("#protect-pinned"), null, "no way to include pinned mail");
  card(/Large mail/).click();
  await wait();
  $("#min-size").value = "0";
  $("#from").value = "bob";
  $("#from").dispatchEvent(new window.Event("input"));
  await wait();
  assert.equal(resultIds().length, 2, "pinned and flagged of Bob left out");
  $("#protect-flagged").checked = false;
  $("#protect-flagged").dispatchEvent(new window.Event("change"));
  await wait();
  assert.equal(resultIds().length, 3);
  assert.ok(!resultIds().includes("Old but pinned"));

  // 6. Unticking a mail keeps it; the button sends the rest, to Deleted Items.
  const boxes = mailRows().map((tr) => tr.querySelector("input"));
  boxes[0].checked = false;
  boxes[0].dispatchEvent(new window.Event("change"));
  await wait();
  assert.match($("#summary").textContent, /^2 mails .* selected \(of 3 found\)/);
  assert.equal($("#trash").textContent, "🗑 Move 2 mails to Deleted Items");
  assert.match($("#act-label").textContent, /^Act on the 2 mails \(.*\) ticked in the list above/);
  $("#trash").click();
  await wait(50);
  assert.equal(applied.at(-1).action, "trash");
  assert.equal(applied.at(-1).ids.length, 2);
  assert.deepEqual(progressSeen.at(-1), ["Moving to Deleted Items: 1 of 2 mails (" + progressSeen.at(-1)[0].match(/\((.*)\)/)[1] + ") — 50%…", false, "50%"]);
  assert.ok($("#act-progress").hidden, "progress bar gone when done");
  assert.match($("#status").textContent, /Moved to Deleted Items: 2 mails/);
  assert.match($("#usage-text").textContent, /Deleted Items: 3 mails/, "they now count in Deleted Items");
  assert.equal(confirms, 0, "no question for a few mails sent to Deleted Items");

  // 7. Deleted Items: only "Empty" (permanent, asked first).
  card(/Deleted Items/).click();
  await wait();
  assert.equal(resultIds().length, 3);
  assert.ok($("#trash").hidden && $("#archive").hidden);
  assert.equal($("#delete").textContent, "Empty Deleted Items (3 mails)…");
  $("#delete").click();
  await wait(50);
  assert.equal(confirms, 1, "permanent delete is confirmed");
  assert.equal(applied.at(-1).action, "delete");
  assert.equal(applied.at(-1).ids.length, 3);
  assert.doesNotMatch($("#usage-text").textContent, /Deleted Items:/, "emptied");

  // 8. Junk card and top senders.
  assert.match(card(/^Junk/).textContent, /1 mail/);
  assert.match($("#senders").textContent, /Bob <bob@x\.com>/);

  // 9b. Month headings: a box per month, for all its mails.
  $("#reset").click();
  await wait();
  const monthRows = () => [...document.querySelectorAll("#results tr.month")];
  const mailsIn = (month) => {
    const out = [];
    for (let tr = month.nextElementSibling; tr && !tr.classList.contains("month"); tr = tr.nextElementSibling) {
      out.push(tr);
    }
    return out;
  };
  const count = () => Number($("#summary").textContent.match(/^(\d+)/)[1]);
  const before = count();
  assert.ok(monthRows().length >= 2, "a heading per month");
  const index = monthRows().findIndex((tr) => mailsIn(tr).length >= 2);
  const month = () => monthRows()[index];
  const size = mailsIn(month()).length;
  assert.match(month().textContent, new RegExp(`— ${size} mails`), month().textContent);
  month().querySelector("input").click();
  await wait();
  assert.equal(count(), before - size, "whole month unticked");
  assert.ok(mailsIn(month()).every((tr) => !tr.querySelector("input").checked));
  mailsIn(month())[0].querySelector("input").click();
  await wait();
  assert.ok(month().querySelector("input").indeterminate, "part of the month");
  assert.match(month().textContent, /1 ticked/);
  month().querySelector("input").click();
  await wait();
  assert.equal(count(), before, "whole month ticked again");
  $("#sort").value = "-size";
  $("#sort").dispatchEvent(new window.Event("change"));
  await wait();
  assert.equal(document.querySelectorAll("#results tr.month").length, 0, "no months when sorted by size");
  $("#sort").value = "date";
  $("#sort").dispatchEvent(new window.Event("change"));

  // 9. No limit from the server (Microsoft 365): typed once, kept.
  serverUsage = null;
  $("#rescan").click();
  await wait(50);
  assert.ok(!$("#limit-box").hidden);
  assert.match($("#usage-text").textContent, /^\d+ mails · .* in this account/);
  assert.ok($("#usage-fill").parentElement.hidden, "no meter without a limit");
  $("#limit").value = "2";
  $("#limit").dispatchEvent(new window.Event("change"));
  await wait();
  assert.match($("#usage-text").textContent, /^.* of 2\.00 GB used \(\d+%\) · .* free · \d+ mails/);
  assert.ok(!$("#usage-fill").parentElement.hidden);
  assert.equal(stored.cleanupLimits.account1, 2 * 1024 ** 3);
  $("#rescan").click();
  await wait(50);
  assert.equal($("#limit").value, "2", "kept after a new scan");
  assert.match($("#usage-text").textContent, /of 2\.00 GB used/);

  console.log("all cleanup scenarios passed");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
