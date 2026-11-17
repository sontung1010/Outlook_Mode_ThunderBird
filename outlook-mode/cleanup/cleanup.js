"use strict";

// Mailbox cleanup page. The background (background/cleanup.js) reads an
// account's mail once; everything here filters that list at once and sends
// the chosen message ids back to be moved or deleted.

const READ = 1;
const FLAGGED = 2;
const PINNED = 4;
const ATTACHMENT = 8;
const DAY = 864e5;
const SHOWN_ROWS = 300;
const SPECIAL = ["trash", "junk"];
// Senders of notifications and newsletters: the address, the domain of a
// service that only sends notifications, or the name.
const AUTOMATED_ADDRESS =
  /^(no-?reply|do-?not-?reply|donotreply|notifications?|notify|alerts?|newsletters?|news|mailer|marketing|promo|bounces?|updates?|digest|info|support|system|jira|confluence)([+._-]|@)|(^|[+._-])(no-?reply|do-?not-?reply|donotreply|notifications?|mailer-daemon|bounces?)([+._-]|@)/i;
// Services that put the name of the person who did something (opened a
// pull request, commented) as the sender name: "Đặng Quốc Bảo
// <notifications-noreply@bitbucket.org>" is Bitbucket, not a person.
const AUTOMATED_DOMAIN =
  /(^|\.)(bitbucket\.org|atlassian\.net|atlassian\.com|github\.com|gitlab\.com|sharepointonline\.com|yammer\.com|linkedin\.com|facebookmail\.com|accountprotection\.microsoft\.com|zoom\.us|slack\.com|trello\.com|asana\.com|notion\.so|figma\.com|docs\.google\.com|calendar-notification\.google\.com)$/i;
const AUTOMATED_NAME = /\b(jira|confluence|sharepoint|teams|viva|newsletter|notification|no-?reply|linkedin|digest)\b|\bvia\b/i;
const SERVICE_NAMES = {
  atlassian: "Atlassian (Jira, Confluence)",
  bitbucket: "Bitbucket",
  facebookmail: "Facebook",
  github: "GitHub",
  gitlab: "GitLab",
  linkedin: "LinkedIn",
  sharepointonline: "SharePoint",
};

let data = null; // { folders, rows: [id, folder, date, size, author, subject, flags], usage }
let filters = null;
let unticked = new Set(); // ids the user left out
let activePreset = null;

const $ = (id) => document.getElementById(id);

function el(tag, props = {}, ...children) {
  const node = document.createElement(tag);
  for (const [key, value] of Object.entries(props)) {
    if (key === "class") {
      node.className = value;
    } else if (key === "text") {
      node.textContent = value;
    } else if (key.startsWith("on")) {
      node.addEventListener(key.slice(2), value);
    } else if (value !== false && value != null) {
      node.setAttribute(key, value === true ? "" : value);
    }
  }
  node.append(...children.filter((child) => child != null));
  return node;
}

function fold(text) {
  return String(text || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .toLowerCase();
}

function formatSize(bytes) {
  if (bytes >= 1024 ** 3) {
    return `${(bytes / 1024 ** 3).toFixed(2)} GB`;
  }
  if (bytes >= 1024 ** 2) {
    return `${(bytes / 1024 ** 2).toFixed(1)} MB`;
  }
  return bytes ? `${Math.max(1, Math.round(bytes / 1024))} KB` : "0 KB";
}

const NUMBER = new Intl.NumberFormat();
const DATE = new Intl.DateTimeFormat(undefined, { year: "numeric", month: "short", day: "numeric" });
const MONTH = new Intl.DateTimeFormat(undefined, { year: "numeric", month: "long" });

function parseAuthor(author) {
  const match = author.match(/^\s*"?([^"<]*?)"?\s*<([^>]+)>/);
  return match ? { name: match[1].trim(), email: match[2].trim().toLowerCase() } : { name: "", email: author.trim().toLowerCase() };
}

function folderOf(row) {
  return data.folders[row[1]];
}

function isSpecial(folder) {
  return SPECIAL.includes(folder.kind);
}

// What a sender is: a person, or an automatic sender with the service it
// belongs to (its domain). Cached: asked for every row on every filter.
const senders = new Map();

function senderOf(author) {
  let sender = senders.get(author);
  if (!sender) {
    const { name, email } = parseAuthor(author);
    const domain = email.split("@")[1] || "";
    const automated = AUTOMATED_ADDRESS.test(email) || AUTOMATED_DOMAIN.test(domain) || AUTOMATED_NAME.test(name);
    const base = baseDomain(domain);
    const service = automated ? SERVICE_NAMES[base.split(".")[0]] || capitalize(base.split(".")[0] || email) : "";
    sender = { name, email, domain: base, automated, service };
    senders.set(author, sender);
  }
  return sender;
}

// "mail.company.com.vn" → "company.com.vn"; "x.atlassian.net" → "atlassian.net".
function baseDomain(domain) {
  const labels = domain.split(".").filter(Boolean);
  const country = labels.length > 2 && labels.at(-1).length === 2 && labels.at(-2).length <= 3;
  return labels.slice(country ? -3 : -2).join(".");
}

function capitalize(text) {
  return text.charAt(0).toUpperCase() + text.slice(1);
}

function isAutomated(row) {
  return senderOf(row[4]).automated;
}

// How the From column shows a sender: the service first for automatic
// mail, with the person it is about.
function fromText(row) {
  const sender = senderOf(row[4]);
  if (sender.automated) {
    const about = sender.name && sender.name.toLowerCase() !== sender.service.toLowerCase() ? ` · ${sender.name}` : "";
    return `🤖 ${sender.service}${about}`;
  }
  return sender.name || sender.email;
}

/* ------------------------------------------------------------------ */
/* Filters                                                             */
/* ------------------------------------------------------------------ */

function defaultFolders() {
  return new Set(data.folders.map((f, i) => (isSpecial(f) ? -1 : i)).filter((i) => i >= 0));
}

function freshFilters() {
  return {
    folders: defaultFolders(),
    before: null,
    after: null,
    from: "",
    subject: "",
    minSize: 0,
    attachment: "any",
    state: "any",
    automated: false,
    protectFlagged: true,
  };
}

function matches(row, f) {
  const flags = row[6];
  if (!f.folders.has(row[1])) {
    return false;
  }
  // Pinned mail is never cleaned up, whatever the filters, until unpinned.
  if (flags & PINNED || (f.protectFlagged && flags & FLAGGED)) {
    return false;
  }
  if ((f.before && row[2] >= f.before) || (f.after && row[2] < f.after)) {
    return false;
  }
  if (f.minSize && row[3] < f.minSize) {
    return false;
  }
  if ((f.attachment === "with" && !(flags & ATTACHMENT)) || (f.attachment === "without" && flags & ATTACHMENT)) {
    return false;
  }
  if ((f.state === "unread" && flags & READ) || (f.state === "read" && !(flags & READ))) {
    return false;
  }
  if (f.from && !fold(row[4]).includes(fold(f.from))) {
    return false;
  }
  if (f.subject && !fold(row[5]).includes(fold(f.subject))) {
    return false;
  }
  if (f.automated && !isAutomated(row)) {
    return false;
  }
  return true;
}

function total(rows) {
  return { count: rows.length, size: rows.reduce((sum, row) => sum + row[3], 0) };
}

function mails(count) {
  return `${NUMBER.format(count)} mail${count === 1 ? "" : "s"}`;
}

function describe({ count, size }) {
  return `${mails(count)} · ${formatSize(size)}`;
}

/* ------------------------------------------------------------------ */
/* Quick clean                                                         */
/* ------------------------------------------------------------------ */

const AGES = [
  [3, "3 months"],
  [6, "6 months"],
  [12, "1 year"],
  [24, "2 years"],
];

function monthsAgo(months) {
  const date = new Date();
  date.setMonth(date.getMonth() - months);
  return date.getTime();
}

function folderIndexes(kind) {
  return new Set(data.folders.map((f, i) => (f.kind === kind ? i : -1)).filter((i) => i >= 0));
}

// Each card: the filters it sets (on top of fresh ones).
function presets() {
  return [
    {
      key: "automated",
      title: "🤖 Notifications & no-reply — start here",
      text: "Automatic mail: Bitbucket, Jira, SharePoint, no-reply, newsletters… The safest to clean first.",
      set: { automated: true },
    },
    {
      key: "old",
      title: "Older than…",
      text: "Mail you are unlikely to look at again.",
      ages: AGES.map(([months, label]) => ({ key: `old-${months}`, label, set: { before: monthsAgo(months) } })),
    },
    { key: "large", title: "Large mail (over 5 MB)", text: "Most of the room, in a few mails.", set: { minSize: 5 * 1024 * 1024 } },
    { key: "unread", title: "Unread for over a month", text: "Never opened, and old.", set: { state: "unread", before: Date.now() - 30 * DAY } },
    { key: "junk", title: "Junk", text: "Mail Thunderbird or the server marked as spam.", set: { folders: folderIndexes("junk") } },
    {
      key: "trash",
      title: "Deleted Items",
      text: "Still counts toward the mailbox size until emptied.",
      set: { folders: folderIndexes("trash"), protectFlagged: false },
    },
  ];
}

function applyPreset(key, set) {
  filters = { ...freshFilters(), ...set };
  activePreset = key;
  unticked = new Set();
  syncControls();
  render();
  $("review").scrollIntoView({ behavior: "smooth" });
}

function renderCards() {
  const box = $("cards");
  box.replaceChildren();
  for (const preset of presets()) {
    if (preset.ages) {
      const chips = preset.ages.map((age) => {
        const stats = total(data.rows.filter((row) => matches(row, { ...freshFilters(), ...age.set })));
        return el("button", {
          class: `chip${activePreset === age.key ? " active" : ""}`,
          title: describe(stats),
          text: `${age.label}: ${NUMBER.format(stats.count)} · ${formatSize(stats.size)}`,
          onclick: (event) => {
            event.stopPropagation();
            applyPreset(age.key, age.set);
          },
        });
      });
      box.append(
        el(
          "div",
          { class: `preset${String(activePreset).startsWith("old-") ? " active" : ""}` },
          el("span", { class: "title", text: preset.title }),
          el("span", { class: "meta", text: preset.text }),
          el("div", { class: "chips" }, ...chips)
        )
      );
      continue;
    }
    const stats = total(data.rows.filter((row) => matches(row, { ...freshFilters(), ...preset.set })));
    if (["junk", "trash"].includes(preset.key) && !preset.set.folders.size) {
      continue; // the account has no such folder
    }
    box.append(
      el(
        "button",
        { class: `preset${activePreset === preset.key ? " active" : ""}`, onclick: () => applyPreset(preset.key, preset.set) },
        el("span", { class: "title", text: preset.title }),
        el("span", { class: "meta", text: preset.text }),
        el("span", { class: "meta", text: describe(stats) })
      )
    );
  }
}

/* ------------------------------------------------------------------ */
/* Top senders                                                         */
/* ------------------------------------------------------------------ */

// Automatic senders are counted per service (Bitbucket sends from
// notifications-noreply@, pullrequests-reply@, commits-noreply@…), people
// per address.
function renderSenders() {
  const groups = new Map();
  for (const row of data.rows) {
    if (isSpecial(folderOf(row)) || row[6] & PINNED) {
      continue;
    }
    const sender = senderOf(row[4]);
    const key = sender.automated ? `auto:${sender.domain || sender.email}` : sender.email;
    const entry = groups.get(key) || { sender, addresses: new Set(), count: 0, size: 0, last: 0 };
    entry.addresses.add(sender.email);
    entry.count++;
    entry.size += row[3];
    entry.last = Math.max(entry.last, row[2]);
    if (!sender.automated && sender.name) {
      entry.sender = sender;
    }
    groups.set(key, entry);
  }
  const all = [...groups.values()].sort((a, b) => b.size - a.size);
  const table = (id, entries, label, filter) =>
    $(id).replaceChildren(
      el("tr", {}, el("th", { text: "Sender" }), el("th", { class: "num", text: "Mails" }), el("th", { class: "num", text: "Size" }), el("th", { class: "num", text: "Latest" }), el("th")),
      ...entries.map((entry) =>
        el(
          "tr",
          {},
          el("td", { class: "cut", title: [...entry.addresses].join(", "), text: label(entry) }),
          el("td", { class: "num", text: NUMBER.format(entry.count) }),
          el("td", { class: "num", text: formatSize(entry.size) }),
          el("td", { class: "num", text: DATE.format(new Date(entry.last)) }),
          el("td", { class: "num" }, el("button", { class: "link", text: "Review", onclick: () => applyPreset(`sender:${filter(entry)}`, { from: filter(entry) }) }))
        )
      )
    );
  const automatic = all.filter((entry) => entry.sender.automated).slice(0, 12);
  $("auto-senders-box").hidden = !automatic.length;
  table(
    "auto-senders",
    automatic,
    (entry) => {
      const addresses = [...entry.addresses];
      return `🤖 ${entry.sender.service} — ${addresses[0]}${addresses.length > 1 ? ` +${addresses.length - 1} more` : ""}`;
    },
    // All the service's addresses: by domain (also a subdomain).
    (entry) => (entry.addresses.size > 1 && entry.sender.domain ? entry.sender.domain : entry.sender.email)
  );
  table(
    "senders",
    all.filter((entry) => !entry.sender.automated).slice(0, 12),
    (entry) => (entry.sender.name ? `${entry.sender.name} <${entry.sender.email}>` : entry.sender.email),
    (entry) => entry.sender.email
  );
}

/* ------------------------------------------------------------------ */
/* Filter controls                                                     */
/* ------------------------------------------------------------------ */

function dateInput(ms) {
  if (!ms) {
    return "";
  }
  const d = new Date(ms);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function syncControls() {
  $("before").value = dateInput(filters.before);
  $("after").value = dateInput(filters.after);
  $("from").value = filters.from;
  $("subject").value = filters.subject;
  $("min-size").value = String(filters.minSize);
  $("attachment").value = filters.attachment;
  $("state").value = filters.state;
  $("protect-flagged").checked = filters.protectFlagged;
  for (const box of document.querySelectorAll("#folder-list input")) {
    box.checked = filters.folders.has(Number(box.value));
  }
  $("filter-label").textContent = filters.automated ? "Only notifications & newsletters." : "";
}

function readControls() {
  const date = (id) => ($(id).value ? new Date(`${$(id).value}T00:00:00`).getTime() : null);
  filters = {
    ...filters,
    before: date("before"),
    after: date("after"),
    from: $("from").value.trim(),
    subject: $("subject").value.trim(),
    minSize: Number($("min-size").value),
    attachment: $("attachment").value,
    state: $("state").value,
    protectFlagged: $("protect-flagged").checked,
    folders: new Set([...document.querySelectorAll("#folder-list input:checked")].map((box) => Number(box.value))),
  };
  activePreset = null;
  unticked = new Set();
  render();
}

function renderFolders() {
  const counts = new Map();
  for (const row of data.rows) {
    counts.set(row[1], (counts.get(row[1]) || 0) + 1);
  }
  $("folder-list").replaceChildren(
    ...data.folders.map((folder, index) =>
      el(
        "label",
        {},
        el("input", { type: "checkbox", value: String(index), onchange: readControls }),
        ` ${folder.path.replace(/^\//, "") || folder.name} `,
        el("span", { class: "hint", text: `(${NUMBER.format(counts.get(index) || 0)})` })
      )
    )
  );
}

/* ------------------------------------------------------------------ */
/* Results and actions                                                 */
/* ------------------------------------------------------------------ */

function matching() {
  const rows = data.rows.filter((row) => matches(row, filters));
  const sort = $("sort").value;
  const key = sort.replace("-", "") === "size" ? 3 : 2;
  const dir = sort.startsWith("-") ? -1 : 1;
  return rows.sort((a, b) => dir * (a[key] - b[key]));
}

function chosen() {
  return matching().filter((row) => !unticked.has(row[0]));
}

function renderResults() {
  const rows = matching();
  const picked = rows.filter((row) => !unticked.has(row[0]));
  const onlyTrash = filters.folders.size && [...filters.folders].every((i) => data.folders[i].kind === "trash");
  $("summary").textContent = `${describe(total(picked))} selected` + (picked.length !== rows.length ? ` (of ${NUMBER.format(rows.length)} found)` : "");
  const allBox = el("input", {
    type: "checkbox",
    title: "Select all",
    onchange: (event) => {
      unticked = event.target.checked ? new Set() : new Set(rows.map((row) => row[0]));
      renderResults();
    },
  });
  allBox.checked = picked.length === rows.length && rows.length > 0;
  allBox.indeterminate = picked.length > 0 && picked.length < rows.length;
  // Sorted by date: a heading per month, with a box for the whole month
  // (also its mails beyond the rows shown).
  const byDate = $("sort").value !== "-size";
  const months = new Map();
  if (byDate) {
    for (const row of rows) {
      const d = new Date(row[2]);
      const key = `${d.getFullYear()}-${d.getMonth()}`;
      const month = months.get(key) || { key, label: MONTH.format(d), rows: [] };
      month.rows.push(row);
      months.set(key, month);
    }
  }
  const monthRow = (month) => {
    const ticked = month.rows.filter((row) => !unticked.has(row[0]));
    const box = el("input", {
      type: "checkbox",
      title: `Select all of ${month.label}`,
      onchange: (event) => {
        for (const row of month.rows) {
          if (event.target.checked) {
            unticked.delete(row[0]);
          } else {
            unticked.add(row[0]);
          }
        }
        renderResults();
      },
    });
    box.checked = ticked.length === month.rows.length;
    box.indeterminate = ticked.length > 0 && ticked.length < month.rows.length;
    const stats = total(month.rows);
    return el(
      "tr",
      { class: "month" },
      el("td", {}, box),
      el(
        "td",
        { colspan: "5" },
        el("span", { class: "month-name", text: month.label }),
        el("span", {
          class: "hint",
          text: ` — ${describe(stats)}` + (ticked.length === month.rows.length ? "" : ` · ${NUMBER.format(ticked.length)} ticked`),
        })
      )
    );
  };
  let lastMonth = null;
  $("results").replaceChildren(
    el("tr", {}, el("th", {}, allBox), el("th", { text: "Received" }), el("th", { text: "From" }), el("th", { text: "Subject" }), el("th", { text: "Folder" }), el("th", { class: "num", text: "Size" })),
    ...rows.slice(0, SHOWN_ROWS).flatMap((row) => {
      const out = [];
      if (byDate) {
        const d = new Date(row[2]);
        const key = `${d.getFullYear()}-${d.getMonth()}`;
        if (key !== lastMonth) {
          lastMonth = key;
          out.push(monthRow(months.get(key)));
        }
      }
      const box = el("input", {
        type: "checkbox",
        onchange: (event) => {
          if (event.target.checked) {
            unticked.delete(row[0]);
          } else {
            unticked.add(row[0]);
          }
          renderResults();
        },
      });
      box.checked = !unticked.has(row[0]);
      out.push(
        el(
          "tr",
          { class: box.checked ? "" : "off" },
          el("td", {}, box),
          el("td", { class: "nowrap", text: DATE.format(new Date(row[2])) }),
          el("td", { class: "cut", title: row[4], text: fromText(row) }),
          el("td", { class: "cut", title: row[5], text: `${row[6] & ATTACHMENT ? "📎 " : ""}${row[5] || "(no subject)"}` }),
          el("td", { class: "nowrap", text: folderOf(row).name }),
          el("td", { class: "num", text: formatSize(row[3]) })
        )
      );
      return out;
    })
  );
  $("more").textContent =
    rows.length > SHOWN_ROWS
      ? `Showing the first ${SHOWN_ROWS} of ${NUMBER.format(rows.length)}. The buttons below act on all ${NUMBER.format(picked.length)} ticked mails, also those not shown.`
      : "";
  // Say on every button exactly how much it acts on: the ticked mails of
  // the list above, nothing else.
  const count = mails(picked.length);
  $("act-label").textContent = picked.length
    ? `Act on the ${count} (${formatSize(total(picked).size)}) ticked in the list above — nothing else is touched`
    : "Tick mails in the list above to act on them";
  $("trash").hidden = onlyTrash;
  $("archive").hidden = onlyTrash;
  $("trash").textContent = `🗑 Move ${count} to Deleted Items`;
  $("archive").textContent = `🗄 Archive ${count}`;
  $("delete").textContent = onlyTrash ? `Empty Deleted Items (${count})…` : `Delete ${count} permanently…`;
  for (const id of ["trash", "archive", "delete"]) {
    $(id).disabled = !picked.length;
  }
}

// Mailbox limits typed by the user, per account (Microsoft 365 does not
// give Thunderbird the quota of a mailbox): { accountId: bytes }.
let limits = {};

async function loadLimits() {
  try {
    limits = (await messenger.storage.local.get("cleanupLimits")).cleanupLimits || {};
  } catch (error) {
    limits = {};
  }
}

// The largest limit accepted (1 TB; Microsoft 365 mailboxes are far smaller).
const MAX_LIMIT_GB = 1024;

function setLimit() {
  let gb = parseFloat($("limit").value);
  if (gb > MAX_LIMIT_GB) {
    gb = MAX_LIMIT_GB;
    $("limit").value = String(MAX_LIMIT_GB);
  }
  if (gb > 0) {
    limits[$("account").value] = Math.round(gb * 1024 ** 3);
  } else {
    delete limits[$("account").value];
  }
  messenger.storage.local.set({ cleanupLimits: limits }).catch(() => {});
  renderUsage();
}

function renderUsage() {
  const all = total(data.rows);
  const trash = total(data.rows.filter((row) => folderOf(row).kind === "trash"));
  const usage = data.usage;
  const serverLimit = usage && usage.limit;
  const limit = serverLimit || limits[$("account").value] || null;
  // The server's number when there is one, else the size of all the mail.
  const used = usage ? usage.used : all.size;
  $("limit-box").hidden = Boolean(serverLimit);
  if (!serverLimit && document.activeElement !== $("limit")) {
    $("limit").value = limit ? String(Math.round((limit / 1024 ** 3) * 100) / 100) : "";
  }
  let text = `${describe(all)} in this account`;
  let ratio = null;
  if (limit) {
    ratio = used / limit;
    text =
      `${formatSize(used)} of ${formatSize(limit)} used (${Math.round(ratio * 100)}%)` +
      (used < limit ? ` · ${formatSize(limit - used)} free` : " · full") +
      ` · ${NUMBER.format(all.count)} mails`;
  } else if (usage) {
    text = `${formatSize(usage.used)} on the server · ${describe(all)}`;
  }
  if (trash.count) {
    text += ` · Deleted Items: ${describe(trash)}`;
  }
  $("usage-text").textContent = text;
  $("usage-fill").style.width = `${Math.min(100, Math.round((ratio ?? 0) * 100))}%`;
  $("usage-fill").classList.toggle("full", ratio != null && ratio > 0.9);
  $("usage-fill").parentElement.hidden = ratio == null;
}

function render() {
  renderUsage();
  renderCards();
  renderResults();
}

function setStatus(text, isError = false) {
  $("status").textContent = text;
  $("status").classList.toggle("error", isError);
}

// The action running, while it runs: { verb, size }.
let working = null;

function showProgress(handled, count, note = "") {
  const percent = count ? Math.floor((handled * 100) / count) : 100;
  $("act-progress").hidden = false;
  $("act-progress-fill").style.width = `${percent}%`;
  setStatus(`${working.verb} ${NUMBER.format(handled)} of ${mails(count)} (${working.size}) — ${percent}%…${note ? ` ${note}` : ""}`);
}

async function stop() {
  if (!working) {
    return;
  }
  working.stopping = true;
  $("stop").disabled = true;
  setStatus(`${$("status").textContent.replace(/ Stopping….*$/, "")} Stopping…`);
  await messenger.runtime.sendMessage({ type: "cleanup-stop" });
}

async function act(action) {
  const rows = chosen();
  if (!rows.length) {
    return;
  }
  const what = describe(total(rows));
  if (action === "delete" && !confirm(`Delete ${what} permanently? This cannot be undone.`)) {
    return;
  }
  if (action !== "delete" && rows.length > 500 && !confirm(`${action === "archive" ? "Archive" : "Move to Deleted Items"}: ${what}?`)) {
    return;
  }
  for (const id of ["trash", "archive", "delete"]) {
    $(id).disabled = true;
  }
  working = { verb: action === "archive" ? "Archiving" : action === "delete" ? "Deleting" : "Moving to Deleted Items:", size: formatSize(total(rows).size) };
  showProgress(0, rows.length);
  $("stop").hidden = false;
  $("stop").disabled = false;
  try {
    const result = await messenger.runtime.sendMessage({ type: "cleanup-apply", ids: rows.map((row) => row[0]), action });
    // Keep the list in step without reading everything again: only what
    // was really done.
    const ids = new Set(result.doneIds);
    const trashIndex = data.folders.findIndex((f) => f.kind === "trash");
    if (action === "trash" && trashIndex >= 0) {
      for (const row of data.rows) {
        if (ids.has(row[0])) {
          row[1] = trashIndex;
        }
      }
    } else {
      data.rows = data.rows.filter((row) => !ids.has(row[0]));
    }
    unticked = new Set();
    const verb = action === "archive" ? "Archived" : action === "delete" ? "Deleted" : "Moved to Deleted Items:";
    const doneRows = rows.filter((row) => ids.has(row[0]));
    setStatus(
      `${verb} ${describe(total(doneRows))}.` +
        (action === "trash" && doneRows.length ? " Empty Deleted Items to free the space." : "") +
        (result.skipped ? ` ${NUMBER.format(result.skipped)} left alone (pinned meanwhile).` : "") +
        (result.stalled
          ? ` Stopped: Thunderbird got no answer from the server for ${mails(result.notDone)}. Restart Thunderbird, then Scan again and act on what is left.`
          : result.stopped
            ? ` Stopped: ${mails(result.notDone)} not done.`
            : result.notDone
              ? ` ${mails(result.notDone)} not done: the server did not confirm them. Scan again and try once more.`
              : "") +
        (result.failed.length ? ` Error: ${result.failed[0]}` : ""),
      Boolean(result.failed.length || result.notDone)
    );
    renderFolders();
    syncControls();
    renderSenders();
    render();
  } catch (error) {
    setStatus(`Failed: ${error.message || error}`, true);
    renderResults();
  } finally {
    working = null;
    $("act-progress").hidden = true;
    $("stop").hidden = true;
  }
}

/* ------------------------------------------------------------------ */
/* Start                                                               */
/* ------------------------------------------------------------------ */

async function scan() {
  $("rescan").disabled = true;
  $("usage-text").textContent = "Reading your mail…";
  try {
    data = await messenger.runtime.sendMessage({ type: "cleanup-scan", accountId: $("account").value });
    filters = freshFilters();
    activePreset = null;
    unticked = new Set();
    renderFolders();
    syncControls();
    renderSenders();
    render();
  } catch (error) {
    $("usage-text").textContent = `Could not read the mail: ${error.message || error}`;
  }
  $("rescan").disabled = false;
}

messenger.runtime.onMessage.addListener((message) => {
  if (message && message.type === "cleanup-progress") {
    $("progress").textContent = message.text;
  }
  if (message && message.type === "cleanup-apply-progress" && working) {
    showProgress(message.handled, message.total, working.stopping ? "Stopping…" : message.note);
  }
  return undefined;
});

async function init() {
  const accounts = await messenger.runtime.sendMessage({ type: "cleanup-accounts" });
  for (const account of accounts) {
    $("account").append(new Option(account.name, account.id));
  }
  await loadLimits();
  $("account").addEventListener("change", scan);
  $("limit").addEventListener("change", setLimit);
  $("rescan").addEventListener("click", scan);
  for (const id of ["before", "after", "min-size", "attachment", "state", "protect-flagged"]) {
    $(id).addEventListener("change", readControls);
  }
  for (const id of ["from", "subject"]) {
    $(id).addEventListener("input", readControls);
  }
  $("sort").addEventListener("change", renderResults);
  $("reset").addEventListener("click", () => {
    filters = freshFilters();
    activePreset = null;
    unticked = new Set();
    syncControls();
    render();
  });
  $("folders-all").addEventListener("click", () => {
    filters.folders = new Set(data.folders.map((_, i) => i));
    syncControls();
    readControls();
  });
  $("folders-none").addEventListener("click", () => {
    filters.folders = new Set();
    syncControls();
    readControls();
  });
  $("trash").addEventListener("click", () => act("trash"));
  $("archive").addEventListener("click", () => act("archive"));
  $("delete").addEventListener("click", () => act("delete"));
  $("stop").addEventListener("click", () => stop());
  if (accounts.length) {
    await scan();
  }
}

init();
