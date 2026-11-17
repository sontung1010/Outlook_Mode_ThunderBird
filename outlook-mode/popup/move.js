"use strict";

let folders = [];
let recent = [];
let shown = [];
let active = 0;

function fold(text) {
  return (text || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D")
    .toLowerCase();
}

function render() {
  const list = document.getElementById("list");
  list.textContent = "";
  const words = fold(document.getElementById("filter").value).split(/\s+/).filter(Boolean);
  const matches = folders.filter((f) => words.every((w) => fold(`${f.path} ${f.account}`).includes(w)));
  const recentSet = new Set(recent);
  const recentFolders = words.length
    ? []
    : recent.map((id) => folders.find((f) => f.id === id)).filter(Boolean);
  shown = [...recentFolders, ...matches.filter((f) => words.length || !recentSet.has(f.id))];
  active = Math.min(active, Math.max(0, shown.length - 1));
  shown.forEach((folder, index) => {
    if (index === 0 && recentFolders.length) {
      list.appendChild(section("Recent"));
    }
    if (index === recentFolders.length && recentFolders.length) {
      list.appendChild(section("All folders"));
    }
    const row = document.createElement("div");
    row.className = `folder${index === active ? " active" : ""}`;
    row.textContent = folder.path;
    const account = document.createElement("span");
    account.className = "account";
    account.textContent = folder.account;
    row.appendChild(account);
    row.addEventListener("click", () => move(folder));
    row.addEventListener("mouseenter", () => {
      active = index;
      for (const el of list.querySelectorAll(".folder")) {
        el.classList.remove("active");
      }
      row.classList.add("active");
    });
    list.appendChild(row);
  });
  const current = list.querySelectorAll(".folder")[active];
  if (current) {
    current.scrollIntoView({ block: "nearest" });
  }
}

function section(text) {
  const el = document.createElement("div");
  el.className = "section";
  el.textContent = text;
  return el;
}

async function move(folder) {
  const status = document.getElementById("status");
  try {
    const result = await messenger.runtime.sendMessage({ type: "move-to", folderId: folder.id });
    status.textContent = `Moved ${result.count} message${result.count > 1 ? "s" : ""} to ${folder.name}.`;
    setTimeout(() => window.close(), 400);
  } catch (error) {
    status.textContent = String(error.message || error);
    status.classList.add("error");
  }
}

async function init() {
  const data = await messenger.runtime.sendMessage({ type: "move-data" });
  folders = data.folders;
  recent = data.recent;
  document.getElementById("heading").textContent =
    `Move ${data.count} message${data.count === 1 ? "" : "s"} to folder`;
  const filter = document.getElementById("filter");
  filter.addEventListener("input", () => {
    active = 0;
    render();
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "ArrowDown") {
      active = Math.min(active + 1, shown.length - 1);
      render();
    } else if (event.key === "ArrowUp") {
      active = Math.max(active - 1, 0);
      render();
    } else if (event.key === "Enter") {
      if (shown[active]) {
        move(shown[active]);
      }
    } else if (event.key === "Escape") {
      window.close();
    } else {
      return;
    }
    event.preventDefault();
  });
  render();
  filter.focus();
}

init();
