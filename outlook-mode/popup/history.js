"use strict";

function formatDate(iso) {
  const date = new Date(iso);
  return date.toLocaleString(undefined, {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
}

async function init() {
  const heading = document.getElementById("heading");
  const sub = document.getElementById("sub");
  const list = document.getElementById("list");
  const [tab] = await messenger.tabs.query({ active: true, currentWindow: true });
  let result;
  try {
    result = await messenger.runtime.sendMessage({ type: "history", tabId: tab.id });
  } catch (error) {
    heading.textContent = "Could not load history.";
    sub.textContent = String(error.message || error);
    return;
  }
  if (result.error) {
    heading.textContent = result.error;
    return;
  }
  heading.textContent = `Mail with ${result.person.name}`;
  sub.textContent = `${result.person.email} · last ${result.days} days · ${result.messages.length} shown`;
  if (!result.messages.length) {
    list.textContent = "No other messages found.";
    return;
  }
  for (const m of result.messages) {
    const item = document.createElement("button");
    item.className = `history-item${m.current ? " current" : ""}`;
    const subject = document.createElement("div");
    subject.className = "subject";
    const dir = document.createElement("span");
    dir.className = `dir ${m.fromMe ? "out" : "in"}`;
    dir.textContent = m.fromMe ? "Sent" : "From";
    subject.append(dir, document.createTextNode(` ${m.subject}`));
    const meta = document.createElement("div");
    meta.className = "meta";
    meta.textContent = `${formatDate(m.date)} · ${m.folder}${m.current ? " · (open now)" : ""}`;
    item.append(subject, meta);
    item.addEventListener("click", async () => {
      await messenger.runtime.sendMessage({ type: "open-message", messageId: m.id });
      window.close();
    });
    list.appendChild(item);
  }
}

init();
