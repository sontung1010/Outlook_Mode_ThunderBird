"use strict";

function at(date, hours, minutes = 0) {
  const d = new Date(date);
  d.setHours(hours, minutes, 0, 0);
  return d;
}

// Outlook-like suggestions.
function suggestions() {
  const now = new Date();
  const list = [];
  const inOneHour = new Date(now.getTime() + 60 * 60 * 1000);
  inOneHour.setSeconds(0, 0);
  list.push({ label: "In 1 hour", when: inOneHour });
  const thisAfternoon = at(now, 13);
  if (thisAfternoon > new Date(now.getTime() + 15 * 60 * 1000)) {
    list.push({ label: "This afternoon", when: thisAfternoon });
  }
  const tomorrow = new Date(now);
  tomorrow.setDate(now.getDate() + 1);
  list.push({ label: "Tomorrow morning", when: at(tomorrow, 8) });
  list.push({ label: "Tomorrow afternoon", when: at(tomorrow, 13) });
  const monday = new Date(now);
  monday.setDate(now.getDate() + ((8 - now.getDay()) % 7 || 7));
  list.push({ label: "Monday morning", when: at(monday, 8) });
  return list;
}

function toLocalInput(date) {
  const pad = (n) => String(n).padStart(2, "0");
  return (
    `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}` +
    `T${pad(date.getHours())}:${pad(date.getMinutes())}`
  );
}

function setStatus(text, isError = false) {
  const status = document.getElementById("status");
  status.textContent = text;
  status.classList.toggle("error", isError);
}

async function schedule(when) {
  // In the future and within a year (the background checks it again).
  const time = when instanceof Date ? when.getTime() : NaN;
  if (!(time > Date.now() && time < Date.now() + 366 * 864e5)) {
    setStatus("Choose a time in the future, within a year.", true);
    return;
  }
  const [tab] = await messenger.tabs.query({ active: true, currentWindow: true });
  if (!tab || tab.type !== "messageCompose") {
    setStatus("Open this from a compose window.", true);
    return;
  }
  for (const button of document.querySelectorAll("button")) {
    button.disabled = true;
  }
  setStatus("Saving…");
  try {
    await messenger.runtime.sendMessage({
      type: "schedule-send",
      tabId: tab.id,
      sendAt: when.toISOString(),
    });
    setStatus(`Scheduled for ${when.toLocaleString()}.`);
    setTimeout(() => window.close(), 900);
  } catch (error) {
    setStatus(String(error.message || error), true);
    for (const button of document.querySelectorAll("button")) {
      button.disabled = false;
    }
  }
}

function init() {
  const box = document.getElementById("quick");
  for (const { label, when } of suggestions()) {
    const button = document.createElement("button");
    button.className = "row";
    button.textContent = label;
    const key = document.createElement("span");
    key.className = "key";
    key.textContent = when.toLocaleString(undefined, {
      weekday: "short",
      hour: "2-digit",
      minute: "2-digit",
    });
    button.appendChild(key);
    button.addEventListener("click", () => schedule(when));
    box.appendChild(button);
  }
  const custom = document.getElementById("custom");
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  custom.value = toLocalInput(at(tomorrow, 8));
  custom.min = toLocalInput(new Date());
  document.getElementById("schedule").addEventListener("click", () => {
    const when = new Date(custom.value);
    if (Number.isNaN(when.getTime())) {
      setStatus("Pick a date and time.", true);
      return;
    }
    schedule(when);
  });
}

init();
