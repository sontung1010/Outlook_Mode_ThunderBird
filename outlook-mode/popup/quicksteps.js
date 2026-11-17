"use strict";

async function currentTab() {
  const [tab] = await messenger.tabs.query({ active: true, currentWindow: true });
  return tab;
}

function setStatus(text, isError = false) {
  const status = document.getElementById("status");
  status.textContent = text;
  status.classList.toggle("error", isError);
}

async function run(message, doneText) {
  const tab = await currentTab();
  if (!tab) {
    setStatus("No mail tab found.", true);
    return;
  }
  try {
    const result = await messenger.runtime.sendMessage({ ...message, tabId: tab.id });
    if (result && (result.ran === false || result.pinned === null)) {
      setStatus("Select one or more messages first.", true);
      return;
    }
    if (result && result.ok === false) {
      setStatus(result.error || "Failed.", true);
      return;
    }
    setStatus(doneText(result));
    setTimeout(() => window.close(), 600);
  } catch (error) {
    setStatus(`Failed: ${error.message || error}`, true);
  }
}

function openOptions(event, section = "") {
  event.preventDefault();
  messenger.runtime.sendMessage({ type: "open-settings", section }).catch(() => messenger.runtime.openOptionsPage());
  window.close();
}

async function init() {
  const steps = await messenger.runtime.sendMessage({ type: "get-steps" });
  const box = document.getElementById("steps");
  if (!steps.length) {
    box.textContent = "No Quick Steps yet.";
  }
  steps.forEach((step, index) => {
    const button = document.createElement("button");
    button.className = "row";
    button.textContent = step.name || `Step ${index + 1}`;
    if (index < 5) {
      const key = document.createElement("span");
      key.className = "key";
      key.textContent = `Ctrl+Shift+${index + 1}`;
      button.appendChild(key);
    }
    button.addEventListener("click", () =>
      run({ type: "run-step", index }, (r) => `Done (${r.count} message${r.count > 1 ? "s" : ""}).`)
    );
    box.appendChild(button);
  });

  document.getElementById("pin").addEventListener("click", () =>
    run({ type: "toggle-pin" }, (r) => (r.pinned ? "Pinned." : "Unpinned."))
  );
  document.getElementById("mentions-show").addEventListener("click", () =>
    run({ type: "mentions-show", show: true }, () => "Showing mail that @mentions you.")
  );
  document.getElementById("mentions-clear").addEventListener("click", () =>
    run({ type: "mentions-show", show: false }, () => "Filter cleared.")
  );
  document.getElementById("manage").addEventListener("click", openOptions);
  document.getElementById("open-settings").addEventListener("click", (event) => openOptions(event));
  document.getElementById("open-signature").addEventListener("click", (event) => openOptions(event, "signature"));
  document.getElementById("open-cleanup").addEventListener("click", (event) => {
    event.preventDefault();
    messenger.runtime.sendMessage({ type: "cleanup-open" }).catch(() => {});
    window.close();
  });

  const scheduled = await messenger.runtime.sendMessage({ type: "schedule-list" });
  const box2 = document.getElementById("scheduled");
  const pending = scheduled.filter((s) => s.status === "pending");
  const failed = scheduled.filter((s) => s.status === "failed");
  if (!scheduled.length) {
    box2.textContent = "Nothing scheduled. Use “Send later” in the compose window.";
  } else {
    const next = pending
      .map((s) => new Date(s.sendAt))
      .sort((a, b) => a - b)[0];
    box2.textContent =
      `${pending.length} waiting` +
      (next ? ` · next ${next.toLocaleString()}` : "") +
      (failed.length ? ` · ${failed.length} failed` : "");
    if (failed.length) {
      box2.classList.add("error");
    }
  }
}

init();
