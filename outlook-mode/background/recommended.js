"use strict";

// Thunderbird settings the add-on's features expect (experiment
// outlookSync.applyRecommended): HTML compose for @mention links, the reply
// above the quote like Outlook, and the Cards View for the pin button.
// Applied once per device (when the add-on is installed or updated to a
// version that adds a setting here), so a later change by the user stays.
// Preferences has a button to apply them again.

const RECOMMENDED_KEY = "recommendedSettingsVersion";
const RECOMMENDED_VERSION = 1; // raise when a setting is added above

const RECOMMENDED_LABELS = {
  html: "HTML compose",
  replyOnTop: "reply above the quote",
  cardsView: "Cards View",
};

function describeRecommended(changed) {
  return [...new Set(changed.map((c) => RECOMMENDED_LABELS[c.what] || c.what))].join(", ");
}

async function applyRecommendedSettings() {
  const { changed } = await messenger.outlookSync.applyRecommended();
  return { ok: true, changed: changed.length, summary: describeRecommended(changed) };
}

(async () => {
  const stored = await messenger.storage.local.get(RECOMMENDED_KEY);
  if ((stored[RECOMMENDED_KEY] || 0) >= RECOMMENDED_VERSION) {
    return;
  }
  const result = await applyRecommendedSettings();
  await messenger.storage.local.set({ [RECOMMENDED_KEY]: RECOMMENDED_VERSION });
  if (result.changed) {
    notify(
      "Outlook Mode",
      `Thunderbird is now set up like Outlook: ${result.summary}. You can change this back in Thunderbird's settings.`
    );
  }
})().catch((error) => logError("recommended settings", error));

handlers["apply-recommended"] = fromOwnPage(() => applyRecommendedSettings());
