"use strict";

// Loaded first, before any other background code, so that an error elsewhere
// can never stop the compose script (@mentions, reply header, Undo Send banner)
// from being registered.

const COMPOSE_SCRIPT = "/compose/compose.js";

// Status shown on the options page (Diagnostics).
const bootStatus = {
  composeScriptRegistered: false,
  composeScriptError: "",
  injectedIntoOpenTabs: 0,
  injectErrors: [],
};

const composeScriptReady = (async () => {
  try {
    await messenger.composeScripts.register({ js: [{ file: COMPOSE_SCRIPT }] });
    bootStatus.composeScriptRegistered = true;
  } catch (error) {
    bootStatus.composeScriptError = String(error && error.message ? error.message : error);
    console.error("[outlook-mode] compose script registration failed", error);
  }
  // Registered compose scripts only reach windows opened from now on. Compose
  // windows that were already open (for example while the add-on was being
  // installed or updated) get the script injected directly. The script
  // guards against running twice.
  try {
    const tabs = await messenger.tabs.query({ type: "messageCompose" });
    for (const tab of tabs) {
      try {
        await messenger.tabs.executeScript(tab.id, { file: COMPOSE_SCRIPT });
        bootStatus.injectedIntoOpenTabs++;
      } catch (error) {
        bootStatus.injectErrors.push(String(error && error.message ? error.message : error));
      }
    }
  } catch (error) {
    bootStatus.injectErrors.push(String(error && error.message ? error.message : error));
  }
})();
