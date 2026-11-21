"use strict";

// Thunderbird's message index (Gloda) for Microsoft Graph accounts.
//
// Thunderbird Conversations shows the messages of a conversation that are in
// other folders (your own replies, in Sent Items) through Gloda. Gloda indexes
// new mail when Thunderbird reports it as classified, which the Graph sync
// does not do for the mail it brings in; so a reply sent from Thunderbird was
// missing from the conversation in the reading pane until Gloda happened to
// sweep its folder. The experiment outlookServer (indexGraphMail) gives such
// mail to Gloda's indexer; this file runs it: at startup and on Refresh for
// every Graph folder, and every 2 minutes for the folders whose message count
// changed.

const GRAPH_INDEX_ALARM = "graph-index";
const GRAPH_INDEX_MINUTES = 2;

async function indexGraphMail(force = false) {
  if (!messenger.outlookServer || !messenger.outlookServer.indexGraphMail) {
    return 0;
  }
  if (!(await graphAccountIds()).length) {
    return 0;
  }
  return messenger.outlookServer.indexGraphMail(force);
}

messenger.alarms.create(GRAPH_INDEX_ALARM, { periodInMinutes: GRAPH_INDEX_MINUTES });
messenger.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === GRAPH_INDEX_ALARM) {
    indexGraphMail().catch((error) => logError("index", error));
  }
});

serverBackHooks.push(() => indexGraphMail(true));

// After Thunderbird has loaded its folders.
setTimeout(() => {
  indexGraphMail(true).catch((error) => logError("startup index", error));
}, 15000);
