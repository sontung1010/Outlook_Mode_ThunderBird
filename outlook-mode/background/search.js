"use strict";

// Outlook-style search (experiment outlookSearch): the search box at the
// top suggests people and mail while typing, and shows the results in the
// message list of the mail tab instead of opening a new tab.

// On or off is the user's choice in Preferences only: it is kept by the
// experiment (a Thunderbird preference), not read from storage.
function configureSearch() {
  return messenger.outlookSearch.configure({});
}

configureSearch().catch((error) => logError("search", error));

handlers["search-status"] = fromPages(["options/options.html"], () => messenger.outlookSearch.status());
handlers["search-set"] = fromPages(["options/options.html"], (message) => {
  if (typeof message.enabled !== "boolean") {
    throw userError("Invalid value.");
  }
  return messenger.outlookSearch.configure({ enabled: message.enabled });
});
