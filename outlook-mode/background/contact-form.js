"use strict";

// New contacts in the company's format in the address book (experiment
// outlookContacts, page script contacts/contact-form.js): email field first,
// the name of a known address filled in, and a "Company – department" field
// that suggests the codes already used. This file keeps it configured.

async function configureContactForm() {
  const settings = await getSettings();
  try {
    await messenger.outlookContacts.configure({ enabled: settings.contactCompanyFormat !== false });
  } catch (error) {
    logError("contact form", error);
  }
}

messenger.storage.onChanged.addListener((changes) => {
  if (changes.settings) {
    configureContactForm();
  }
});

configureContactForm();
