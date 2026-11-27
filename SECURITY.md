# Outlook Mode add-on: security notes

What the **Outlook Mode** Thunderbird add-on (version **3.9.44**) does with the network, your credentials and your data. Every statement below was checked against the package.

## 1. Identification

| Item | Value |
| --- | --- |
| Name | Outlook Mode |
| Add-on ID | `outlook-mention@tung.local` |
| Version | 3.9.44 |
| File | `outlook-mode.xpi` (227,480 bytes) |
| SHA-256 | `28a32ac98413937e64722f76d8c4a14830750b3b9749b303cec60fb8dc2749a6` |
| Source | The package is a plain zip of the add-on's source files and its `LICENSE`. Unzip it to read the code. |
| License | MIT (see `LICENSE`). |
| Code | About 18,700 lines of plain, unminified JavaScript, HTML and CSS written for this add-on. No bundled third-party libraries. |
| Distribution | Not published on addons.thunderbird.net and **not signed by Mozilla**. It is installed by hand: Add-ons Manager → Install Add-on From File. |
| Target | Thunderbird 140 or later (tested on 159 and 160, Linux), with a Microsoft 365 mailbox set up as a *Microsoft Graph* account. Minimum Thunderbird version declared: 128. |

## 2. Purpose

Microsoft 365 users who move from Outlook to Thunderbird lose many everyday functions. This add-on brings Outlook-style working back inside Thunderbird:

- @mentions in the message body, and tagging of mail that mentions you;
- Undo Send (a 10-second hold) and Schedule send;
- Pin, with "pinned first" sorting, synced with Outlook's own pin;
- Outlook-style reading pane and inline reply;
- Outlook reply/forward header;
- Quick Steps, Outlook keyboard shortcuts and "Move to folder";
- Outlook-style search;
- Mailbox cleanup;
- Settings sync between the user's devices (off by default);
- new contacts in the company format "Full Name (Company-Department)", from data already on the computer (section 4a);
- one notice when Microsoft's server is failing and one when it is back, instead of Thunderbird's repeated error notices (section 4b).

## 3. Network connections

The add-on itself contacts exactly **one** host. It sends no telemetry and uses no analytics or other third-party services.

| Host | When | What |
| --- | --- | --- |
| `graph.microsoft.com` (HTTPS, `/v1.0`) | Only for Microsoft Graph mail accounts | Calls the user's own mailbox (`/me/...`) with Thunderbird's existing sign-in. See section 4. |

The only network call in the code is one `fetch()`, in `background/graph.js`; it refuses any URL outside `https://graph.microsoft.com/v1.0/`. The manifest's only host permission is `https://graph.microsoft.com/*`.

A test on Thunderbird 159 recorded every HTTP request with Thunderbird's `http-on-modify-request` observer. With the default settings and no Graph account, the add-on made no external request at all.

## 4. Microsoft Graph usage and credentials

**Token.** `getGraphToken` in `experiments/outlookPin` asks Thunderbird's own `OAuth2Module` for the access token of the user's account. It never shows a sign-in window and never requests new scopes or consent; it uses the token Thunderbird already holds for mail. The token:

- stays in memory;
- is sent only in the `Authorization` header to `graph.microsoft.com`;
- is never stored, logged or synced.

The add-on reads no passwords and stores none.

**Endpoints called.** All calls go to the signed-in user's own mailbox:

| Feature | Default | Calls |
| --- | --- | --- |
| Pin sync | **On** for Graph accounts | `GET /me/messages` (`id`, `subject`, `receivedDateTime`, two pin properties), and `PATCH /me/messages/{id}` to set Outlook's pin: the MAPI extended properties `0x0F01` and `0x0F02`, set to a date in the year 4500, which is how Outlook marks a pinned item. No other message field is changed. |
| Mailbox cleanup page | Opened by the user | `GET /me/mailFolders` (`displayName`, `sizeInBytes`) to show folder sizes. Deleting is done through Thunderbird's normal message API, only after the user confirms. |
| One copy of sent mail (3.9.34) | **On** for Graph accounts | `GET /me/mailFolders/sentitems/messages` (`id`, `internetMessageId`, `isDraft`, `createdDateTime` only; after the first check, only messages created since the last one), and `POST /me/messages/{id}/move` to Deleted Items for a second copy of a sent mail. See section 4c. |
| Settings sync | **Off**; the user must turn it on | Creates a hidden folder `Outlook Mode Sync` (`POST /me/mailFolders`, `isHidden: true`) and stores one message there with a gzip JSON attachment (`POST /me/mailFolders/{id}/messages`). Reads it back (`GET .../messages`, `GET /me/messages/{id}/attachments`). Older copies of that sync message are removed with `POST /me/messages/{id}/permanentDelete`. |

## 4a. New contacts in the company format (3.9.28)

In many organizations, a person's display name is "Full Name (Company-Department)", for example "Phạm Minh Châu (Contoso-RND-HN)". In Thunderbird's address book, when the user creates or edits a contact, the experiment `outlookContacts` and its page script `contacts/contact-form.js`:

- put the Email Addresses field first; a new contact starts there;
- split a pasted "Name (Code) <address>" (as copied from a To/From line) into the address, the display name, Company and Department;
- for a new contact, look up the name of a typed address in the local address books and, if not there, in Thunderbird's local mail index (Gloda: names seen in mail headers), and fill an empty display name with it;
- add a "Company – department" field that suggests the codes already used by the contacts of the address books (counted when the form opens; those of the same mail domain first). A code sets Company, Department and the "(…)" after the display name.

**Security relevance**:

- Only data already on the computer is read: the address books (no LDAP directory) and the Gloda index.
- No network access, no new permission, no sign-in.
- Nothing is written by the add-on: the user saves the contact with Thunderbird's own Save.
- The form's fields are changed only as typing would (value plus an `input` event).
- The page script is loaded from the package (`chrome://` URL registered to the add-on's own `contacts/` folder).

(3.9.25 to 3.9.27 had a lookup in the Microsoft 365 directory with an extra sign-in for `User.ReadBasic.All` and `People.Read`. It needed a Microsoft 365 administrator's consent and was removed in 3.9.28, together with its use of the password manager.)

## 4b. Microsoft's server failing (3.9.26)

**What it does.** During an outage of Microsoft's server, Thunderbird shows *Server graph.microsoft.com responded with an error.* at every mail check. The experiment `outlookServer`:

- Sees Thunderbird's notices by wrapping `onAlert` of Thunderbird's own `alertHook` module (`resource:///modules/activity/alertHook.sys.mjs`). The original function is kept and called for every notice the add-on does not hold back, and it is put back when the add-on stops. Only notices of Microsoft Graph accounts are looked at.
- Reports a notice to the add-on when it is exactly Thunderbird's own text for an error answer of the server (`errorResponseError` of `messenger.properties`, in the user's language). "Could not connect" and time-outs are not counted, since they are often caused by the computer's own network.
- Checks the server by running Thunderbird's own folder list sync of the account (`IExchangeIncomingServer.syncFolderHierarchy`). This is the same request Thunderbird makes at each mail check, with Thunderbird's own sign-in. A notice during that run means it failed.
- Holds back Thunderbird's notices for one account while the add-on says its server is failing, and the notice of its own check. The add-on then shows two notices of its own (`notifications` permission): "Microsoft's server is failing" with Thunderbird's reason, and "Microsoft's server is back".

**While failing**:

- The check runs again after 30 s, 1 and 2 minutes, then every 5 minutes, or at once with ⟳ Refresh.
- The add-on's own Graph requests wait (at most 15 minutes, then they fail as before). Pin and settings syncs do not start; pin changes stay in the local queue.
- A request a person is looking at (the cleanup page's folder sizes) fails at once with a message instead of waiting.
- A Graph request answered 500/502/503/504 counts as a sign too. It is sent once more after the check, except a `POST`, which is never sent twice. A 429 (throttling) answer is waited out as Microsoft asks (`Retry-After`, at most 2 minutes), once.

**When it is back**: Thunderbird gets the account's new mail (`getNewMessages` on the account), and the put-off pin and settings syncs run.

**Security relevance**:

- No new host, permission, token or stored data. Only the notices of Graph accounts are affected, and only while the server is failing or being checked.
- The history of the last 20 changes ("failing", "back") is kept in memory for Diagnostics.

**What settings sync stores** (only when enabled; each part can be switched off):

- Thunderbird preferences;
- account and identity settings;
- message filters;
- local address books (including Collected Addresses);
- this add-on's own settings.

It never includes:

- filters that forward or reply automatically, or run another add-on's custom action (each filter file is checked with Thunderbird's own parser; such a file is refused);
- automatic Cc/Bcc and Reply-To of identities;
- protections: phishing and junk detection, remote content blocking, how HTML mail is shown;
- passwords, tokens or sign-in data;
- encryption keys or certificates;
- server host names or ports;
- file paths.

The exclusion rules are in `experiments/outlookSync/implementation.js` (`token|password|secret|oauth|...`, and the server and identity block lists). The data stays inside the user's own Microsoft 365 mailbox, in the user's own Microsoft 365 tenant.

## 4c. One copy of sent mail (3.9.34)

**The problem.** Thunderbird sends mail of a Microsoft Graph account with `POST /me/messages` and `/send`, and Exchange keeps the sent mail in Sent Items itself. Thunderbird then also filed its own copy (Fcc), uploaded as a new message, so every mail sent from Thunderbird was in Sent Items twice. Exchange stores the uploaded copy as an unsent draft, so Outlook (web, Windows, mobile) showed a "[Draft]" in the conversation although the mail had been sent.

**What it does** (Preferences → *Sending*, on by default):

- When a message is sent through a Graph outgoing server and its Fcc folder is on a Graph account, Thunderbird files no copy of its own (`compFields.fcc = "nocopy://"`). This covers the reading pane and the compose dock (experiment `outlookPane`, which builds the compose fields) and compose windows (a listener for the compose window's own `compose-send-message` event). Drafts, templates and mail sent through SMTP are not touched, and neither is an Fcc folder outside Graph accounts (e.g. Local Folders).
- The second copies already in Sent Items are moved to Deleted Items: messages with the same Message-ID as another one in Sent Items, created at most 10 minutes apart. The copy kept is the oldest one that is not a draft (Exchange's own). The check runs 20 s after startup, every 30 minutes, on ⟳ Refresh and when Microsoft's server is back. A notification says how many copies were moved. Moving to Deleted Items can be undone there.

**Security relevance**: no new host or permission. Only message ids, Message-IDs, the draft flag and creation times are read; no subject, body or recipient. The time of the last check per account is kept in `storage.local`.

## 4d. Message index for Graph accounts (3.9.35)

Thunderbird Conversations finds the messages of a conversation in other folders (your own replies in Sent Items) through Thunderbird's message index, Gloda. Gloda indexes new mail when Thunderbird reports it as classified, which Thunderbird's Microsoft Graph sync does not do, so mail it brings in (such as the Sent Items copy of a mail sent from Thunderbird) was missing from the conversation view. The experiment `outlookServer` (`indexGraphMail`):

- listens for mail added to folders of Graph accounts (`nsIMsgFolderListener.msgAdded`) and passes those messages to Gloda's own indexer (`GlodaMsgIndexer.indexMessages`);
- marks Graph folders whose message count changed (all Graph folders at startup and on ⟳ Refresh) as dirty in Gloda (`_ensureFolderDirty`), so Gloda's normal sweep indexes the messages it does not know yet.

**Security relevance**: no network access, no new permission. Nothing is read or written by the add-on itself: Gloda indexes the messages the way Thunderbird always does, into its own local database. Other accounts are not touched, and nothing is done when Gloda is turned off in Thunderbird. The listener is removed when the add-on stops.

## 5. Permissions

**Standard WebExtension permissions:**

| Permission | Why |
| --- | --- |
| `accountsRead` | List accounts and folders (folder picker, Inbox detection, own addresses). |
| `addressBooks` | Contact suggestions for @mention; adding people you write to into Collected Addresses. |
| `alarms` | Timers for Schedule send and periodic pin/settings sync. |
| `compose`, `compose.save`, `compose.send` | @mention and reply header in the editor; Undo Send (hold, then send); Schedule send (save as draft, send at the chosen time). |
| `menus` | Right-click menu entries (Pin, Move to folder, Quick Steps). |
| `messagesRead` | Read mail headers and bodies **locally**, to detect @mentions, show history with a sender, run search and cleanup. |
| `messagesUpdate`, `messagesTags`, `messagesTagsList` | Mark read/unread, the 📌 Pinned and @ Mentioned tags. |
| `messagesMove`, `messagesDelete` | Move, archive and delete, always started by the user (Quick Steps, Ctrl+D, Move to folder, cleanup). |
| `notifications` | Desktop notice when a scheduled mail is sent or fails. |
| `storage` | The add-on's own settings, stored locally. |
| `https://graph.microsoft.com/*` | Section 4. |

**Experiment APIs (privileged code).** The add-on ships six WebExtension *Experiments*. They run with Thunderbird's full privileges, which is why Thunderbird shows a full-access warning when the add-on is installed. They are needed because the standard API cannot change the message list, the menus or the reading pane.

| Experiment | What it does | Privileged operations |
| --- | --- | --- |
| `outlookPin` | Pin column and "Pinned first" sort; pin and refresh buttons in the message list; Outlook-style card view; Graph access token (section 4); thread helpers. | Thunderbird UI changes; `OAuth2Module` (token of the existing sign-in). |
| `outlookPane` | Outlook-style reading pane and inline reply/new-message pane; signature import; reads whether Thunderbird Conversations is installed (section 6). | Loads its own `reading-pane/pane.js` from the package (`loadSubScript` on a `chrome://` URL registered to the add-on's own folder; never remote code). Reads only a file the user picks in a file dialog (Outlook signature `.htm`) and the images that signature references. Every signature (from that file or from a sent mail) goes through Thunderbird's HTML sanitizer before it is stored and again before it is shown (section 5a). `AddonManager` read only (state of Conversations). |
| `outlookSync` | Reads and applies preferences, account/identity settings, filters and address books for settings sync; applies three recommended settings once (HTML compose, reply above quote, Cards view). | Writes Thunderbird preferences (with the block lists above). Reads and writes the account's message filter file (`msgFilterRules.dat`) in the profile. |
| `outlookSearch` | Outlook-style search box with suggestions; results shown in the message list. | Thunderbird UI changes; one UI preference. |
| `outlookServer` | Notices when Microsoft's server is failing and when it is back; the Graph folder list sync for checks and for ⟳ Refresh (section 4b). | Wraps `onAlert` of Thunderbird's `alertHook` (put back when the add-on stops); runs `syncFolderHierarchy` and `getNewMessages` of the user's Graph accounts. No network access of its own. |
| `outlookContacts` | New contacts in the company format in the address book (section 4a). | Loads its own `contacts/contact-form.js` into the address book page (`loadSubScript` on a `chrome://` URL registered to the add-on's own `contacts/` folder; never remote code). Reads the address books and the Gloda index. |

**5a. Signatures (fixed in 3.9.2).** A signature can be taken from an `.htm` file or from a mail the user selects, so its HTML is treated as untrusted:

- It goes through Thunderbird's own sanitizer (`nsIParserUtils.sanitize`, flags `SanitizerAllowStyle | SanitizerDropForms`) when it is stored and again each time it is shown. This removes scripts, event handler attributes, `javascript:` URLs, frames, objects, embeds, forms and `<meta>`/`<base>`.
- Since 3.9.2, "Use the signature in the selected mail" only accepts a mail whose sender is one of the user's own addresses. The sanitizer remains the main control, because a From line can be forged.
- Remote (`https:`) images and CSS in a signature are kept, as Outlook signatures often load their logo that way.

A test on Thunderbird 159 imported a signature containing about 20 XSS payloads (script tags, `on*` handlers, `javascript:` and `data:text/html` links, SVG, iframe `srcdoc`, forms, `<meta>` refresh, mutation XSS with `<noscript>`/`<math>`) from a mail with a forged own address: no script survived and nothing ran when the stored signature was rendered in an extension page. A mail from another sender was refused.

The code does not use `eval`, `new Function`, process execution, the password manager or remote script loading. File access is limited to the cases listed above.

## 6. Third-party components

The add-on does **not** download, install or enable any software.

The reading-pane features use **Thunderbird Conversations** (`gconversation@xulforum.org`), an open-source add-on (MPL 2.0) published on Mozilla's add-on site, if the user has it.

- The add-on only reads whether Conversations is installed and enabled (`AddonManager.getAddonByID`, read only).
- If it is missing or turned off, the user gets one notification, and a note in Preferences says to install or turn it on through your usual channel for add-ons, or to turn the reading-pane features off.
- Without Conversations, those features do nothing and replies open Thunderbird's normal compose window.
- To show every message of a conversation open (setting *Open every message*), `reading-pane/pane.js` passes Conversations' page its own display action `messages/msgExpand` (`{ id, expand: true }`) for each closed message, newest first and one loading at a time. A message that loaded empty (a Microsoft Graph download cut off) is loaded again, at most three times (at once, after 3 s, after 10 s) and again when the user clicks the "could not be loaded" notice, with Conversations' own `messages/setHasRemoteContent` action (`true`, then `false`), the way Conversations reloads a message after "show remote content"; remote content stays blocked. These actions only change what the page shows. It does not mark mail read, change a message, or call any API.
- To fold the earlier mail quoted below a message (setting *Fold the earlier mail quoted*), `pane.js` moves that message's own nodes into hidden boxes inside the message's frame and adds a "show quoted text" button there. Nothing from the message is copied into the add-on's page, the stored message is not changed, and replies are still built from the original message by the experiment.

Versions up to 3.9.0 could install Conversations automatically from addons.thunderbird.net. That code was removed in 3.9.1, both the download in the background script and the install function in the `outlookPane` experiment. The privileged code has no install path left.

Tests on Thunderbird 159 confirmed that:

- with Conversations missing, nothing is downloaded or installed;
- with Conversations installed but turned off, it stays off;
- with Conversations on, the reading pane works.

## 7. Data stored on the computer

All of the following is stored in the Thunderbird profile, in the add-on's own `storage.local`:

- settings and Quick Steps;
- the list of scheduled messages (subject, recipients, send time); the message itself waits in the Drafts folder;
- the pin sync queue (message IDs);
- the time of the last Sent Items check per Graph account (section 4c);
- settings-sync bookkeeping: the last synced state and one backup for *Undo last sync*.

The add-on also changes Thunderbird's own data, all of it local:

- local tags on messages (📌 Pinned, @ Mentioned);
- added contacts in Collected Addresses;
- a hidden custom column.

Diagnostics (recent errors) are kept only in memory and shown on the Preferences page.

## 8. Actions on mail

**Saving a conversation (3.9.40; one file since 3.9.41; the conversation as its text since 3.9.43).** The reading pane's ⤓ button saves the conversation shown as one `.eml` file, only when the user clicks it, at a place the user picks in Thunderbird's own file picker (which asks before replacing a file). Each message is read as stored through Thunderbird's own `MessageSaver.saveToFile` into a temporary file created in the system temp folder (unique name, readable only by the user), which is removed right after, and parsed with Thunderbird's own MIME parser. The file written has the conversation as its HTML text (newest first, each message under an Outlook-style header, without the earlier mail it quotes), with scripts, frames, forms, event-handler attributes and `javascript:` URLs removed; the pictures it shows; and the messages' attachments, unchanged. The experiment saves only messages of the conversation selected in the message list (its thread, or linked to it by Message-ID / References, such as replies in Sent Items); ids of other messages sent by the page are ignored. The file name is the subject and the date, with characters that are not allowed in file names replaced. Nothing is sent anywhere.

Mail is sent, moved or deleted only on user action, with one exception (second copies of sent mail, below):

- Send with Undo Send (the message is held for 10 seconds, then sent through Thunderbird);
- a scheduled time the user chose;
- Quick Steps, keyboard shortcuts and menu entries;
- the cleanup page, which lists the matching mail before acting.

The exception: on Graph accounts, a second copy of a sent mail in Sent Items (same Message-ID, created within 10 minutes of the copy kept) is moved to Deleted Items automatically (section 4c; can be turned off in Preferences → *Sending*). Nothing is deleted permanently.

On the cleanup page, permanent deletion always asks for confirmation. Moving to Deleted Items or archiving asks when more than 500 messages are involved. Cleanup never runs on its own, and pinned mail is never cleaned. Nothing is forwarded or sent automatically to addresses the user did not enter.

## 8a. Input checks (3.9.3 to 3.9.18)

Hardening:

- **Every handler (3.9.4).** All message handlers except the six used by the compose script are wrapped in `fromOwnPage()`, so each one checks its caller itself.
- **Message routing.** The `sender` of a message (`id`, `url`, `tab`) is filled in by Thunderbird, not by the page that sends it, so a page cannot claim to be another page, another tab or another add-on. `runtime.onMessage` only receives messages from this add-on; there is no `onMessageExternal` listener, so other add-ons and web pages cannot call it. The add-on's own pages (`moz-extension:`) may call every handler. The compose script, which runs inside compose windows next to mail content, may call only the six handlers it needs (mentions, Undo Send, compose info). Handlers that read or change mail, settings or sync (cleanup, move, Quick Steps, scheduled sends, signature import, settings sync, refresh, contact collection) also check at the handler itself that the caller is one of the add-on's own pages.
- **Undo Send.** The delay is capped at 60 seconds (the longest choice in Preferences), whatever the stored value; the same cap applies in the reading pane.
- **Move to folder and Quick Steps.** The destination must be an existing folder of one of the user's own accounts; otherwise nothing is moved.
- **Quick Steps from storage or settings sync.** Each field is checked: unknown actions become "none", and a forward keeps only well-formed addresses. A forward Quick Step only opens a compose window with the recipients filled in; the user checks and sends it. Nothing is forwarded automatically.
- **Scheduled sends.** Cancel and "send now" act only on an id that exists in the schedule list, and since 3.9.8 only when asked from the page that shows that list (see below).
- **Mailbox cleanup.** The scan only accepts one of the user's mail accounts; the apply step only accepts the actions archive, trash and delete, and numeric message ids.
- **Pictures in the reading pane editor.** Only PNG, JPEG, GIF, BMP and WebP pictures are accepted, recognised by their first bytes, not by the file name or the type the file claims; SVG and other files are refused. The picture's `data:` URL carries the detected type.
- **Pop-out of a reply written in the reading pane.** Its HTML goes through the same sanitizer as signatures (section 5a) in the experiment and again in the background script, right before it is put into the compose window.
- **Reading pane editor.** A signature or an earlier draft is sanitized again each time it is put into the editor. A @mention link is only made for a plain e-mail address; the address and the name are set as an attribute and as text, so they are escaped.
- **Microsoft Graph.** `graphFetch` refuses any URL outside `https://graph.microsoft.com/v1.0/`, so the access token cannot be sent elsewhere, even through a paging link.
- **Settings sync.** The downloaded sync file is refused when larger than 3 MB, or larger than 32 MB once unpacked (unpacking stops at that point, so a decompression bomb cannot exhaust memory). Add-on settings from another device are only applied when the setting exists in this add-on, is not a per-device setting, and the value has the same type as the default (true/false, number or text); Quick Steps are checked as above. Thunderbird preferences go through the block lists of section 4.
- **Settings sync, contacts (3.9.4).** A contact from another device must be a single vCard of at most 64 KB made only of property lines, with a plain id; it is imported in the clean form (no photo or UID line).
- **Settings sync, filters and identities (3.9.4).** See section 4: filters that forward, reply or run custom actions, automatic Cc/Bcc, Reply-To and protection settings are never applied from another device.
- **Pop-out (3.9.4).** The fields sent to the compose window are checked: known type, numeric message id, one of the user's own identities, well-formed addresses, subject of at most 998 characters. The compose window shows them; the user sends.
- **Quick Step fields (3.9.5).** The destination is a Thunderbird folder id (`account<n>://Folder/Sub`), never a file path; ids of another form, or with a `..` part or control characters, are dropped, and runStep() only moves to an existing folder of the user's accounts. Names are one line of plain text; menu labels are not HTML, so markup in a name shows as text.
- **Cleanup (3.9.7).** The apply step only acts on mail listed by the last scan in the background (the page can only choose among those); any other id refuses the whole request. All accounts it can reach are the user's own accounts in this Thunderbird profile.
- **Signature preview (3.9.7).** The Preferences preview shows the sanitized signature in an iframe sandboxed without `allow-scripts`, with its own Content-Security-Policy (`default-src 'none'`; images from `data:`/`https:`; since 3.9.17 only `style` attributes, no style sheets).
- **Pin sync queries (3.9.7).** The Message-ID is sent as an OData string literal (single quotes doubled, which is OData's escaping, then URL-encoded), so it cannot add operators to the `$filter`; Message-IDs with spaces, angle brackets or control characters are not looked up at all.
- **Undo Send "Send now" (3.9.8).** The banner's *Send now* only finishes a send the user already started: it acts only for a compose tab (`sender.tab.type` is `messageCompose`) that has an Undo Send countdown running, and only for that tab. Otherwise nothing is sent.
- **@mention recipient (3.9.8).** The person added to To must be in the user's address books (as the mention popup offers); the address and name are taken from the address book, not from the page. A value such as `a@x.com>, b@y.com` (which would add a second recipient) is refused. Quotes, angle brackets, backslashes and control characters are removed from the name. Only a compose tab may ask.
- **Scheduled mail handlers (3.9.8).** Each handler accepts only the add-on pages that use it: scheduling from the *Send later* popup; the list from Preferences and the Quick Steps popup; cancel and "send now" only from Preferences. Other add-on pages, such as the new-message pane, cannot call them. The add-on runs in one user's Thunderbird profile, so there is no other user whose scheduled mail could be reached.
- **Thread root lookup (3.9.8).** When pinning, the References and In-Reply-To ids of a mail are looked up in the same folder only when they are well-formed Message-IDs (no spaces, angle brackets or control characters). `messages.query` compares the whole Message-ID (it has no query language), and the result is only used to pin that thread; nothing is returned to the mail's sender.
- **Mark conversation read (3.9.8).** The mail marked read is read from the mail tab's own selection in the privileged code; the ids passed in must equal that selection, or nothing is marked. One row must be selected.
- **Pin buttons (3.9.8).** In the message list, only Thunderbird's own `toggle-flag` event (trusted, on the list's table body, for an existing row) pins a conversation. In the reading pane, each page gets its own bridge object that knows its page (read-only); its pin button can only pin the conversation selected in the message list of the mail tab the page is in, and refuses any other message id, so a page showing mail content cannot pin or unpin other mail.
- **History (3.9.4).** Only a plain e-mail address is looked up (`messages.query` matches text in local headers; it has no query language).
- **Pin sync.** At most 100 pin changes are uploaded per run, one request at a time; the rest stay queued for the next run (every 5 minutes).
- **Contact collection.** A full scan that is already queued or running is reused, so repeated requests do not stack scans; the refresh button runs one refresh at a time.

## 8b. Hardening in 3.9.8

| Finding | Assessment | 3.9.8 |
| --- | --- | --- |
| `undo-send-now` uses `sender.tab.id` without checks (CWE-346) | `sender` is set by Thunderbird and cannot be spoofed by a page; Conversations' reading pane is another add-on and is refused by the router. Real gap: the handler sent a compose tab even with no countdown running. | Sends only a compose tab with its own countdown running (section 8a). |
| XSS through the `email` of `mention-add` (CWE-79) | Not XSS: recipients are set with `compose.setComposeDetails` and shown as text, not HTML. Real gap: a crafted value could add extra recipients. | Only address-book people, with the address book's address and name (section 8a). |
| No authorization for scheduled mail actions (CWE-284) | Single-user profile: there are no other users' scheduled mails. Handlers were limited to the add-on's own pages. | Each handler limited to the page that uses it (section 8a). |
| Crafted References / In-Reply-To in `messages.query` (CWE-200) | Not exploitable: `messages.query` compares the whole `headerMessageId`, in the same folder, and the result is never returned to anyone. | Only well-formed Message-IDs are looked up (section 8a). |
| IDOR in mark read on `onSelectedMessagesChanged` (CWE-284) | The event is raised by Thunderbird for the user's own selection; a page cannot raise it. | The privileged code marks only the tab's real selection (section 8a). |
| No authorization for pin/unpin (`onPinButton`, CWE-284) | The list event is raised by the privileged code from Thunderbird's own click event; pages cannot raise it. Real gap: the reading-pane bridge accepted any message id from the page. | Trusted list events only; the reading pane can pin only the conversation it shows (section 8a). |

The checks are covered by `tests/handlers.test.js` and were tried on Thunderbird 159: a `toggle-flag` event made by a script, a reading-pane pin of a mail that is not shown, and a mark-read call with ids other than the selection all do nothing, while the real click, the real selection and the reading-pane pin of the shown conversation work.

## 8c. Hardening in 3.9.9

| Finding | Assessment | 3.9.9 |
| --- | --- | --- |
| Signature import XSS, `signature.js:116` and `:139` (CWE-79) | `DOMParser` makes an inert document, and `outlookPane.setSignature` already sanitized the HTML before storing it (section 5a), but the background code passed the HTML on unsanitized. | `signature.js` now sanitizes the HTML itself (`outlookPane.sanitizeHtml`, Thunderbird's sanitizer) right before `setSignature`. Tested on Thunderbird 159 with about 20 payloads from a mail with a forged own address: no script, event handler, `javascript:` URL or frame is stored, nothing runs; Outlook layout (cell padding, `MsoNormal`, pictures) is kept. |
| Quick Step menu title XSS, `main.js:49` (CWE-79) | `menus.create` shows `title` as plain text; it is never parsed as HTML. | Markup characters (`< > & " '` and backtick) are also removed from the name before it becomes a menu label. |
| `configureSearch` applies the stored setting without validation (CWE-284) | Storage is written only by the add-on's own pages and by settings sync, which checks each value. | Only a true/false value is used (anything else means the default), and only changes to the add-on's local storage are applied. |
| `history` uses `tabId` without checks (CWE-284) | Called only from the add-on's own popup. | The tab id must be an integer naming an existing tab that is active (the one the user is looking at); any other id gets "No message is open." and no data. |

## 8d. Hardening in 3.9.10

| Finding | 3.9.10 |
| --- | --- |
| Unbounded Undo Send delay, `options.js:465` (CWE-1284) | Preferences stores only 0 to 60 seconds (whole seconds; anything else becomes 0). The background clamps the stored value the same way. |
| Settings sync timestamp, `settings-sync.js:507` (CWE-693) | A synced entry must have a numeric time that is not negative and not more than one day in the future; otherwise it is ignored and this device's value is kept (and uploaded again). |
| `run-step` with any `tabId`, `quicksteps.js:45` (CWE-284) | Accepted only from the Quick Steps popup, for an existing mail tab that is active (the one the user is looking at), and only for an existing step index. |
| Pin sync IDOR, `pin-sync.js:505` (CWE-284) | Each message is looked up again in Thunderbird by its id; its account and Message-ID come from Thunderbird's copy, not from the object passed in. Only messages that exist in the user's own Microsoft 365 accounts are queued. |
| `history` with any `tabId`, `history.js:11` (CWE-284) | Accepted only from the History popup. The tab is found by the background (the active tab of the last focused window); the tab id in the request is not used. |
| Own pages may send every request, `main.js:163` (CWE-284) | Each add-on page may send only the requests it uses (a fixed list per page in `main.js`); any other request from it is refused and logged. The compose script keeps its own short list. |
| Quick Step actions without authorization, `popup/quicksteps.js:59` (CWE-284) | Same checks as `run-step` above. Steps are checked again when they run: unknown actions become "none", a move needs a folder of the user's accounts, *delete* moves to Trash, and *forward* only opens a compose window for the user to send. |
| Scheduled time not checked, `popup/schedule.js:55` (CWE-20) | The popup and the background both require a time in the future and within a year. |

## 8e. Hardening in 3.9.11

| Finding | 3.9.11 |
| --- | --- |
| Quick Step forward to any address, `quicksteps.js:27` (CWE-20) | The forward compose window is filled in only with recipients found in the user's address books; other addresses of the step are left out. It still only opens the window; the user checks and sends. |
| Quick Step move to any folder, `quicksteps.js:31` (CWE-284) | A step moves mail only to a folder of the same account as the mail (and the folder must exist in the user's accounts); a folder of another account is refused with a message. |
| ReDoS in the mention scan, `mentioned-me.js:28` (CWE-1333) | The anchor pattern has bounded parts (attributes up to 1000 characters without `<`, link text up to 300), so the scan is linear: 1 MB of crafted `<a>` tags takes under 0.1 s (measured with Node.js). Real mention links are still found. |
| Pin sync with any message id, `pin-sync.js:501` (CWE-284) | Besides re-reading the message in Thunderbird, a change is uploaded only when Thunderbird's own copy already has (or no longer has) the pin tag, so the server only follows what Thunderbird shows. |
| History query with an address from the mail, `history.js:32` (CWE-200) | The search is limited to the account of the displayed mail (`messages.query` with `accountId`). The results are shown only to the user in their own popup; nothing is sent anywhere. |
| Reading pane `messageDetails` with any id, `pane.js:400` (CWE-284) | The privileged bridge returns To/Cc details only for mail of the conversation the page shows (the same check as the pin button), and nothing for any other id. |

## 8f. Hardening in 3.9.12

| Finding | 3.9.12 |
| --- | --- |
| Contact sync uses the book name and contact id from the sync data, `settings-sync.js:267` (CWE-284) | Contacts are changed or deleted only when the contact id is in this device's own local, writable address book that the change names (read from Thunderbird in the same sync run); a new contact must have an id no local contact has. CardDAV, LDAP and read-only books are never touched. A book name with control characters is refused, and one sync creates at most 5 new address books. The same checks apply to *Undo last sync*. |

## 8g. Hardening in 3.9.13

| Finding | 3.9.13 |
| --- | --- |
| Pop-out reply HTML, `reading-pane.js:115` (CWE-79) | After Thunderbird's sanitizer, the background checks the HTML again right before it goes into the compose body: script, frame, object, embed, form, input, meta, base, link, SVG, MathML and template elements are removed, and so are every `on*` attribute, `srcdoc`, and `javascript:` / `vbscript:` / `data:` URLs (only `data:image/` pictures are kept). Tested on Thunderbird 159: a pop-out with script, `onerror`, `onclick` and `javascript:` links ends up in the compose window with none of them, while text formatting, links, styles and quotes stay. |
| Crafted References / In-Reply-To in the thread lookup, `pin.js:51` (CWE-200) | Only a plain `left@right` Message-ID is looked up (letters, digits and the usual Message-ID punctuation such as `.`, `-`, `_`, `+`, `=`); ids with `*`, `?`, `%`, brackets, spaces or other characters are not looked up at all. `messages.query` compares the whole id in the same folder, and the result is only used to pin that thread. |

## 8h. Hardening in 3.9.16

| Finding | 3.9.16 |
| --- | --- |
| Search turned on or off from storage, `search.js:11` (CWE-284) | The on/off choice is no longer read from the add-on's storage. The `outlookSearch` experiment keeps it in a Thunderbird preference (`extensions.outlookmode.search.enabled`, default on) and changes it only when the Preferences page asks (`search-set`, accepted only from `options/options.html`, true/false only). Settings sync does not cover `extensions.` preferences, and the setting is no longer synced. Tested on Thunderbird 159: writing the setting into storage changes nothing; another add-on page asking is refused. |
| Path traversal in synced Quick Step folders, `settings-sync.js:133` (CWE-22) | A synced folder path must start with `//` and pass the folder id check (no `..` part, no control characters); otherwise the step is kept without a folder, to be picked again. When it runs, the folder must also exist in the user's accounts and be in the same account as the mail (since 3.9.11). |
| OData injection through the Message-ID, `pin-sync.js:76` (CWE-74) | Only a plain `left@right` Message-ID (letters, digits and `.`, `-`, `_`, `+`, `=` and similar; no quotes, parentheses, brackets, spaces, `*`, `?`, `%`) is looked up on the server; it is still sent as an OData string literal and URL-encoded. The same pattern is used for the thread lookup in `pin.js`. |

## 8i. Hardening in 3.9.17

| Finding | 3.9.17 |
| --- | --- |
| CSS injection in the signature preview, `options.js:537` (CWE-79) | The preview's policy allows only `style` attributes (`style-src-attr 'unsafe-inline'`), no `<style>` sheet, so CSS selectors cannot be run on the content; scripts, frames, forms and connections stay blocked and the frame stays sandboxed without `allow-scripts`. The signature is still sanitized first (which already drops `<style>`). |
| Quick Step run without authorization, `quicksteps.js:55` (CWE-284) | A step runs only for the user's own action on the tab they are looking at: the tab must be an active mail or message tab, whether the step comes from the menu, a shortcut or the popup. It acts on the mail selected there, with the checks of 8d and 8e. Tested on Thunderbird 159: an inactive or non-mail tab, or no tab, runs nothing. |
| Forged From in the signature import, `signature.js:106` (CWE-346) | Besides the From address being one of the user's own, the mail must be in a Sent folder of the user's accounts (where received mail does not go). A mail in the Inbox with a forged own address is refused. The HTML is still sanitized before it is stored. |
| Script URL in CSS of the pop-out reply, `reading-pane.js:139` (CWE-79) | `inertHtml` also removes `style` attributes, and `<style>` elements, whose CSS has `javascript:`, `vbscript:`, `expression(`, `-moz-binding`, `behavior:` or a CSS escape (`\`, which could hide those words). Plain styles stay. |
| Local file path in reading-pane drafts, `pane.js:1506` (CWE-200) | The pane no longer reads or keeps the path of an attached file. The privileged experiment reads it and keeps it in memory; the pane and its drafts hold only a token (`file-1`, ...) with the name, size and type. Drafts are never sent anywhere (they stay in memory). Tested on Thunderbird 159: the draft has no path; the file is attached and sent; the sent mail has no path. |
| Unvalidated Thunderbird settings from the sync, `settings-sync.js:347` (CWE-20) | Before they reach the experiment, the background checks each Thunderbird setting: the key must be one this device syncs (`outlookSync.supportedKeys`: allowed preference branches, existing accounts and identities, block lists of section 4), and the value must be true/false, a whole number, text up to 512 KB, or empty (reset). Others are refused and listed as failed. The experiment checks the same again. The same applies to *Undo last sync*. |
| Markup in synced vCards, `settings-sync.js:308` (CWE-79) | A synced vCard is refused when any value or parameter has `<` or `>`, or the card has `javascript:` or `vbscript:`. Thunderbird shows contact fields as text in any case. |

## 8j. Hardening in 3.9.18

| Finding | 3.9.18 |
| --- | --- |
| No upper bound for the cleanup mailbox limit, `cleanup.js:585` (CWE-20) | The limit is at most 1024 GB (1 TB; Microsoft 365 mailboxes are far smaller): a larger number is stored as 1024 GB and shown so in the field, which also has `max="1024"`. Values that are not a positive number (including `Infinity`) remove the limit. |
| IDOR when pinning, `pin.js:118` (CWE-639) | Before tags are changed, every message is read again from Thunderbird by its id, and only mail in one of the user's own accounts of this profile is changed; its tags are taken from Thunderbird's copy, not from the object passed in. An id that does not exist changes nothing. (The pin events themselves come from Thunderbird's own UI, see 8b.) Tested on Thunderbird 159. |

## 9. Summary

| Area | Assessment |
| --- | --- |
| Data leaving your Microsoft 365 tenant | None. The only host contacted is `graph.microsoft.com`, for the user's own mailbox. |
| Credentials | Uses Thunderbird's existing OAuth token in memory, only towards `graph.microsoft.com`. Never stored, logged or synced. |
| Privilege level | High: six Experiments with full Thunderbird privileges, needed for the UI changes. Mitigation: readable source in the package; no remote code. |
| Supply chain | None at run time. The add-on installs no software; the optional Thunderbird Conversations add-on is installed by the user through your usual channel for add-ons (section 6). |
| Package integrity | Not Mozilla-signed. Verify with the SHA-256 above. |
| Write access to the mailbox | Outlook pin properties (default on, Graph accounts), moving second copies of sent mail to Deleted Items (default on, Graph accounts) and one hidden sync folder (settings sync, default off). |

## 10. How to verify the package

```bash
sha256sum outlook-mode.xpi                     # must match section 1
mkdir /tmp/om && unzip -q outlook-mode.xpi -d /tmp/om
grep -rnE "fetch\(|XMLHttpRequest|WebSocket|sendBeacon" /tmp/om --include=*.js    # one fetch(), background/graph.js
grep -rnoE "https?://[^\"'\` )]+" /tmp/om --include=*.js | grep -v w3.org          # graph.microsoft.com only (plus "https://" link prefix in the editor)
grep -rn "getInstallForURL\|getInstallForFile" /tmp/om                             # nothing: no install code
grep -rn "Services.logins" /tmp/om                                                 # nothing: no password manager use
```
