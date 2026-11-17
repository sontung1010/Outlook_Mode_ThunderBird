# Outlook Mode for Thunderbird

`outlook-mode.xpi` (v3.9.33) is a Thunderbird add-on that makes Thunderbird behave like Outlook. It adds:

- reply and new mail inside the main window;
- newest-first threads;
- pins synced with Outlook;
- @mentions, Undo Send, Schedule send and Quick Steps;
- sync of settings and contacts between your devices;
- new contacts in the company format, Name (Company-Department).

## 1. Requirements

| Dependency | Required? | Needed for |
| --- | --- | --- |
| **Thunderbird 140–159** | Yes | Everything. 128–139 works without the reading pane, because Conversations needs 140+. Tested on 159; the 140 ESR from `apt` is untested. |
| **Microsoft 365 account as *Microsoft Graph*** | For sync | Pin sync with Outlook, and settings/contacts sync between devices. Add it in Account Settings → Add Mail Account → Microsoft → **Microsoft Graph** (OAuth2). |
| **[Thunderbird Conversations](https://addons.thunderbird.net/thunderbird/addon/gmail-conversation-view/)** add-on | Yes, for the reading pane. **Install it yourself** through your usual channel for add-ons; Outlook Mode never downloads or installs it | Newest-first thread view, reply/forward inside the thread, pin button in the thread. Version 4.3.18 supports Thunderbird 140 to 159.x. Without it, Reply opens a normal compose window. |

## 2. Install

1. Add the Microsoft 365 account (Microsoft Graph).
2. Install **Outlook Mode**: ☰ → Add-ons and Themes → ⚙ → *Install Add-on From File…* → `outlook-mode.xpi`.
3. Wait about 15 seconds. Outlook Mode then sets things up by itself. A notification says what it did.
   - If **Thunderbird Conversations** is missing or turned off while a reading pane feature is on, a notification says so (once). Outlook Mode does not install or turn it on.
   - It sets the Thunderbird settings listed below.
4. Optional: open **Preferences** (Outlook Mode's settings page) and review it. Everything is on by default except *Sync between devices*. Ways to open it:
   - the **⚡ Preferences** button in the toolbar → **⚙ Settings** (or **✍ Signature** to go straight to the signature);
   - **Tools → Outlook Mode preferences…** (press Alt to show the menu bar);
   - in a reply or new mail, **✍ Signature ▾ → Signatures…**;
   - Add-ons and Themes → Outlook Mode → ⋯ → *Options*.

   The page has a list of sections at the top.

To update, install the new `.xpi` over the old one. Your settings are kept.

**Conversations** is checked at every start while a reading pane feature is on. Outlook Mode only reports its state: it never downloads, installs or turns on any add-on, and makes no network request for it. If Conversations is missing or turned off, a notification and Preferences → *Reading pane* say what to do: install or turn it on through your usual channel for add-ons, or turn the reading pane features off.

**Thunderbird settings** are set once on each device, the way the features expect:

- *Compose messages in HTML format*, needed for @mention links;
- *Start my reply above the quote*, as in Outlook;
- *Cards View*, needed for the pin button and the unread highlight in the message list.

If you change these back later, they stay as you set them. Preferences → *Apply recommended Thunderbird settings* sets them again.

## 3. Features

| Feature | How to use |
| --- | --- |
| **Reply inside the thread** | Reply, Reply all, Forward (buttons or Ctrl+R / Ctrl+Shift+R / Ctrl+L) open a reply box above the newest mail, with no pop-up window. It has To/Cc/Bcc, the formatting tools below, signature, attachments and address suggestions. Ctrl+Enter sends. **Pop out** moves the reply to a compose window (for encryption, read receipts or inline images). |
| **New mail in the main window** | Write / New Message / Ctrl+N opens a pane at the bottom of the window, one tab per message. Use ▾ to minimize, □ to maximize, or drag the top edge to resize. |
| **Formatting (like Outlook's Format Text)** | In replies, forwards and new mail. The toolbar has: undo/redo; font and size in points; grow/shrink; bold, italic, underline, strikethrough, sub/superscript; highlight and font color (click to reuse the last color, ▾ for more); clear formatting; bullets, numbering, indent; align left/center/right/justify; styles (Normal, Heading 1–3, Quote, Code); link, picture (or paste one), table (pick the size on a grid), horizontal line, emoji; attach. Outlook's keys work in the text: Ctrl+B/I/U, Ctrl+K link, Ctrl+E/L/R/J align, Ctrl+]/Ctrl+[ font size, Ctrl+Shift+L bullets, Ctrl+M/Ctrl+Shift+M indent, Ctrl+= / Ctrl+Shift++ sub/superscript, Ctrl+Space clear, Ctrl+Shift+F/P font/size, Ctrl+Shift+V paste as text only (in the font where it goes). Mail is written in Calibri 11 pt, like Outlook, unless Thunderbird has another default font. The message area is always white with black text, as recipients see it, in light and dark theme alike. |
| **Thread view** | The newest mail is on top. Like Outlook, every mail of the conversation is open and shown on its own, with its own text and its own attachments (the earlier mail a reply quotes below its text, such as Outlook's *From / Sent* block, is folded under *- show quoted text -*, since it is already shown above as a message of its own; the oldest message and a forward with no text of its own keep theirs). Mail is opened newest first, one at a time: on a Microsoft Graph account, mail not yet stored on the computer is downloaded first, and a second load at the same time cuts that download off (the mail would show empty). A mail that still loads empty is loaded again by itself (at once, then after 3 and 10 seconds); while it loads it says *Loading this message…*, and if the server still does not send it, *This message could not be loaded from the server. Click here to try again.* The first time, a long thread of such mail takes a few seconds per mail; after that it is stored and opens at once. As in Outlook, each mail's attachments are shown above its text, right under its header. Click a mail's header (or press O on it) to close it; it stays closed while that conversation is shown. With *Open every message* turned off in Preferences, Conversations decides what is open, and *Show only the newest message* puts earlier mail behind *Show N earlier messages*. Like Outlook, the reading pane is white (also with a dark theme); the ◐ button switches the whole pane, header and mail text, to dark and back (Conversations' button, or Thunderbird's own toggle in the message header when Conversations is not used), and an open mail's header has separate **To** and **Cc** lines with the full date; click them, or **▾ N**, to see everyone with their addresses. Pictures wider than the pane are shown at the pane's width (aspect ratio kept), and a mail laid out wider than the pane as a whole (newsletters of a fixed 600–800 px, text over pictures) is shown smaller to fit the pane's width, as in Outlook; both follow the pane when it gets wider or narrower, without moving where you are in the mail. |
| **Pin** | Click 📌 on a mail (in the list or in the thread; it replaces the star), or press Alt+Shift+P. Red means pinned. Pinning applies to the whole conversation: every mail in it, as Outlook does, and new replies to it are pinned too. The Inbox keeps pinned conversations on top (View → Sort by → *📌 Pinned first*). A conversation that gets new mail moves up by its newest mail, as in Outlook. Pins sync both ways with Outlook (Microsoft Graph accounts): a pin made in Thunderbird shows in Outlook on the web, the new Outlook for Windows and the Outlook mobile app, and a pin made there shows in Thunderbird, within 5 minutes or at once with ⟳ Refresh. Classic Outlook for Windows has no pins. Preferences → Diagnostics shows the last pin sync. |
| **Message list like Outlook** | In the cards view, pinned conversations sit in their own **📌 Pinned** section at the top, and the rest is grouped by date as **Today**, **Yesterday**, **This week** (the last 7 days, today included), **Older**. **One card per conversation**, showing who sent its newest mail and when: threads stay closed (no *N replies* button), and the reading pane shows the whole conversation. Unread mail (and a conversation with unread mail) has Outlook's blue bar on the left, a light blue card, a bold sender and a blue bold subject and time, in place of Thunderbird's dot. Opening a mail marks its whole conversation as read, as in Outlook (Thunderbird alone marks only the mail you click). Each can be turned off in Preferences. |
| **@mention** | Type `@name` in the body and pick a person (↑ ↓, Enter). Diacritics are ignored, so `@tung` finds "Tùng". A link is inserted and the person is added to To. |
| **@Mentioned me** | Mail that @mentions you gets the **@ Mentioned** tag. Alt+Shift+M shows only those. |
| **Undo Send** | After Send, a bar shows *Sending in 10s… Undo / Send now*. The delay is set in Preferences (off, 5–60 s). |
| **Schedule send** | Compose window → **Send later** → pick a time. Thunderbird must be running at that time; mail that became due while it was closed is sent at the next start. |
| **Signature from Outlook** | Preferences → *Signature* (fastest: ⚡ → **✍ Signature**) → pick the identity, then either: **Use the signature in the selected mail** — select (in the message list, e.g. Sent Items) a mail you sent from Outlook with your signature; works with new Outlook and Outlook on the web, whose signatures live in the cloud; or **Import Outlook signature (.htm)…** — classic Outlook's file from `%APPDATA%\Microsoft\Signatures` (copy the `.htm` with its `_files` folder). Tables, fonts, colours and pictures are kept, and the preview shows the result. The pictures are stored inside the signature, so it shows everywhere, is sent as inline images, and syncs to your other devices. |
| **Outlook reply header** | Replies and forwards quote the original with an Outlook `From / Sent / To / Cc / Subject` block and no quote bar. |
| **Search (like Outlook)** | Type in the search box at the top (Ctrl+K). While you type, it suggests **people** (also without diacritics: `hoang` finds "Hoàng") and **matching mail**, plus recent searches. Press Enter: the results show **in the message list of the same tab**, from all folders, without opening a new tab. The bar above them switches between **All folders** and **Current folder**, adds quick filters (**Unread, Attachments, Pinned, To me, From me**), and **✕ Exit search** (or Esc, or clicking a folder) goes back. It searches subject, sender, To/Cc names and addresses, and the message text (Thunderbird's index). Keywords as in Outlook: `from:` `to:` `cc:` `subject:` `body:` `has:attachment` `is:unread` `is:pinned` `before:2026-10-01` `after:` `"exact phrase"`. |
| **Clean up mailbox 🧹** | ⚡ → **🧹 Clean up mailbox**, Tools → *Clean up mailbox…*, or Preferences → *Mailbox cleanup*. Shows how full the mailbox is, then one click to review, starting with **🤖 notifications & no-reply** (Bitbucket, Jira, SharePoint, newsletters…; recognised by address and by service domain, so a Bitbucket mail named after the colleague who commented still counts as Bitbucket), then mail **older than 3 / 6 months / 1 / 2 years**, **large mail** (over 5 MB), mail **unread for over a month**, **Junk**, and the **top senders** by size (automatic senders per service, apart from people). Or pick folders (Outlook system folders such as Sync Issues, RSS Subscriptions and Conversation History are left out), dates, sender, subject, size, attachments, read state. Every mail is listed before anything happens, grouped by month with a box for each month; untick what to keep. 📌 Pinned mail is **always** kept, by every option including Empty Deleted Items, until you unpin it (checked again right before acting, so a pin made on another device counts too); 🚩 flagged mail is kept unless you untick it. **Move to Deleted Items** (can be undone), **Archive**, or **Delete permanently** (asked first); then **Empty Deleted Items**, since Deleted Items still counts toward the mailbox size. |
| **Refresh ⟳ / F5** | The button above the message list, or F5, gets new mail for every account, updates the open folders and syncs settings, pins and contacts now. For a Microsoft Graph account it also runs the **folder list sync** (Thunderbird's own *Get Messages* does not), which restarts a sync that was failing and tells at once whether Microsoft's server is back. It also puts the add-on's additions back into any window that missed them. The tooltip shows the result. Shift+F5 keeps Thunderbird's own "get new mail for all accounts"; F5 can be given back to Thunderbird in Preferences. |
| **Microsoft's server failing** | When Microsoft's server answers with errors for a while (as in an Outlook outage), Thunderbird would show *Server graph.microsoft.com responded with an error.* at every mail check. Instead, after the first one, one notice says **Microsoft's server is failing** (with the reason), Thunderbird's repeated notices for that account are held back, and the add-on's own requests (pin sync, settings sync) wait instead of failing. The add-on tries the folder list sync again after 30 s, 1 and 2 minutes, then every 5 minutes (or at once with ⟳ / F5); when it works, one notice says **Microsoft's server is back**, the account's new mail is fetched and the waiting syncs run. A lost network on your own computer is not counted as Microsoft's server failing. The ⟳ button turns red with the reason in its tooltip meanwhile; Preferences → Diagnostics shows the history. |
| **Collected contacts** | Everyone you wrote to or replied to is added to *Collected Addresses*, so address suggestions know them. |
| **New contact in the company format** | Address Book → **New Contact**. The **Email Addresses** field comes first and has the cursor. Paste *Phạm Minh Châu (Contoso-RND-HN) <chau.pm@contoso.example>* (copy it from a To/From line) and the address, display name, Company (*Contoso*) and Department (*RND-HN*) are filled in at once; or type just the address and its name is taken from your address books (Collected Addresses…) or your mail, if known. The **Company – department** field under the name suggests the codes your contacts already use (those of the same mail domain first; type a few letters to filter, Enter takes the first): picking one sets Company, Department and the *(…)* after the display name. A new code you save is suggested next time. Only data on your computer is used. Turn it off in Preferences → *Sync between devices* (contacts part). |
| **Quick Steps** | The **⚡ Preferences** toolbar button (its *Quick Steps* list) or Ctrl+Shift+1…5 runs a configured action: mark read, pin, reply/forward, then move, archive or delete. Configure them in Preferences. |
| **Move / History** | Ctrl+Shift+V in the message list opens a searchable folder picker. The **History** button above a mail lists recent mail with that person. |

Shortcuts (Thunderbird's own keys take priority, so some differ from Outlook):

| Action | Shortcut |
| --- | --- |
| Delete | Ctrl+D |
| Forward / forward as attachment | Ctrl+L / Ctrl+Alt+F |
| Mark read / unread | Ctrl+Shift+Q / Ctrl+Shift+U |

To change the keys: Add-ons and Themes → ⚙ → *Manage Extension Shortcuts*.

## 4. Sync between devices

1. On the device whose settings you want to keep, open Preferences, tick **Sync between devices**, click **Save**, then click **Sync now**.
2. Do the same on your other devices. They take the synced settings; **Undo last sync** puts a device's own settings back.

What is synced:

- general settings;
- signatures and identity settings;
- message filters;
- local address books (Personal, Collected Addresses);
- this add-on's settings and Quick Steps.

What is never synced: the accounts themselves (add each account once per device), passwords and sign-ins, encryption keys and file paths.

The data lives in a hidden folder *Outlook Mode Sync* in your mailbox. Sync runs at startup, every 10 minutes, and when you click Refresh.

## 5. Good to know

- **Unsent mail:** replies and new messages you have not sent are kept in memory while Thunderbird runs, not in Drafts. Thunderbird asks before quitting if any are left.
- **Thunderbird updates:** if Thunderbird updates past what Conversations supports (4.3.18 stops at 159), the reading pane is off until a newer Conversations is published and installed (through your usual channel for add-ons). In the meantime, replies open a compose window, and a notification says why (once per Thunderbird version).
- **Not available:** Recall, Focused Inbox and Exchange-only features (shared calendars, server-side rules) are not available to Thunderbird add-ons.
- **Turning features off:** every feature can be turned off in Preferences.

## 6. Troubleshooting

- **Preferences → Diagnostics** shows the status of each part and recent errors. Include a screenshot when you report a problem.
- **Reply still opens a window:** Preferences → *Reading pane* shows whether Conversations is running. If it is not, install or turn on Thunderbird Conversations (through your usual channel for add-ons), and check that *Reply … inline in the reading pane* is ticked.
- **@mention shows nothing:**
  - Make sure the message is in HTML mode (Preferences → *Apply recommended Thunderbird settings*).
  - Make sure the person is in an address book. New contacts appear within 5 minutes; *Reload contacts* in Diagnostics refreshes them.

## 7. Development

- **Source:** `outlook-mode/`
  - `background/`: feature logic
  - `compose/`: compose-window script
  - `reading-pane/`: inline reply and new-mail pane
  - `experiments/`: privileged Thunderbird APIs (`outlookPane`: reading pane, compose dock, Conversations state (read only); `outlookPin`: pin column and buttons, refresh; `outlookSync`: settings sync, collected contacts, recommended settings; `outlookSearch`: the search box and results)
  - `options/`, `popup/`: UI
- **Build:** `./build.sh` writes `outlook-mode.xpi` (needs `bash`, `python3`, `zip`).
- **Tests:** need Node.js, plus jsdom (`npm install jsdom@24` in any temporary folder):

  ```bash
  node tests/settings-sync.test.js
  NODE_PATH=<folder>/node_modules node tests/reading-pane.test.js
  NODE_PATH=<folder>/node_modules node tests/compose-dock.test.js
  NODE_PATH=<folder>/node_modules node tests/cleanup.test.js
  node tests/cleanup-apply.test.js
  node tests/handlers.test.js
  node tests/pin-sync.test.js
  ```

- **Version history:** in the git log.

## 8. License

[MIT](LICENSE). The license file is also included in `outlook-mode.xpi`.
