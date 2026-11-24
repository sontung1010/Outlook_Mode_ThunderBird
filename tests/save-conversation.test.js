"use strict";
// Run: NODE_PATH=<dir with jsdom>/node_modules node tests/save-conversation.test.js
// (jsdom: `npm install jsdom@24` in any temporary folder). The test uses
// Thunderbird's own MIME parser (jsmime): set JSMIME=<path to jsmime.mjs>, or
// THUNDERBIRD_OMNI=<path to omni.ja>; common install places are tried, and
// the test is skipped when no Thunderbird is found.
//
// experiments/outlookPane: the reading pane's "save the conversation" button
// saves the conversation shown as ONE .eml file, "<subject> - <today>.eml",
// whose text is the whole conversation, newest first: each message under an
// Outlook-style header, without the earlier mail it quotes (the oldest keeps
// its quote), with the pictures it shows; then the messages' attachments.
// Only messages of that conversation (the selected thread, plus mail linked
// to it by Message-ID / References, e.g. your replies in Sent Items) are in it.
const fs = require("fs");
const os = require("os");
const path = require("path");
const vm = require("vm");
const assert = require("assert");
const { execFileSync } = require("child_process");
const { JSDOM } = require("jsdom");

function findJsmime() {
  if (process.env.JSMIME) {
    return process.env.JSMIME;
  }
  const omnis = [
    process.env.THUNDERBIRD_OMNI,
    "/opt/thunderbird/omni.ja",
    "/usr/lib/thunderbird/omni.ja",
    "/usr/lib64/thunderbird/omni.ja",
    "/snap/thunderbird/current/omni.ja",
    "/Applications/Thunderbird.app/Contents/Resources/omni.ja",
  ].filter(Boolean);
  for (const omni of omnis) {
    if (fs.existsSync(omni)) {
      const out = path.join(os.tmpdir(), `outlookmode-jsmime-${process.pid}.mjs`);
      fs.writeFileSync(out, execFileSync("unzip", ["-p", omni, "modules/jsmime/jsmime.mjs"]));
      process.on("exit", () => fs.rmSync(out, { force: true }));
      return out;
    }
  }
  return null;
}

const code = fs.readFileSync(path.join(__dirname, "../outlook-mode/experiments/outlookPane/implementation.js"), "utf8");
const b64 = (text) => Buffer.from(text, "utf8").toString("base64");
const PNG = (n) => Buffer.from([0x89, 0x50, 0x4e, 0x47, n, n, n]).toString("base64");
const crlf = (lines) => lines.join("\r\n") + "\r\n";

// The conversation: Tung's plan (Sent), Truong's answer (Inbox, Outlook
// desktop quote), Tung's plain-text reply (Sent, Outlook underscore quote),
// Miya's answer (Inbox, new Outlook quote after a line, a forwarded mail
// attached). Plus a mail of another conversation.
const raws = {
  "mailbox:sent#1": crlf([
    "From: Tùng <t@x>", "To: Truong <tr@x>", "Subject: Plan", "Message-ID: <orig@x>",
    'Content-Type: multipart/mixed; boundary="m1"', "",
    "--m1", 'Content-Type: multipart/related; boundary="r1"', "",
    "--r1", "Content-Type: text/html; charset=UTF-8", "Content-Transfer-Encoding: base64", "",
    b64('<html><head><style>p.Mso{margin:0}</style></head><body><p class="Mso">Kế hoạch gốc</p><img src="cid:logo@x"></body></html>'),
    "--r1", "Content-Type: image/png", "Content-ID: <logo@x>", "Content-Transfer-Encoding: base64", "", PNG(1),
    "--r1--",
    "--m1", 'Content-Type: application/pdf; name="plan.pdf"', 'Content-Disposition: attachment; filename="plan.pdf"', "Content-Transfer-Encoding: base64", "", b64("%PDF plan"),
    "--m1--",
  ]),
  "mailbox:inbox#2": crlf([
    "From: Truong <tr@x>", "To: Tùng <t@x>", "Cc: Ha <ha@x>", "Subject: RE: Plan", "Message-ID: <ans@y>", "References: <orig@x>",
    'Content-Type: multipart/related; boundary="r2"', "",
    "--r2", "Content-Type: text/html; charset=UTF-8", "Content-Transfer-Encoding: base64", "",
    b64(
      '<p>Đồng ý với kế hoạch</p><img src="cid:sig@y" onerror="alert(1)"><script>bad()</script>' +
        '<div style="border:none;border-top:solid #E1E1E1 1.0pt"><p><b>From:</b> Tùng<br><b>Sent:</b> Monday</p></div>' +
        '<p>Kế hoạch gốc (quoted)</p><img src="cid:logo2@y">'
    ),
    "--r2", "Content-Type: image/png", "Content-ID: <sig@y>", "Content-Transfer-Encoding: base64", "", PNG(2),
    "--r2", "Content-Type: image/png", "Content-ID: <logo2@y>", "Content-Transfer-Encoding: base64", "", PNG(3),
    "--r2", 'Content-Type: application/pdf; name="plan.pdf"', 'Content-Disposition: attachment; filename="plan.pdf"', "Content-Transfer-Encoding: base64", "", b64("%PDF plan"),
    "--r2--",
  ]),
  "mailbox:sent#3": crlf([
    "From: Tùng <t@x>", "To: Truong <tr@x>", "Subject: RE: Plan", "Message-ID: <mine@x>", "References: <orig@x> <ans@y>",
    "Content-Type: text/plain; charset=UTF-8", "Content-Transfer-Encoding: base64", "",
    b64("Cảm ơn anh <ok>\r\n\r\n________________________________\r\nFrom: Truong\r\nSent: Monday\r\n\r\nĐồng ý (quoted)\r\n"),
  ]),
  "mailbox:inbox#5": crlf([
    "From: Miya <m@q>", "To: Tùng <t@x>", "Subject: RE: Plan", "Message-ID: <miya@q>", "References: <orig@x> <ans@y> <mine@x>",
    'Content-Type: multipart/mixed; boundary="m5"', "",
    "--m5", "Content-Type: text/html; charset=UTF-8", "Content-Transfer-Encoding: base64", "",
    b64('<div>Miya agrees</div><hr style="display:inline-block;width:98%"><div><b>差出人:</b> Tùng<br><b>件名:</b> RE: Plan</div><div>Cảm ơn anh (quoted)</div>'),
    "--m5", "Content-Type: message/rfc822", "Content-Disposition: attachment", "",
    "Subject: Shipping: address?", "From: a@b", "", "Forwarded body",
    "--m5--",
  ]),
  "mailbox:inbox#4": crlf(["From: Eve <e@z>", "Subject: Other", "Message-ID: <other@z>", "", "Elsewhere"]),
};

function hdr(key, uri, messageId, refs, date, author, subject, flags = 0) {
  const [folderName] = uri.split("#");
  return {
    key, uri, messageId, mime2DecodedAuthor: author, mime2DecodedSubject: subject, flags,
    date: Date.parse(date) * 1000, numReferences: refs.length, getStringReference: (i) => refs[i],
    folder: { name: folderName, server: { key: "server2" }, getUriForMsg: () => uri },
  };
}
const HAS_RE = 0x10;
const all = [
  hdr(1, "mailbox:sent#1", "orig@x", [], "2026-11-20T11:38:00", "Tùng <t@x>", "Plan"),
  hdr(2, "mailbox:inbox#2", "ans@y", ["orig@x"], "2026-11-20T11:49:00", "Truong <tr@x>", "Plan", HAS_RE),
  hdr(3, "mailbox:sent#3", "mine@x", ["orig@x", "ans@y"], "2026-11-20T18:20:00", "Tùng <t@x>", "Plan", HAS_RE),
  hdr(5, "mailbox:inbox#5", "miya@q", ["orig@x", "ans@y", "mine@x"], "2026-11-21T09:00:00", "Miya <m@q>", "Plan", HAS_RE),
  hdr(4, "mailbox:inbox#4", "other@z", [], "2026-11-20T09:00:00", "Eve <e@z>", "Other"),
];
const [original, answer] = all;
const byUri = new Map(all.map((h) => [h.uri, h]));
const convIdToHdr = new Map(all.map((h) => [10 + h.key, h]));

(async () => {
  const jsmimePath = findJsmime();
  if (!jsmimePath) {
    console.log("save-conversation: skipped (no Thunderbird omni.ja found; set JSMIME or THUNDERBIRD_OMNI)");
    return;
  }
  const { jsmime } = await import("file://" + jsmimePath);
  const MimeParser = {
    parseSync(input, emitter, opts) {
      const parser = new jsmime.MimeParser(emitter, opts);
      parser.deliverData(input);
      parser.deliverEOF();
    },
    extractHeaders(input) {
      let headers = null;
      MimeParser.parseSync(input, { startPart: (n, h) => n === "" && (headers = h) }, { pruneat: "", bodyformat: "none" });
      return headers;
    },
    getParameter(header, name) {
      return jsmime.headerparser.parseParameterHeader(header, true, true).get(name.toLowerCase()) || "";
    },
  };
  const window = new JSDOM("").window;

  // Fake files and picker.
  const files = new Map();
  let pickerAnswer = { result: 0, path: "/home/me/Saved/Plan - 2026-11-24.eml" };
  const pickers = [];
  const prefs = new Map();
  const FilePicker = { modeSave: 1, returnOK: 0, returnCancel: 1, returnReplace: 2, filterAll: 1 };
  const deep = () =>
    new Proxy(function () {}, { get: (t, k) => (k === Symbol.toPrimitive ? () => 0 : deep()), apply: () => deep(), construct: () => deep() });
  const modules = {
    "resource:///modules/MailServices.sys.mjs": {
      MailServices: {
        messageServiceFromURI: () => ({ messageURIToMsgHdr: (uri) => byUri.get(uri) || null }),
        headerParser: { makeMimeAddress: (name, email) => `${name} <${email}>` },
        accounts: { getFirstIdentityForServer: () => ({ fullName: "Đỗ Sơn Tùng", email: "me@x" }) },
      },
    },
    "resource://gre/modules/ExtensionParent.sys.mjs": {
      ExtensionParent: {
        GlobalManager: {
          getExtension: () => ({ messageManager: { get: (id) => convIdToHdr.get(id), convert: (h) => ({ id: 10 + h.key }) } }),
        },
      },
    },
    "resource://gre/modules/Timer.sys.mjs": { setTimeout, clearTimeout },
    "resource:///modules/mimeParser.sys.mjs": { MimeParser },
    "moz-src:///comm/mail/modules/MessageSaver.sys.mjs": {
      MessageSaver: {
        saveToFile: async (uri, filePath, type) => {
          assert.equal(type, "eml");
          assert.ok(files.has(filePath), "into the temporary file created for it");
          files.set(filePath, Buffer.from(raws[uri], "utf8"));
        },
      },
    },
  };
  const ctx = {
    ChromeUtils: { importESModule: (url) => modules[url] || new Proxy({}, { get: () => deep() }), generateQI: () => () => {} },
    Ci: new Proxy(
      { nsIFilePicker: FilePicker, nsIFile: { NORMAL_FILE_TYPE: 0 }, nsMsgMessageFlags: { HasRe: HAS_RE } },
      { get: (t, k) => t[k] || deep() }
    ),
    Cc: new Proxy({}, {
      get: (t, k) =>
        k === "@mozilla.org/filepicker;1"
          ? {
              createInstance: () => ({
                filters: [],
                init(bc, title, mode) { Object.assign(this, { title, mode }); },
                appendFilter(label, ext) { this.filters.push(ext); },
                appendFilters() {},
                open(done) {
                  pickers.push(this);
                  this.file = pickerAnswer.result === FilePicker.returnCancel ? null : { path: pickerAnswer.path, parent: { path: path.dirname(pickerAnswer.path) } };
                  done(pickerAnswer.result);
                },
              }),
            }
          : deep(),
    }),
    Services: {
      prefs: {
        getComplexValue: (name) => {
          if (!prefs.has(name)) {
            throw new Error("not set");
          }
          return prefs.get(name);
        },
        setComplexValue: (name, type, value) => prefs.set(name, value),
        getStringPref: () => "",
        getBoolPref: () => false,
      },
      wm: { getMostRecentWindow: () => ({ btoa: (s) => Buffer.from(s, "binary").toString("base64"), DOMParser: window.DOMParser }) },
      dirsvc: {
        get: (key) => {
          assert.equal(key, "TmpD");
          let leaf = "";
          return {
            append(name) { leaf = name; },
            createUnique() {
              for (let n = 1; files.has(`/tmp/${leaf}`); n++) {
                leaf = `outlookmode-message-${n}.eml`;
              }
              files.set(`/tmp/${leaf}`, Buffer.alloc(0));
            },
            get path() { return `/tmp/${leaf}`; },
          };
        },
      },
    },
    IOUtils: {
      read: async (p) => new Uint8Array(files.get(p)),
      write: async (p, bytes) => files.set(p, Buffer.from(bytes)),
      remove: async (p) => files.delete(p),
    },
    ExtensionCommon: { ExtensionAPI: class {} },
    // Like the real experiment scope: no TextEncoder, no PathUtils.
    console, Date, Math, JSON, Promise, Map, Set, Object, Array, String, Number, Boolean, Error, URL, Uint8Array,
    encodeURIComponent, decodeURIComponent,
  };
  vm.createContext(ctx);
  vm.runInContext(code, ctx);
  vm.runInContext(
    "this.api = { paneBridge, paneRelated, paneThreadSubject, paneConversationFileName, paneEncodeHeader, paneUtf8, paneOwnText };",
    ctx
  );
  const { api } = ctx;

  // Helpers on their own.
  assert.ok(api.paneRelated(all[2], [answer]) && api.paneRelated(original, [answer]) && !api.paneRelated(all[4], [answer]));
  assert.equal(api.paneThreadSubject([{ mime2DecodedSubject: "Re: FW: Trả lời: [2] Plan" }]), "[2] Plan");
  assert.equal(api.paneConversationFileName('A/B: "x"?', new Date(2026, 10, 24, 9, 5)), "A_B_ _x_ - 2026-11-24.eml");
  assert.equal(api.paneEncodeHeader("Plain"), "Plain");
  assert.match(api.paneEncodeHeader("Đề xuất"), /^=\?UTF-8\?B\?[A-Za-z0-9+/=]+\?=$/);
  for (const text of ["abc", "Đỗ Sơn Tùng", "返信 件名", "emoji 😀 ok", "lone \ud800 surrogate"]) {
    assert.deepEqual([...api.paneUtf8(text)], [...new TextEncoder().encode(text)], `UTF-8 of ${text}`);
  }
  assert.equal(api.paneOwnText("Answer\n\nOn Mon, Ann wrote:\n> old\n>\n> more\n"), "Answer");
  assert.equal(api.paneOwnText("Line > not a quote\nend"), "Line > not a quote\nend");
  assert.equal(api.paneOwnText("> only a quote\n"), "> only a quote\n", "nothing but a quote stays");

  // The page shows the Inbox thread of `answer` (selected in the list).
  const about3Pane = { browsingContext: {} };
  const bridge = Object.create(api.paneBridge, {
    page: { value: { browsingContext: { embedderElement: { ownerDocument: { defaultView: about3Pane } } } } },
  });
  bridge.shownConversation = () => [answer];

  const file = "/home/me/Saved/Plan - 2026-11-24.eml";
  const result = await bridge.saveConversation([13, 12, 14, 11, 15, 12]);
  assert.deepEqual(result, { saved: 4, file }, "one file, four messages");
  assert.equal(pickers[0].mode, FilePicker.modeSave);
  assert.match(pickers[0].defaultString, /^Plan - \d{4}-\d{2}-\d{2}\.eml$/, "subject + today");
  assert.deepEqual(pickers[0].filters, ["*.eml"]);
  assert.equal(prefs.get("messenger.save.dir").path, "/home/me/Saved", "folder remembered");
  assert.deepEqual([...files.keys()], [file], "temporary copies removed");

  // Read the file back with Thunderbird's parser.
  const out = files.get(file);
  const parts = new Map();
  let top = null;
  MimeParser.parseSync(
    out.toString("latin1"),
    {
      startPart(n, h) {
        if (n === "") {
          top = h;
        }
        parts.set(n, { type: h.contentType.type, headers: h, data: [] });
      },
      deliverPartData(n, d) {
        parts.get(n).data.push(d);
      },
    },
    { bodyformat: "decode", strformat: "unicode", decodeSubMessages: false }
  );
  assert.equal(top.contentType.type, "multipart/mixed");
  assert.equal(top.get("Subject"), "Plan");
  assert.deepEqual(top.get("From").map((a) => [a.name, a.email]), [["Đỗ Sơn Tùng", "me@x"]], "From: the account, only its name encoded");
  const html = parts.get("1.1").data.join("");
  assert.equal(parts.get("1").type, "multipart/related");
  assert.equal(parts.get("1.1").type, "text/html");

  // Newest first, each under its header.
  const order = ["Miya agrees", "Cảm ơn anh &lt;ok&gt;", "Đồng ý với kế hoạch", "Kế hoạch gốc"].map((t) => html.indexOf(t));
  assert.ok(order.every((at) => at > 0), `all messages in the text: ${order}`);
  assert.deepEqual([...order].sort((a, b) => a - b), order, "newest first");
  assert.match(html, /<b>From:<\/b> Miya &lt;m@q&gt;<br><b>Sent:<\/b> [^<]+<br><b>To:<\/b> Tùng &lt;t@x&gt;<br><b>Subject:<\/b> RE: Plan/);
  assert.match(html, /<b>To:<\/b> Tùng &lt;t@x&gt;<br><b>Cc:<\/b> Ha &lt;ha@x&gt;/, "Cc line");
  assert.match(html, /4 messages, newest first/);

  // The earlier mail each quotes is left out; the oldest keeps everything.
  for (const quoted of ["Kế hoạch gốc (quoted)", "Đồng ý (quoted)", "Cảm ơn anh (quoted)", "差出人"]) {
    assert.ok(!html.includes(quoted), `no ${quoted}`);
  }
  assert.ok(!html.includes("<script") && !html.includes("onerror") && !html.includes("bad()"), "no scripts or handlers");
  assert.ok(html.includes("p.Mso{margin:0}"), "the messages' styles kept");

  // Pictures: those shown, renamed per message; the quoted one left out.
  const pictures = [...parts.values()].filter((p) => p.type === "image/png");
  const pictureIds = pictures.map((p) => String(p.headers.get("Content-ID"))).sort();
  assert.deepEqual(pictureIds, ["<m1-logo@x>", "<m2-sig@y>"]);
  assert.ok(html.includes('src="cid:m1-logo@x"') && html.includes('src="cid:m2-sig@y"'));
  assert.ok(!html.includes("logo2@y"));

  // Attachments: the same file once, the attached mail named by its subject.
  const attachments = [...parts.entries()].filter(([n]) => /^\d+$/.test(n) && n !== "1").map(([, p]) => p);
  const names = attachments.map((p) => MimeParser.getParameter(String(p.headers.get("Content-Disposition")), "filename"));
  assert.deepEqual(names.sort(), ["Shipping_ address_.eml", "plan.pdf"]);
  const forwarded = attachments.find((p) => p.type === "message/rfc822");
  assert.match(String(forwarded.headers.get("Content-Transfer-Encoding")), /^7bit$/, "an attached mail is not base64 (MIME allows none)");
  assert.match(Buffer.concat(forwarded.data.map((d) => Buffer.from(d))).toString("latin1"), /^Subject: Shipping: address\?\r\nFrom: a@b\r\n\r\nForwarded body/);

  // Cancelled in the picker: nothing read, nothing written.
  pickerAnswer = { result: FilePicker.returnCancel, path: "" };
  files.clear();
  assert.deepEqual(await bridge.saveConversation([12]), { saved: 0, file: "" });
  assert.equal(files.size, 0);

  // ".eml" added when the name typed has none.
  pickerAnswer = { result: FilePicker.returnOK, path: "/home/me/Saved/thread" };
  assert.deepEqual(await bridge.saveConversation([12]), { saved: 1, file: "/home/me/Saved/thread.eml" });

  // Nothing shown (another tab, several rows selected): nothing saved.
  bridge.shownConversation = () => null;
  const before = pickers.length;
  assert.deepEqual(await bridge.saveConversation([12]), { saved: 0, file: "" });
  assert.equal(pickers.length, before, "not even asked");

  console.log("save-conversation: all checks passed");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
