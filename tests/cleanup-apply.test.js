"use strict";
// Run: node tests/cleanup-apply.test.js
// Runs cleanupApply of background/cleanup.js against a fake Thunderbird
// whose deletes may never report back (as when Microsoft 365 slows down
// many changes in a row), with short timeouts, and checks that the cleanup
// goes on, retries, stops when nothing moves any more, and reports only
// what was really done.
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const assert = require("assert");

const original = fs.readFileSync(path.join(__dirname, "../outlook-mode/background/cleanup.js"), "utf8");
assert.ok(original.includes("const CLEANUP_POLL = 1500;") && original.includes("const CLEANUP_STALL = 30e3;"));
const sourceWith = (stall) =>
  original.replace("const CLEANUP_POLL = 1500;", "const CLEANUP_POLL = 10;").replace("const CLEANUP_STALL = 30e3;", `const CLEANUP_STALL = ${stall};`);

const PIN = "$ol_pinned";

// behaviour(batchNumber, ids) → "ok", "hang" (does the work, never reports
// back), "hang-nothing" (never reports back, does nothing), "hang-half".
function setup(count, behaviour, { pinned = [], stall = 60 } = {}) {
  const store = new Map();
  for (let id = 1; id <= count; id++) {
    store.set(id, { id, folder: { id: "inbox", specialUse: ["inbox"] }, tags: pinned.includes(id) ? [PIN] : [] });
  }
  const progress = [];
  let batch = 0;
  const context = {
    console,
    setTimeout,
    clearTimeout,
    handlers: {},
    PIN_TAG_KEY: PIN,
    logError() {},
    fromOwnPage: (handler) => handler,
    userError: (text) => Object.assign(new Error(text), { userFacing: true }),
    messenger: {
      runtime: {
        async sendMessage(message) {
          progress.push(message);
        },
      },
      messages: {
        async get(id) {
          const message = store.get(id);
          if (!message) {
            throw new Error(`Message not found: ${id}.`);
          }
          return structuredClone(message);
        },
        delete(ids) {
          const how = behaviour(++batch, ids);
          const work = how === "hang-nothing" ? [] : how === "hang-half" ? ids.slice(0, ids.length / 2) : ids;
          for (const id of work) {
            store.delete(id);
          }
          return how === "ok" ? Promise.resolve() : new Promise(() => {});
        },
      },
    },
  };
  vm.createContext(context);
  vm.runInContext(sourceWith(stall), context);
  // As if the page had scanned this mail (apply only acts on scanned mail).
  vm.runInContext(`cleanupScanned = new Set(${JSON.stringify([...store.keys()])});`, context);
  return { context, store, progress, batches: () => batch };
}

(async () => {
  // 1. Everything answers: all done, progress up to 100%.
  {
    const t = setup(60, () => "ok");
    const result = await t.context.cleanupApply({ ids: [...t.store.keys()], action: "delete" });
    assert.equal(result.doneIds.length, 60);
    assert.equal(result.notDone, 0);
    assert.ok(result.ok);
    assert.equal(t.store.size, 0);
    assert.deepEqual(t.progress.map((p) => p.handled), [0, 25, 50, 60]);
  }

  // 2. The second batch never reports back but was done: seen at once by
  //    watching the mail, no waiting for the stall time.
  {
    const t = setup(60, (n) => (n === 2 ? "hang" : "ok"), { stall: 10e3 });
    const start = Date.now();
    const result = await t.context.cleanupApply({ ids: [...t.store.keys()], action: "delete" });
    assert.ok(Date.now() - start < 2e3, "did not wait for the stall time");
    assert.equal(result.doneIds.length, 60);
    assert.equal(result.notDone, 0);
  }

  // 3. Half of a batch was done before it hung: the other half is tried
  //    again at the end, and done then.
  {
    const t = setup(60, (n) => (n === 1 ? "hang-half" : "ok"));
    const result = await t.context.cleanupApply({ ids: [...t.store.keys()], action: "delete" });
    assert.equal(t.batches(), 4, "3 batches + 1 retry");
    assert.equal(result.doneIds.length, 60);
    assert.equal(result.notDone, 0);
    assert.equal(t.store.size, 0);
  }

  // 4. The server does not answer any more: stops after 2 batches with
  //    nothing done, and says how many are left.
  {
    const t = setup(100, (n) => (n === 1 ? "ok" : "hang-nothing"));
    const result = await t.context.cleanupApply({ ids: [...t.store.keys()], action: "delete" });
    assert.equal(t.batches(), 3);
    assert.ok(result.stalled);
    assert.ok(!result.stopped);
    assert.equal(result.doneIds.length, 25);
    assert.equal(result.notDone, 75);
    assert.equal(t.store.size, 75);
  }

  // 5. Pinned mail is left alone, also when pinned after the scan.
  {
    const t = setup(30, () => "ok", { pinned: [3, 27] });
    const result = await t.context.cleanupApply({ ids: [...t.store.keys()], action: "delete" });
    assert.equal(result.skipped, 2);
    assert.equal(result.doneIds.length, 28);
    assert.deepEqual([...t.store.keys()], [3, 27]);
  }

  // 6. Stop: ends after the batch running, the rest is not touched.
  {
    let t;
    t = setup(100, (n) => {
      if (n === 2) {
        t.context.handlers["cleanup-stop"]();
      }
      return "ok";
    });
    const result = await t.context.cleanupApply({ ids: [...t.store.keys()], action: "delete" });
    assert.ok(result.stopped);
    assert.equal(result.doneIds.length, 50);
    assert.equal(result.notDone, 50);
    assert.equal(t.store.size, 50);
  }

  // 7. Stop while a batch hangs: answered at once, not after the stall time.
  {
    const t = setup(60, () => "hang-nothing", { stall: 10e3 });
    setTimeout(() => t.context.handlers["cleanup-stop"](), 50);
    const start = Date.now();
    const result = await t.context.cleanupApply({ ids: [...t.store.keys()], action: "delete" });
    assert.ok(Date.now() - start < 1e3, `stopped in ${Date.now() - start} ms`);
    assert.ok(result.stopped);
    assert.equal(result.doneIds.length, 0);
    assert.equal(result.notDone, 60);
  }

  // 8. Progress never goes back, also while trying again.
  {
    const t = setup(60, (n) => (n === 1 ? "hang-half" : "ok"));
    await t.context.cleanupApply({ ids: [...t.store.keys()], action: "delete" });
    const handled = t.progress.map((p) => p.handled);
    assert.deepEqual(handled, [...handled].sort((a, b) => a - b), handled.join(","));
    assert.equal(handled.at(-1), 60);
  }

  // 9. Mail the last scan did not list is refused, as a whole.
  {
    const t = setup(10, () => "ok");
    await assert.rejects(t.context.cleanupApply({ ids: [1, 2, 99], action: "delete" }), /not in the last scan/);
    assert.equal(t.store.size, 10, "nothing deleted");
    await assert.rejects(t.context.cleanupApply({ ids: [1], action: "format" }), /Invalid cleanup request/);
  }

  console.log("all cleanup-apply scenarios passed");
})().catch((error) => {
  console.error(error);
  process.exit(1);
});
