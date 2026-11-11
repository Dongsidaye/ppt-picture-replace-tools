"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const importScript = fs.readFileSync(path.join(__dirname, "..", "web", "dwg-import.js"), "utf8");
const firstDocument = "document:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa:window:101";
const secondDocument = "document:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb:window:202";
const svgPreview = Buffer.from(
  '<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100"><path d="M0 0L100 100"/></svg>'
).toString("base64");

function delay(milliseconds) {
  return new Promise(resolve => setTimeout(resolve, milliseconds));
}

async function waitFor(condition, description) {
  const deadline = Date.now() + 3000;
  while (!condition()) {
    if (Date.now() >= deadline) throw new Error("Timed out waiting for " + description);
    await delay(10);
  }
}

class FakeElement {
  constructor() {
    this.hidden = false;
    this.disabled = false;
    this.textContent = "";
    this.listeners = new Map();
    this.attributes = new Map();
  }

  addEventListener(name, callback) {
    this.listeners.set(name, callback);
  }

  click() {
    const callback = this.listeners.get("click");
    assert(callback, "The normal button entry point must have a click handler");
    // Deliver even a queued second click so the handler's own guard is tested.
    return callback();
  }

  removeAttribute(name) {
    this.attributes.delete(name);
  }

  set innerHTML(value) {
    throw new Error("CAD content must not be assigned through innerHTML");
  }
}

class FakePreview extends FakeElement {
  set src(value) {
    this.attributes.set("src", value);
    this.naturalWidth = 100;
    this.naturalHeight = 100;
    queueMicrotask(() => {
      if (this.onload) this.onload();
    });
  }

  get src() {
    return this.attributes.get("src") || "";
  }
}

function createFixture() {
  const elements = {};
  for (const name of ["Choose", "Cancel", "Insert", "File", "Status", "Warning", "PreviewWrap"]) {
    elements["cadImport" + name] = new FakeElement();
  }
  elements.cadImportPreview = new FakePreview();

  const events = new Map();
  const calls = { pick: 0, begin: [], poll: [], cancel: [], insert: [] };
  let documentKey = firstDocument;
  let resolvePicker = null;
  let pollResult = { state: "pending", status: "Converting model space" };
  const fixture = {
    elements,
    calls,
    complete: {
      state: "done",
      ok: true,
      previewBase64: svgPreview,
      warning: "<b>DWG warning remains plain text</b>",
      entityCount: 7
    },
    selectFile(result) {
      assert(resolvePicker, "A native file picker must be pending");
      const resolve = resolvePicker;
      resolvePicker = null;
      resolve(result);
    },
    setPollResult(result) {
      pollResult = result;
    },
    changeDocument(nextDocument) {
      const previous = documentKey;
      documentKey = nextDocument;
      for (const callback of events.get("ppt-document-changed") || []) {
        callback({ detail: { previous, current: nextDocument } });
      }
    },
    dispose() {
      this.changeDocument("");
      if (resolvePicker) this.selectFile({ ok: false, cancelled: true });
    }
  };

  const context = {
    console,
    setTimeout,
    clearTimeout,
    location: { origin: "https://ppt-toolbox.local" },
    document: { getElementById: id => elements[id] },
    addEventListener(name, callback) {
      if (!events.has(name)) events.set(name, []);
      events.get(name).push(callback);
    },
    __pptDocumentKey: () => documentKey,
    __pptChooseFile(options) {
      calls.pick += 1;
      assert.equal(options.kind, "file");
      assert.equal(options.filter, "*.dwg");
      return new Promise(resolve => { resolvePicker = resolve; });
    },
    __pptHost: {
      BeginDwgImport(options) {
        const request = JSON.parse(options);
        assert.equal(request.documentKey, documentKey);
        calls.begin.push(request);
        return "job-" + calls.begin.length;
      },
      PollDwgImport(id) {
        calls.poll.push(id);
        return JSON.stringify(pollResult);
      },
      CancelDwgImport(id) {
        calls.cancel.push(id);
        return true;
      },
      InsertDwgImport(id) {
        calls.insert.push(id);
        return JSON.stringify({ ok: true });
      }
    },
    wps: {
      CreateTaskPane(url) {
        fixture.paneUrl = url;
        fixture.pane = { Visible: false };
        return fixture.pane;
      }
    }
  };
  context.window = context;
  vm.createContext(context);
  vm.runInContext(importScript, context, { filename: "dwg-import.js" });
  fixture.openPane = () => context.OpenCadImportPane();
  return fixture;
}

async function withFixture(work) {
  const fixture = createFixture();
  try {
    await work(fixture);
  } finally {
    fixture.dispose();
  }
}

async function testNormalEntryAndDuplicateClicks() {
  await withFixture(async fixture => {
    const { elements, calls } = fixture;
    fixture.openPane();
    assert.equal(fixture.paneUrl, "https://ppt-toolbox.local/taskpane.html#tools/cad");
    assert.equal(fixture.pane.Visible, true);

    const choosing = elements.cadImportChoose.click();
    await elements.cadImportChoose.click();
    assert.equal(calls.pick, 1, "A second click must not open another file picker");
    fixture.selectFile({ ok: true, path: "C:/CAD/example.dwg" });
    await waitFor(() => calls.begin.length === 1, "the conversion to start");
    fixture.setPollResult(fixture.complete);
    await waitFor(() => !elements.cadImportInsert.disabled, "the converted preview");
    await choosing;

    assert.equal(elements.cadImportFile.textContent, "example.dwg");
    assert.equal(elements.cadImportPreview.src, "data:image/svg+xml;base64," + svgPreview);
    assert.equal(elements.cadImportPreviewWrap.hidden, false);
    assert.equal(elements.cadImportWarning.textContent, fixture.complete.warning);
    assert.equal(elements.cadImportWarning.hidden, false);

    const inserting = elements.cadImportInsert.click();
    await elements.cadImportInsert.click();
    await inserting;
    assert.deepEqual(calls.insert, ["job-1"], "A converted drawing must be inserted only once");
    assert.equal(elements.cadImportInsert.disabled, true);
    assert.equal(elements.cadImportChoose.disabled, false);
  });
}

async function testCancellationWaitsForCompletion() {
  await withFixture(async fixture => {
    const { elements, calls } = fixture;
    const choosing = elements.cadImportChoose.click();
    fixture.selectFile({ ok: true, path: "C:/CAD/cancel.dwg" });
    await waitFor(() => calls.begin.length === 1, "the cancellable conversion");

    elements.cadImportCancel.click();
    assert.deepEqual(calls.cancel, ["job-1"]);
    assert.equal(elements.cadImportChoose.disabled, true, "Wait for the cancelled worker before accepting a retry");
    assert.equal(elements.cadImportInsert.disabled, true);
    fixture.setPollResult({ state: "done", ok: false, cancelled: true });
    await choosing;

    assert.equal(elements.cadImportChoose.disabled, false);
    assert.equal(elements.cadImportCancel.hidden, true);
    assert.equal(elements.cadImportPreviewWrap.hidden, true);
    assert.equal(calls.insert.length, 0);
  });
}

async function testDocumentSwitchSuppressesLateResult() {
  await withFixture(async fixture => {
    const { elements, calls } = fixture;
    const choosing = elements.cadImportChoose.click();
    fixture.selectFile({ ok: true, path: "C:/CAD/previous-document.dwg" });
    await waitFor(() => calls.begin.length === 1, "the original document conversion");

    fixture.changeDocument(secondDocument);
    fixture.setPollResult(fixture.complete);
    await choosing;

    assert.deepEqual(calls.cancel, ["job-1"]);
    assert.equal(elements.cadImportChoose.disabled, false);
    assert.equal(elements.cadImportInsert.disabled, true);
    assert.equal(elements.cadImportPreviewWrap.hidden, true);
    assert.equal(elements.cadImportPreview.src, "");
    assert.equal(calls.insert.length, 0, "An old conversion must not insert into the newly active document");
  });
}

async function testStalePickerCannotStartConversion() {
  await withFixture(async fixture => {
    const { elements, calls } = fixture;
    const choosing = elements.cadImportChoose.click();
    fixture.changeDocument(secondDocument);
    fixture.selectFile({ ok: true, path: "C:/CAD/stale-picker.dwg" });
    await choosing;

    assert.equal(calls.begin.length, 0, "A file chosen for the previous document must be discarded");
    assert.equal(elements.cadImportChoose.disabled, false);
    assert.equal(elements.cadImportInsert.disabled, true);
  });
}

async function testMissingDocumentKeepsControlsUsable() {
  await withFixture(async fixture => {
    const { elements, calls } = fixture;
    fixture.changeDocument("");
    await elements.cadImportChoose.click();

    assert.equal(calls.pick, 0, "A missing document must be handled before opening the file picker");
    assert.equal(calls.begin.length, 0);
    assert.equal(elements.cadImportChoose.disabled, false);
    assert.equal(elements.cadImportInsert.disabled, true);
    assert(elements.cadImportStatus.textContent.length > 0, "Explain why the operation cannot start");
  });
}

(async function () {
  const tests = [
    testNormalEntryAndDuplicateClicks,
    testCancellationWaitsForCompletion,
    testDocumentSwitchSuppressesLateResult,
    testStalePickerCannotStartConversion,
    testMissingDocumentKeepsControlsUsable
  ];
  for (const test of tests) {
    await test();
    console.log("PASS " + test.name);
  }
}()).catch(error => {
  console.error(error);
  process.exitCode = 1;
});
