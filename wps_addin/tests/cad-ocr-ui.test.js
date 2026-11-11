"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ocrScript = fs.readFileSync(path.join(__dirname, "..", "cad-ocr-ui.js"), "utf8");
const originalSvg = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="10 20 200 100"><title>庭院图纸</title></svg>';
const previewBase64 = Buffer.from(originalSvg, "utf8").toString("base64");

class Element {
  constructor(tagName) {
    this.tagName = tagName;
    this.children = [];
    this.attributes = new Map();
    this.listeners = new Map();
    this.hidden = false;
    this.checked = false;
    this.value = "";
  }

  addEventListener(name, callback) {
    this.listeners.set(name, callback);
  }

  dispatch(name, event = {}) {
    const callback = this.listeners.get(name);
    assert(callback, "Expected a handler for " + name);
    return callback(event);
  }

  click() {
    return this.dispatch("click");
  }

  setAttribute(name, value) {
    this.attributes.set(name, value);
  }

  append(...children) {
    this.children.push(...children);
  }

  appendChild(child) {
    this.children.push(child);
    return child;
  }

  replaceChildren(...children) {
    this.children = children;
  }

  set innerHTML(value) {
    throw new Error("OCR candidates must be rendered as input values and plain text");
  }
}

class Dialog extends Element {
  constructor() {
    super("dialog");
    this.open = false;
  }

  showModal() {
    assert.equal(this.open, false, "A prior OCR dialog must settle before another opens");
    this.open = true;
  }

  close() {
    this.open = false;
  }
}

function createFixture() {
  const elements = {
    cadOcrChoice: new Dialog(),
    cadOcrResults: new Element("section"),
    cadOcrItems: new Element("div"),
    cadImportPreview: new Element("img"),
    cadOcrEnable: new Element("button"),
    cadOcrSkip: new Element("button"),
    cadOcrAbort: new Element("button")
  };
  const calls = { parsed: [], documents: [], urls: [], revoked: [] };
  const context = {
    console,
    Blob,
    TextDecoder,
    Uint8Array,
    atob,
    document: {
      getElementById: id => elements[id],
      createElement: name => new Element(name)
    },
    DOMParser: class {
      parseFromString(text, mimeType) {
        calls.parsed.push({ text, mimeType });
        const document = {
          documentElement: new Element("svg"),
          createElementNS(namespace, tagName) {
            assert.equal(namespace, "http://www.w3.org/2000/svg");
            return new Element(tagName);
          }
        };
        calls.documents.push(document);
        return document;
      }
    },
    XMLSerializer: class {
      serializeToString(document) {
        return JSON.stringify(document.documentElement.children.map(child => ({
          name: child.tagName,
          attributes: Object.fromEntries(child.attributes)
        })));
      }
    },
    URL: {
      createObjectURL(blob) {
        const url = "blob:cad-ocr-" + (calls.urls.length + 1);
        calls.urls.push({ url, blob });
        return url;
      },
      revokeObjectURL(url) {
        calls.revoked.push(url);
      }
    }
  };
  context.window = context;
  vm.createContext(context);
  vm.runInContext(ocrScript, context, { filename: "dwg-ocr-ui.js" });
  return { elements, calls, api: context.WpsCadOcr };
}

function sampleOcr() {
  return {
    width: 1000,
    height: 500,
    viewBox: [10, 20, 200, 100],
    lines: [
      { text: "庭院入囗", confidence: 0.82, points: [{ x: 10, y: 20 }, { x: 110, y: 20 }, { x: 110, y: 70 }, { x: 10, y: 70 }] },
      { text: "<b>原线条待核对</b>", confidence: 0.64, points: [{ x: 500, y: 200 }, { x: 600, y: 200 }, { x: 600, y: 250 }, { x: 500, y: 250 }] }
    ]
  };
}

function selections(api) {
  // Normalize the VM realm boundary while preserving the actual API payload.
  return JSON.parse(JSON.stringify(api.selections()));
}

async function testExplicitEnableSkipAndAbort() {
  const fixture = createFixture();
  for (const [button, expected] of [["cadOcrEnable", true], ["cadOcrSkip", false], ["cadOcrAbort", null]]) {
    const chosen = fixture.api.choose();
    assert.equal(fixture.elements.cadOcrChoice.open, true);
    fixture.elements[button].click();
    assert.equal(await chosen, expected);
    assert.equal(fixture.elements.cadOcrChoice.open, false);
  }

  let prevented = false;
  const escaped = fixture.api.choose();
  fixture.elements.cadOcrChoice.dispatch("cancel", { preventDefault() { prevented = true; } });
  assert.equal(await escaped, null);
  assert.equal(prevented, true, "Escape must resolve the OCR choice as an abort");
}

async function testCandidatesRequireSelectionAndUseEditedText() {
  const fixture = createFixture();
  fixture.api.show(sampleOcr(), previewBase64);
  assert.equal(fixture.elements.cadOcrResults.hidden, false);
  assert.equal(fixture.elements.cadOcrItems.children.length, 2);
  assert.deepEqual(selections(fixture.api), [], "OCR candidates must never be checked automatically");

  const first = fixture.elements.cadOcrItems.children[0];
  const second = fixture.elements.cadOcrItems.children[1];
  const [check, text, confidence] = first.children;
  assert.equal(check.checked, false);
  assert.equal(second.children[0].checked, false);
  assert.equal(second.children[1].value, "<b>原线条待核对</b>");
  assert.equal(confidence.textContent, "82%");
  text.value = "庭院入口";
  assert.deepEqual(selections(fixture.api), [], "Editing alone must not authorize restoring an OCR region");
  check.checked = true;
  check.dispatch("change");
  assert.deepEqual(selections(fixture.api), [{ index: 0, text: "庭院入口" }]);
  fixture.api.setDisabled(true);
  assert.equal(check.disabled, true, "Insertion must freeze the reviewed OCR choices");
  assert.equal(text.disabled, true, "Insertion must freeze the reviewed OCR content");
  fixture.api.setDisabled(false);
  assert.equal(check.disabled, false);
  assert.equal(text.disabled, false);

  const insertionOptions = JSON.parse(JSON.stringify({ font: "Microsoft YaHei", ocrSelections: fixture.api.selections() }));
  assert.deepEqual(insertionOptions.ocrSelections, [{ index: 0, text: "庭院入口" }], "The host payload must carry the user's corrected content");
}

async function testSelectedRegionsHighlightAndReleaseOldUrls() {
  const fixture = createFixture();
  fixture.api.show(sampleOcr(), previewBase64);
  const firstCheck = fixture.elements.cadOcrItems.children[0].children[0];
  firstCheck.checked = true;
  firstCheck.dispatch("change");

  assert.equal(fixture.calls.parsed[0].text, originalSvg, "SVG decoding must preserve Chinese text");
  assert.equal(fixture.calls.parsed[0].mimeType, "image/svg+xml");
  const firstPolygon = fixture.calls.documents[0].documentElement.children[0];
  assert.equal(firstPolygon.tagName, "polygon");
  assert.equal(firstPolygon.attributes.get("points"), "12,24 32,24 32,34 12,34");
  assert.equal(firstPolygon.attributes.get("fill"), "none");
  assert.equal(firstPolygon.attributes.get("vector-effect"), "non-scaling-stroke");
  assert.equal(fixture.calls.urls[0].blob.type, "image/svg+xml");
  assert.equal(fixture.elements.cadImportPreview.src, fixture.calls.urls[0].url);

  const secondCheck = fixture.elements.cadOcrItems.children[1].children[0];
  secondCheck.checked = true;
  secondCheck.dispatch("change");
  assert.deepEqual(fixture.calls.revoked, [fixture.calls.urls[0].url]);
  assert.equal(fixture.calls.documents[1].documentElement.children.length, 2, "A new overlay must contain exactly the selected regions");
  firstCheck.checked = false;
  firstCheck.dispatch("change");
  assert.equal(fixture.calls.documents[2].documentElement.children.length, 1, "Deselection must remove the old overlay");
  fixture.api.clear();
  assert.equal(fixture.calls.revoked.at(-1), fixture.calls.urls.at(-1).url);
}

async function testClearSettlesDialogAndRemovesPriorCandidates() {
  const fixture = createFixture();
  const firstChoice = fixture.api.choose();
  const replacementChoice = fixture.api.choose();
  assert.equal(await firstChoice, null, "A replaced OCR prompt must not remain pending");
  fixture.api.clear();
  assert.equal(await replacementChoice, null);
  assert.equal(fixture.elements.cadOcrChoice.open, false);

  fixture.api.show(sampleOcr(), previewBase64);
  fixture.elements.cadOcrItems.children[0].children[0].checked = true;
  fixture.api.clear();
  assert.deepEqual(selections(fixture.api), []);
  assert.equal(fixture.elements.cadOcrItems.children.length, 0);
  assert.equal(fixture.elements.cadOcrResults.hidden, true);
  fixture.api.show(sampleOcr(), previewBase64);
  assert.deepEqual(selections(fixture.api), [], "A subsequent drawing must not inherit selected OCR candidates");
  fixture.api.show({ lines: [] }, previewBase64);
  assert.equal(fixture.elements.cadOcrResults.hidden, true);
}

(async function () {
  for (const test of [
    testExplicitEnableSkipAndAbort,
    testCandidatesRequireSelectionAndUseEditedText,
    testSelectedRegionsHighlightAndReleaseOldUrls,
    testClearSettlesDialogAndRemovesPriorCandidates
  ]) {
    await test();
    console.log("PASS " + test.name);
  }
}()).catch(error => {
  console.error(error);
  process.exitCode = 1;
});
