"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const root = path.resolve(__dirname, "..");
const main = fs.readFileSync(path.join(root, "shared/main.js"), "utf8");
const html = fs.readFileSync(path.join(root, "shared/taskpane.html"), "utf8");
const adapter = fs.readFileSync(path.join(root, "web/host-adapter.js"), "utf8");
const bootstrap = fs.readFileSync(path.join(root, "web/host-bootstrap.js"), "utf8");
const actions = JSON.parse(fs.readFileSync(path.join(root, "shared/action-map.json"), "utf8"));

function environment() {
  const events = new Map(), timers = new Map(), files = new Map();
  let timerId = 0;
  const native = { ActiveWindow: { HWND: 101 }, ActivePresentation: { Name: "A" } };
  const host = {
    TempFolder: "C:\\port-tests", shown: 0, hidden: 0,
    FileExists: p => files.has(p), ReadBinary: p => files.get(p), WriteBinary: (p, v) => files.set(p, v),
    ReadText: p => files.get(p), WriteText: (p, v) => files.set(p, v), RemoveFile: p => files.delete(p),
    ReadBinaryBase64: p => Buffer.from(files.get(p), "latin1").toString("base64"), WriteBinaryBase64: (p, v) => files.set(p, Buffer.from(v, "base64").toString("latin1")),
    ShowPane() { this.shown++; }, HidePane() { this.hidden++; }, Log() {}
  };
  host.ReadMember = (target, name) => { if (typeof target[name] === "function") throw new Error("method"); return target[name]; };
  host.WriteMember = (target, name, value) => { target[name] = value; };
  const context = {
    chrome: { webview: { hostObjects: { sync: { pptHost: host, pptApplication: native } } } },
    location: { href: "https://ppt-toolbox.local/taskpane.html#idle", origin: "https://ppt-toolbox.local", hash: "#idle" },
    URL, console, atob, btoa, alert() {}, confirm: () => false,
    Event: class { constructor(type) { this.type = type; } },
    CustomEvent: class { constructor(type, options) { this.type = type; this.detail = options.detail; } },
    setTimeout(fn) { timers.set(++timerId, fn); return timerId; },
    clearTimeout(id) { timers.delete(id); },
    addEventListener(name, fn) { if (!events.has(name)) events.set(name, []); events.get(name).push(fn); },
    dispatchEvent(event) { (events.get(event.type) || []).forEach(fn => fn(event)); },
    fetch: async () => ({ ok: true, json: async () => actions })
  };
  context.window = context;
  vm.createContext(context);
  vm.runInContext(adapter, context);
  return { context, native, host, timers, files, events };
}
function engine(env, fixture = "") {
  vm.runInContext(main.replace("  global.WpsPictureReplace = {", fixture + "\n  global.WpsPictureReplace = {"), env.context);
}
function extract(start, end) {
  const at = html.indexOf(start), until = html.indexOf(end, at);
  assert(at >= 0 && until > at, "Expected generated UI boundary");
  return html.slice(at, until);
}
function classList() {
  const values = new Set();
  return { add: v => values.add(v), remove: v => values.delete(v), contains: v => values.has(v) };
}
async function flush() { for (let i = 0; i < 12; i++) await Promise.resolve(); }

async function testObjectOwnership() {
  const e = environment();
  e.context.fixtureShape = { Name: "Picture A", Type: 13, Id: 7, Visible: -1, Left: 0, Top: 0, Width: 100, Height: 80, Locked: false };
  engine(e, `
    layerEnsureSelectionGuard = function () { return false; };
    layerCurrentContext = function () { return {kind:"slide",slideIndex:1,label:"A",container:{}}; };
    objectFilterShapes = function () { return [global.fixtureShape]; };
  `);
  const W = e.context.WpsPictureReplace;
  const row = W.layerList().items[0];
  assert.equal(row.ownerDocumentKey, "window:101");
  assert.equal(W.layerSetVisible(row, false).ok, true);
  assert.equal(e.context.fixtureShape.Visible, 0);
  e.native.ActiveWindow.HWND = 202;
  assert.equal(W.layerSetVisible(row, true).ok, false);
  assert.equal(W.layerSetLocked(row, true).ok, false);
  assert.equal(e.context.fixtureShape.Visible, 0, "B must not mutate an A row");
  assert.equal(e.context.fixtureShape.Locked, false);
  assert.equal(W.layerSetVisible({ shape: e.context.fixtureShape }, true).ok, false, "unbound row must fail closed");
  e.native.ActiveWindow.HWND = 101;
  assert.equal(W.layerSetVisible(row, true).ok, true, "same-window row is still usable after returning");
}

async function testRibbonCompletion() {
  const e = environment();
  let starts = 0, complete;
  e.context.runFixture = () => { starts++; return new Promise(resolve => { complete = resolve; }); };
  engine(e, `
    selectedPicture = function () { return {}; };
    tell = function () {};
    replaceAllFromClipboard = async function () { await global.runFixture(); return {matched:1,success:1,failed:0}; };
    runBatchWithProgress = async function (title, work) { return await work(function(){},function(){return false;}); };
  `);
  e.context.OnAddInLoad = () => {};
  vm.runInContext(bootstrap, e.context);
  const first = e.context.__pptDispatch("ReplaceAllClipboard");
  await flush();
  const blocked = e.context.__pptDispatch("ReplaceAllClipboard");
  await blocked;
  assert.equal(starts, 1, "Ribbon mutex must remain held throughout the generated asynchronous wrapper");
  complete(); await first;
  const next = e.context.__pptDispatch("ReplaceAllClipboard");
  await flush(); assert.equal(starts, 2, "next operation is allowed after real completion");
  complete(); await next;
  e.native.ActiveWindow.HWND = 202;
  await e.context.__pptDispatch("ReplaceAllClipboard", "window:101");
  assert.equal(starts, 2, "a queued command from document A must not run on document B");
}

async function testScratchEvents() {
  const e = environment();
  let zoomEnded = 0, cancelled = 0, switches = 0;
  e.context.WpsPictureReplace = { addinVersion: "3.0.0", smartZoomEnd: () => zoomEnded++, readTaskState: () => ({ running: true }), requestCancelTask: () => cancelled++ };
  e.context.OnAddInLoad = () => {};
  e.context.addEventListener("ppt-document-changed", () => switches++);
  vm.runInContext(bootstrap, e.context);
  e.context.__pptEmit("NewPresentation");
  e.context.__pptEmit("PresentationClose");
  e.context.__pptEmit("WindowActivate");
  assert.equal(zoomEnded, 0, "scratch/nonactive document events must keep the active zoom alive");
  e.native.ActiveWindow.HWND = 202;
  e.context.__pptEmit("WindowActivate");
  e.context.__pptEmit("PresentationOpen");
  assert.equal(zoomEnded, 1); assert.equal(switches, 1); assert.equal(cancelled, 1);
}

async function testDestructiveConfirmation() {
  const e = environment();
  const c = e.context;
  c.pptUiDocumentKey = c.__pptDocumentKey;
  vm.runInContext(extract("      function armButton(", "      var refreshSeq ="), c);
  let click, mutations = 0;
  const button = { classList: classList(), textContent: "删除", addEventListener: (_, fn) => { click = fn; } };
  c.armButton(button, "删除", "再次确认", () => mutations++);
  click(); assert(button.classList.contains("armed"));
  e.native.ActiveWindow.HWND = 202;
  click(); assert.equal(mutations, 0, "a late document-change event must not allow A confirmation to mutate B");
  click(); assert.equal(mutations, 1, "two new B clicks may confirm B");
  click();
  c.document = { querySelectorAll: () => [button] };
  c.refreshSeq = 1; c.layerState = { refreshSeq: 1, result: {}, items: [{}] };
  c.el = { replaceSelected: { classList: classList(), textContent: "old", _armed: true }, groups: { textContent: "old" } };
  c.layerEl = { list: { textContent: "old" } };
  c.state = { groups: [{}], selected: { old: true }, docKey: "old", scanning: true };
  c.zoomTimer = null; c.zoomSessionId = 9;
  c.zoomEndDisarm = c.zoomSetControlsDisabled = c.zoomSetStatus = c.hideProgress = () => {};
  c.location.hash = "#tools/cleanup";
  vm.runInContext(extract('      window.addEventListener("ppt-document-changed"', '      window.addEventListener("hashchange"'), c);
  c.dispatchEvent(new c.CustomEvent("ppt-document-changed", { detail: { previous: "window:202", current: "window:303" } }));
  assert.equal(button.classList.contains("armed"), false);
  assert.equal(c.el.replaceSelected._armed, false);
  assert.equal(c.layerState.items.length, 0); assert.equal(c.state.groups.length, 0); assert.equal(c.zoomSessionId, null);

  let release, wrote = false, status = "";
  c.toolState = { busy: false };
  c.yieldPane = () => new Promise(resolve => { release = resolve; });
  c.toolStatus = (_, message) => { status = message; };
  vm.runInContext(extract("      async function toolRun(", "      function setupTools()"), c);
  const work = c.toolRun({ textContent: "清理", getAttribute: () => "status" }, () => { wrote = true; });
  e.native.ActiveWindow.HWND = 303;
  release(); await work;
  assert.equal(wrote, false); assert.match(status, /文稿已切换/);
  assert.equal(c.toolState.busy, false);
}

async function testProgressLifecycle() {
  const e = environment();
  const panel = e.context.wps.CreateTaskPane("https://ppt-toolbox.local/taskpane.html#tools/style", "tools");
  panel.Visible = true;
  e.context.location.hash = "#tools/color";
  const progress = e.context.wps.CreateTaskPane("https://ppt-toolbox.local/taskpane.html#progress", "progress");
  progress.Visible = true; progress.Visible = false;
  assert.equal(e.context.location.hash, "#tools/color", "closing progress restores the preceding route");
  assert.equal(e.host.hidden, 0, "progress completion must not hide the whole toolbox");

  engine(e);
  const W = e.context.WpsPictureReplace;
  const first = W.openProgressPanel("第一任务");
  W.closeProgressPanel(first);
  const oldClose = [...e.timers.values()].at(-1);
  const second = W.openProgressPanel("第二任务");
  oldClose();
  assert.equal(second.Visible, true, "the preceding task's delayed close cannot hide the next one");
  assert.equal(W.readTaskState().running, true);
  assert.equal(W.readTaskState().title, "第二任务", "task text uses the text file API");

  let state = { runId: 10, running: false, label: "done" }, tick;
  const elements = { taskPane: { classList: classList() }, taskCancel: { disabled: false, addEventListener() {} }, taskTitle: {}, taskText: {}, taskFill: { style: {} } };
  const poll = {
    window: {}, location: { hash: "#progress" }, document: { getElementById: id => elements[id] },
    W: { readTaskState: () => state }, setTimeout: fn => { tick = fn; return 1; }, clearTimeout() {}
  };
  vm.createContext(poll);
  vm.runInContext(extract("      function startTaskPoll()", "      function layerSetStatus("), poll);
  poll.startTaskPoll(); tick(); assert.equal(elements.taskCancel.disabled, true);
  state = { runId: 11, running: true, done: 1, total: 10 };
  poll.startTaskPoll(); tick(); assert.equal(elements.taskCancel.disabled, false, "second task can be cancelled");
  poll.location.hash = "#panel"; tick = null;
  vm.runInContext("startTaskPoll()", poll);
  const lastTick = tick; tick = null; lastTick();
  assert.equal(tick, null, "no polling is rescheduled outside the progress route");
}

async function testScratchOwnershipAndRibbonSchema() {
  const e = environment();
  const tracked = new Set();
  e.host.TrackScratch = raw => tracked.add(raw);
  e.host.ReleaseScratch = raw => tracked.delete(raw);
  let userSavedWrites = 0, closed = 0, deleted = 0;
  const user = { Name: "User", set Saved(_) { userSavedWrites++; } };
  const shapeCollection = { Count: 1, Item: () => ({ Delete() { deleted++; shapeCollection.Count = 0; } }) };
  const scratch = {
    Name: "Owned scratch", Saved: 0,
    Slides: { Count: 0, Add: () => ({ Shapes: shapeCollection }) },
    Close() { closed++; }
  };
  e.native.ActivePresentation = user;
  e.native.Presentations = { Add: () => scratch };
  engine(e, "  global.scratchFixture = createScratchManager;\n");
  const manager = e.context.scratchFixture();
  manager.ensure();
  assert(tracked.has(scratch), "adapter must register the raw owned COM object");
  assert(!tracked.has(user));
  manager.clear();
  assert.equal(deleted, 1);
  assert.equal(scratch.Saved, -1, "pooled scratch clear suppresses save prompts only for the owned document");
  assert.equal(userSavedWrites, 0);
  manager.dispose();
  assert.equal(closed, 1); assert.equal(tracked.size, 0);

  scratch.Close = () => { throw new Error("host still busy"); };
  const retry = e.context.scratchFixture();
  retry.ensure(); retry.dispose();
  assert(tracked.has(scratch), "failed Close must remain in the native shutdown registry");
  e.native.Presentations.Add = () => user;
  assert.throws(() => e.context.scratchFixture().ensure(), /独立/);
  assert(!tracked.has(user)); assert.equal(userSavedWrites, 0);

  const ribbon = fs.readFileSync(path.join(root, "shared/ribbon.xml"), "utf8");
  const sizes = [...ribbon.matchAll(/\bsize="([^"]+)"/g)].map(match => match[1]);
  assert(sizes.length > 0);
  assert(sizes.every(size => size === "normal" || size === "large"), "Office rejects non-schema sizes such as small");
  assert.match(ribbon, /id="AddonVersion"[^>]*size="normal"/);
  assert.match(ribbon, /id="AddonAuthor"[^>]*size="normal"/);
}

async function testPictureInputOwnership() {
  const e = environment();
  const c = e.context;
  c.fixtureMode = "success";
  c.atob = value => Buffer.from(value, "base64").toString("binary");
  c.FileReader = class {
    readAsDataURL() { this.result = "data:image/png;base64,aW1hZ2U="; this.onload(); }
  };
  engine(e, `
    selectedPicture = function () { return {}; };
    replacePictureKeepCrop = function () { if (global.fixtureMode === "failure") throw new Error("insert failed"); return {}; };
    replaceAllMatching = async function () {
      if (global.fixtureMode === "failure") throw new Error("scan failed");
      if (global.fixtureMode === "unmatched") return {matched:0,success:0,failed:0};
      return {matched:2,success:global.fixtureMode === "cancelled" ? 1 : 2,failed:0,cancelled:global.fixtureMode === "cancelled"};
    };
  `);
  const W = c.WpsPictureReplace;
  const original = "G:\\pictures\\user-original.png";
  const userTemp = "C:\\port-tests\\user-original.png";
  e.files.set(original, "original-bytes"); e.files.set(userTemp, "user-temp-bytes");
  for (const mode of ["success", "failure", "unmatched", "cancelled"]) {
    c.fixtureMode = mode;
    try { await W.replaceAllFromFile(original); } catch (_) {}
    assert.equal(e.files.get(original), "original-bytes", "batch must preserve user input after " + mode);
    try { await W.replaceSelectedFromFile(original); } catch (_) {}
    assert.equal(e.files.get(original), "original-bytes", "single must preserve user input after " + mode);
  }
  assert.equal(W.releaseBrowserFile(original), false);
  assert.equal(W.releaseBrowserFile(userTemp), false, "TEMP is not sufficient proof of ownership");
  const allocated = await W.writeBrowserFile({ size: 5 });
  assert(e.files.has(allocated));
  assert.equal(W.releaseBrowserFile(allocated), true);
  assert(!e.files.has(allocated));
  assert.equal(W.releaseBrowserFile(allocated), false, "release is idempotent");

  c.W = W;
  vm.runInContext(extract("      function setupLegacy(", "      Object.keys(LEGACY)"), c);
  for (const mode of ["success", "failure", "cancelled"]) {
    c.fixtureMode = mode;
    let click;
    const cfg = {
      batch: true, file: { files: [{ size: 5 }], addEventListener() {} },
      run: { disabled: false, addEventListener: (_, fn) => { click = fn; } }, status: {}
    };
    c.setupLegacy(cfg); await click();
    assert.equal(e.files.size, 2, "UI must release browser-owned temporary data after " + mode);
    assert.equal(cfg.run.disabled, false);
  }
}

async function testClipboardFailureSafety() {
  const preview = environment();
  engine(preview, "  global.previewFixture = exportUncroppedPreview;\n");
  let pasteCalls = 0, cleared = 0;
  const output = "C:\\port-tests\\stale-preview.png";
  preview.files.set(output, "stale preview");
  const scratch = {
    ensure: () => ({ Shapes: { Count: 0, Paste() { pasteCalls++; } } }),
    clear() { cleared++; }
  };
  const result = preview.context.previewFixture({ Copy() { throw new Error("Copy failed"); } }, scratch, output, false, false);
  assert.equal(result.ok, false);
  assert.equal(pasteCalls, 0, "failed Copy must never fingerprint old clipboard contents");
  assert.equal(preview.files.has(output), false, "failed preview must discard a stale output");
  assert.equal(cleared, 2, "scratch is cleaned before use and after failure");

  for (const mode of ["capture", "apply", "delete", "success"]) {
    const e = environment(), c = e.context;
    let oldDeletes = 0, newDeletes = 0, disposed = 0;
    const inserted = { PictureFormat: {}, Width: 100, Height: 80, Left: 0, Top: 0, HorizontalFlip: 0, VerticalFlip: 0, ZOrderPosition: 2,
      Delete() { newDeletes++; }, Select() {} };
    const shapes = { Count: 1, PasteSpecial() { this.Count++; }, Item: () => inserted };
    const original = { Parent: { Shapes: shapes }, Delete() { oldDeletes++; if (mode === "delete") throw new Error("delete ambiguous"); } };
    c.fixtureMode = mode;
    c.fixtureDispose = () => disposed++;
    engine(e, `
      createScratchManager = function () { return {dispose:global.fixtureDispose}; };
      readClipboardPixelSize = async function () { return null; };
      captureState = function () { return {captureFailed:global.fixtureMode === "capture",flipH:false,flipV:false,left:10,top:20,rotation:0,zOrder:1,name:"source",alternativeText:""}; };
      applyPreservedCrop = function () { if (global.fixtureMode === "apply") throw new Error("apply failed"); };
      copyCosmetics = function () {};
      invalidatePanelInventoryCache = function () {};
    `);
    const operation = c.WpsPictureReplace.pasteReplacePictureKeepCrop(original);
    if (mode === "success") assert.equal(await operation, inserted);
    else await assert.rejects(operation, mode === "capture" ? /裁剪信息/ : mode === "apply" ? /apply failed/ : /delete ambiguous/);
    assert.equal(disposed, 1);
    assert.equal(oldDeletes, mode === "capture" || mode === "apply" ? 0 : 1);
    assert.equal(newDeletes, mode === "capture" || mode === "apply" ? 1 : 0,
      "rollback only before deleting the original: " + mode);
  }
}

async function testSlideExtractionSnapshot() {
  for (const mode of ["unsaved", "macros", "unknown-macros", "save-failure", "insert-zero", "insert-throws", "wrong-count", "same-target", "close-failure", "collision"]) {
    const e = environment(), c = e.context;
    let sourceSavedWrites = 0, sourceClosed = 0, targetClosed = 0, targetSavedWrites = 0, adds = 0;
    const saves = [], imports = [], tracked = new Set();
    const preexisting = "C:\\port-tests\\preexisting.pptx";
    e.files.set(preexisting, "user data");
    e.host.TrackScratch = value => tracked.add(value);
    e.host.ReleaseScratch = value => tracked.delete(value);
    const target = {
      PageSetup: { SlideWidth: 720, SlideHeight: 540 },
      set Saved(_) { targetSavedWrites++; },
      Close() { targetClosed++; if (mode === "close-failure") throw new Error("close failed"); },
      Slides: {
        Count: 0,
        InsertFromFile(snapshot, after, start, end) {
          assert.equal(e.files.get(snapshot), "unsaved source snapshot");
          assert.equal(after, this.Count); assert.equal(start, end);
          imports.push(start);
          if ((mode === "insert-throws" || mode === "close-failure") && imports.length === 2) throw new Error("insert failed");
          if (mode === "insert-zero" && imports.length === 2) return 0;
          this.Count += mode === "wrong-count" ? 2 : 1;
          return 1;
        }
      }
    };
    const source = {
      Slides: { Count: 4, Range() { throw new Error("must not use clipboard"); } },
      PageSetup: { SlideWidth: 960, SlideHeight: 540 },
      HasVBProject: mode === "macros" ? true : mode === "unknown-macros" ? undefined : false,
      get Path() { throw new Error("unsaved path must not be read"); },
      get FullName() { throw new Error("unsaved filename must not be read"); },
      get Saved() { return 0; },
      set Saved(_) { sourceSavedWrites++; },
      Close() { sourceClosed++; },
      SaveCopyAs(snapshot, format, embed) {
        assert.equal(adds, 0, "snapshot must be captured before Add activates target");
        saves.push({ snapshot, format, embed });
        e.files.set(snapshot, "unsaved source snapshot");
        if (mode === "save-failure") throw new Error("partial save failed");
      }
    };
    c.fixtureCurrent = source;
    c.fixtureApp = { FileSystem: c.wps.FileSystem, Presentations: { Add() { adds++; c.fixtureCurrent = target; return mode === "same-target" ? source : target; } } };
    c.fixtureMode = mode;
    engine(e, `
      activePresentation = function () { return global.fixtureCurrent; };
      application = function () { return global.fixtureApp; };
      designSlides = function (scope) { if (scope !== "selected") throw new Error("wrong scope"); return [{SlideIndex:3},{SlideIndex:1},{SlideIndex:3}]; };
      if (global.fixtureMode === "collision") tempPath = function () { return "C:\\\\port-tests\\\\preexisting.png"; };
    `);
    const operation = c.WpsPictureReplace.designExtractSlides("selected");
    const success = ["unsaved", "macros", "unknown-macros"].includes(mode);
    if (success) {
      const result = await operation;
      assert.equal(result.count, 2);
      assert.deepEqual(imports, [3, 1], "selected order is preserved and duplicate indexes are omitted");
      assert.equal(target.PageSetup.SlideWidth, 960); assert.equal(target.PageSetup.SlideHeight, 540);
      assert.equal(targetClosed, 0); assert.equal(targetSavedWrites, 0, "successful output stays unsaved");
      const macro = mode !== "unsaved";
      assert.equal(saves[0].format, macro ? 25 : 24);
      assert.equal(path.extname(saves[0].snapshot), macro ? ".pptm" : ".pptx");
    } else {
      await assert.rejects(operation);
      const ownTarget = !["save-failure", "same-target", "collision"].includes(mode);
      assert.equal(targetClosed, ownTarget ? 1 : 0);
      assert.equal(targetSavedWrites, ownTarget ? 1 : 0);
    }
    assert.equal(sourceSavedWrites, 0); assert.equal(sourceClosed, 0);
    assert.equal(source.Saved, 0, "an unsaved source remains unsaved");
    assert.equal(e.files.get(preexisting), "user data");
    assert.equal(e.files.size, 1, "only the owned snapshot may be removed, including a partial failed save");
    assert.equal(tracked.has(target), mode === "close-failure", "failed target Close remains tracked for shutdown cleanup");
    assert.equal(tracked.has(source), false);
  }
}

(async () => {
  for (const test of [testObjectOwnership, testRibbonCompletion, testScratchEvents, testDestructiveConfirmation, testProgressLifecycle, testScratchOwnershipAndRibbonSchema, testPictureInputOwnership, testClipboardFailureSafety, testSlideExtractionSnapshot]) {
    await test(); console.log("PASS " + test.name);
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
