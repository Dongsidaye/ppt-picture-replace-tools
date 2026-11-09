"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const root = path.resolve(__dirname, "..");
let shown = 0, hidden = 0, eventCount = 0;
const files = new Map();
const host = {
  TempFolder: "C:\\temp\\PowerPointTools\\test", RootFolder: "C:\\toolbox",
  FileExists: p => files.has(p), ReadBinary: p => files.get(p), WriteBinary: (p, value) => files.set(p, value),
  ReadText: p => files.get(p), WriteText: (p, value) => files.set(p, value), RemoveFile: p => files.delete(p),
  ReadBinaryBase64: p => Buffer.from(files.get(p), "latin1").toString("base64"), WriteBinaryBase64: (p, value) => files.set(p, Buffer.from(value, "base64").toString("latin1")),
  StatJson: () => '{"size":3,"mtimeMs":1000}', ShowPane: () => shown++, HidePane: () => hidden++, Log: () => {}
};
host.ReadMember = (target, name) => { if (typeof target[name] === "function") throw new Error("method"); return target[name]; };
host.WriteMember = (target, name, value) => { target[name] = value; };
const native = { Version: "16.0", Presentations: { Count: 1 } };
Object.defineProperty(native, "MissingOptionalMember", { get() { throw new Error("未知名称。 (0x80020006)"); } });
const window = {
  chrome: { webview: { hostObjects: { sync: { pptHost: host, pptApplication: native } } } },
  addEventListener() {}, dispatchEvent() { eventCount++; }, alert() {}, atob, btoa
};
const location = { href: "https://ppt-toolbox.local/taskpane.html#idle", origin: "https://ppt-toolbox.local", hash: "#idle" };
const context = vm.createContext({ window, location, URL, Event, CustomEvent: class { constructor(type, options) { this.type = type; this.detail = options.detail; } }, Date, Map, Set, Proxy });
vm.runInContext(fs.readFileSync(path.join(root, "web/host-adapter.js"), "utf8"), context);
assert.equal(window.wps.Version, "16.0");
assert.equal(window.wps.Presentations.Count, 1);
assert.equal(window.wps.MissingOptionalMember, undefined, "unsupported COM properties must support capability probes");
window.wps.FileSystem.writeAsBinaryString("binary", "\0\x80\xff");
assert.equal(window.wps.FileSystem.readAsBinaryString("binary"), "\0\x80\xff");
assert.throws(() => window.wps.CreateTaskPane("https://evil.example/taskpane.html", "x"), /不受信任/);
assert.throws(() => window.wps.CreateTaskPane("https://ppt-toolbox.local/other.html", "x"), /不受信任/);
const a = window.wps.CreateTaskPane(location.href.replace("#idle", "#layers"), "objects");
a.Visible = true;
const b = window.wps.CreateTaskPane(location.href.replace("#idle", "#zoom"), "zoom");
b.Visible = true; a.Visible = false;
assert.equal(hidden, 0, "old pane reference must not hide current pane");
b.Visible = false;
assert.equal(shown, 2); assert.equal(hidden, 1);
let events = 0;
const handler = () => events++;
window.wps.ApiEvent.AddApiEventListener("WindowActivate", handler);
window.wps.ApiEvent.AddApiEventListener("WindowActivate", handler);
window.__pptEmit("WindowActivate"); assert.equal(events, 1);
window.wps.ApiEvent.RemoveApiEventListener("WindowActivate", handler);
window.__pptEmit("WindowActivate"); assert.equal(events, 1);
assert.throws(() => { window.wps.FileSystem = {}; }, /不可覆盖/);
const actions = JSON.parse(fs.readFileSync(path.join(root, "shared/action-map.json")));
const main = fs.readFileSync(path.join(root, "shared/main.js"), "utf8");
new vm.Script(main);
Object.entries(actions).forEach(([id, callback]) => assert(main.includes("global." + callback + " ="), "Missing command " + id));
const icons = JSON.parse(fs.readFileSync(path.join(root, "shared/icon-map.json")));
Object.entries(icons).forEach(([id, filename]) => {
  const bytes = fs.readFileSync(path.join(root, "shared", filename));
  assert.equal(bytes.subarray(1, 4).toString(), "PNG", "Missing PNG " + id);
});
assert(!main.includes('const ADDIN_VERSION = "2.'), "wrong product version");
assert(!main.includes('"/PictureReplaceTools-WPS-'), "wrong installer platform");
assert(/return (?:await )?global\.PptHostUpdate\(\)/.test(main), "native updater missing");
console.log("PASS adapter: COM forwarding, binary files, trusted pane routing, stale pane handle, event de-duplication, platform update isolation");
console.log("PASS " + Object.keys(actions).length + " Ribbon actions and " + Object.keys(icons).length + " icon mappings");

const dialogTimers = [];
window.setTimeout = callback => { dialogTimers.push(callback); return dialogTimers.length; };
host.DocumentKey = "document:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa:window:101";
assert.equal(window.__pptDocumentKey(),host.DocumentKey,"native session token must survive the bridge unchanged");
host.DocumentKey = "window:101";
assert.equal(window.__pptDocumentKey(),"","reusable HWND alone cannot authorize an operation");
host.DocumentKey = "document:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb:window:101";
assert.equal(window.__pptDocumentKey(),host.DocumentKey,"a reused HWND must retain the new document session identity");
host.DocumentKey = "document:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa:window:101";
let nextDialogResult = { state: "done", ok: true, cancelled: false, path: "C:\\exports", error: "" };
let dialogBegins = 0, pendingPolls = 0, writes = 0;
host.BeginFileDialog = requestJson => {
  const request = JSON.parse(requestJson);
  assert.equal(request.documentKey, "document:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa:window:101");
  assert(["file", "folder"].includes(request.kind));
  return "dialog-" + (++dialogBegins);
};
host.PollFileDialog = () => JSON.stringify(pendingPolls-- > 0 ? { state: "pending" } : nextDialogResult);
native.Mutate = () => { writes++; };
async function tickDialog() {
  const tick = dialogTimers.shift();
  assert(tick, "file selection must yield rather than synchronously show a modal dialog");
  tick();
  for (let i = 0; i < 8; i++) await Promise.resolve();
}
async function testAsyncFileDialogs() {
  assert.throws(() => window.wps.FileDialog(4), /异步文件选择/, "old synchronous picker paths must fail without calling COM Show");
  pendingPolls = 1;
  const choose = window.__pptChooseFile({ kind: "folder", title: "导出页面" });
  assert.equal(window.__pptFileDialogPending(), true);
  assert.throws(() => window.wps.Mutate(), /文件选择/);
  assert.throws(() => { window.wps.Version = "changed"; }, /文件选择/);
  assert.equal(window.wps.Version, "16.0", "property reads remain available");
  await assert.rejects(window.__pptChooseFile({ kind: "file" }), /文件选择/);
  const imagePane = window.wps.CreateTaskPane("https://ppt-toolbox.local/taskpane.html#panel", "图片清单");
  imagePane.Visible = true;
  assert.equal(location.hash, "#panel", "navigation remains available during async file selection");
  await tickDialog(); assert.equal(window.__pptFileDialogPending(), true);
  await tickDialog();
  const result = await choose;
  assert.equal(result.path, "C:\\exports"); assert.equal(result.ok, true);
  assert.equal(window.__pptFileDialogPending(), false);
  window.wps.Mutate(); assert.equal(writes, 1);

  for (const mode of ["cancel", "error", "switch", "bad-json"]) {
    nextDialogResult = mode === "cancel" ? { state: "done", ok: false, cancelled: true, path: "" }
      : { state: "done", ok: true, path: "C:\\a.png", error: mode === "error" ? "原生选择失败" : "" };
    const beforePoll = host.PollFileDialog;
    if (mode === "bad-json") host.PollFileDialog = () => "bad json";
    const operation = window.__pptChooseFile({ kind: "file", filter: "*.png" });
    const observed = operation.then(value => ({ value }), error => ({ error }));
    if (mode === "switch") host.DocumentKey = "document:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb:window:202";
    await tickDialog();
    const outcome = await observed;
    if (mode === "cancel") assert.equal(outcome.value.cancelled, true);
    else assert(outcome.error, mode + " must fail closed");
    assert.equal(window.__pptFileDialogPending(), false, mode + " must release the lock");
    host.DocumentKey = "document:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa:window:101"; host.PollFileDialog = beforePoll;
    imagePane.Visible = true; assert.equal(location.hash, "#panel");
  }
  const begin = host.BeginFileDialog;
  host.BeginFileDialog = () => { throw new Error("native unavailable"); };
  await assert.rejects(window.__pptChooseFile({ kind: "file" }), /unavailable/);
  assert.equal(window.__pptFileDialogPending(), false, "Begin failure must release the lock");
  host.BeginFileDialog = begin;
  console.log("PASS asynchronous picker: yield/poll, cancellation, COM mutation gate, document switch, exceptions, re-open image panel");
}
async function testNavigationMutex() {
  context.fetch = async () => ({ ok: true, json: async () => actions });
  window.WpsPictureReplace = { addinVersion: "3.0.0" };
  window.OnAddInLoad = () => {};
  let mutationStarts = 0, release, navigations = 0;
  window.ReplaceAllFromClipboard = () => { mutationStarts++; return new Promise(resolve => { release = resolve; }); };
  window.OpenPicturePanel = () => { navigations++; };
  vm.runInContext(fs.readFileSync(path.join(root, "web/host-bootstrap.js"), "utf8"), context);
  const job = window.__pptDispatch("ReplaceAllClipboard", "document:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa:window:101");
  for (let i = 0; i < 8; i++) await Promise.resolve();
  await window.__pptDispatch("OpenPicturePanelButton", "document:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa:window:101");
  await window.__pptDispatch("ReplaceAllClipboard", "document:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa:window:101");
  assert.equal(navigations, 1, "navigation is not blocked by a long operation");
  assert.equal(mutationStarts, 1, "navigation cannot release the mutation lock");
  release(); await job;
  nextDialogResult = { state: "done", ok: false, cancelled: true, path: "" };
  const picker = window.__pptChooseFile({ kind: "folder" });
  await window.__pptDispatch("ReplaceAllClipboard", "document:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa:window:101");
  assert.equal(mutationStarts, 1, "a pending native picker blocks a second Ribbon mutation");
  await window.__pptDispatch("OpenPicturePanelButton", "document:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa:window:101");
  assert.equal(navigations, 2);
  await tickDialog(); await picker;
  await window.__pptDispatch("OpenPicturePanelButton", "document:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa:window:101");
  assert.equal(navigations, 3, "cancelled export must not prevent reopening the image list");
  console.log("PASS Ribbon navigation remains available without unlocking concurrent mutations");
}
testAsyncFileDialogs().then(testNavigationMutex).catch(error => { console.error(error); process.exitCode = 1; });
