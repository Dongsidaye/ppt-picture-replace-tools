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
