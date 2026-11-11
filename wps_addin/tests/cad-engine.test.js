"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");
const script = fs.readFileSync(path.join(__dirname, "..", "cad-engine.js"), "utf8");
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));
const key = value => String(value).replace(/\\/g, "/").toLowerCase();

function fixture(settings = {}) {
  const files = new Map(), listeners = {}, events = {}, requests = [];
  let shapes = [], serial = 0, serverRunning = false;
  const calls = { text: 0, picture: 0 };
  function shape(type) {
    const value = { Id: ++serial, Type: type, Name: "shape-" + serial, Delete() {
      const index = shapes.indexOf(value);
      if (index >= 0) shapes.splice(index, 1);
    }, ZOrder() {}, Select() {} };
    shapes.push(value);
    return value;
  }
  const original = shape(1);
  const slide = { SlideID: 100, Shapes: {
    AddPicture() { calls.picture++; return shape(28); },
    Range(names) { return { Group() {
      const members = shapes.filter(item => names.includes(item.Name));
      members.forEach(item => item.Delete());
      const group = shape(6); group.members = members;
      return group;
    } }; }
  } };
  const windowObject = { Active: -1, HWND: 101, ViewType: 9, View: { Slide: slide } };
  const presentation = { Name: "test.pptx", FullName: "C:\\test.pptx", Windows: { Count: 1 }, PageSetup: { SlideWidth: 960, SlideHeight: 540 } };
  const app = { ActivePresentation: presentation, ActiveWindow: windowObject,
    CurrentWPSAddIn: { Path: "C:\\addon", Name: "test" }, Env: { GetTempPath: () => "C:\\Temp\\" },
    ApiEvent: { AddApiEventListener(name, callback) { events[name] = callback; } },
    FileSystem: { tmpdir: () => "C:\\Temp\\", Exists: file => files.has(key(file)),
      WriteFile(file, content) { files.set(key(file), content); if (serverRunning && file.endsWith(".ready")) processReady(file); }, ReadFile: file => files.get(key(file)),
      unlinkSync(file) { files.delete(key(file)); } },
    OAAssist: { ShellExecute(exe, parameters) {
      assert.equal(exe, "C:\\addon\\cad\\WpsCadBridge.exe");
      assert.equal(parameters, "--serve");
      if (settings.launchError) throw new Error("launch rejected");
      serverRunning = true;
      files.set(key("C:\\Temp\\DongsidayeWpsCad.aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa.server.json"), JSON.stringify({protocol:1,exe,workerHash:"a".repeat(64),heartbeatAt:Date.now()}));
      for (const file of files.keys()) if (file.endsWith(".ready")) processReady(file);
    } }
  };
  function processReady(file) {
      const marker = JSON.parse(files.get(key(file)));
      const requestPath = marker.request;
      files.delete(key(file));
      const request = JSON.parse(files.get(key(requestPath)));
      requests.push(request);
      const prefix = requestPath.replace(/\.[a-z]+\.json$/, "");
      const statusPath = prefix + "." + request.phase + ".status.json";
      let result;
      if (request.phase === "convert") result = { textCount: 2, entityCount: 3, previewBase64: "PHN2Zy8+" };
      if (request.phase === "prepare") result = { left: 10, top: 20, width: 700, height: 400,
        texts: Array.from({ length: settings.rows || 2 }, (_, n) => ({ item: { SourceId: settings.groupFailure && n < 2 ? "paragraph" : "line-" + n, Text: "text-" + n }, ocr: false })) };
      if (request.phase === "finalize") {
        if (settings.changePageSize) presentation.PageSetup.SlideWidth = 800;
        result = { path: prefix + "\\insertion.svg", hasGeometry: true };
      }
      const name = { convert: "converted.json", prepare: "plan.json", finalize: "finalized.json" }[request.phase];
      const output = prefix + "\\" + name;
      if (request.phase !== "cleanup") files.set(key(output), JSON.stringify(result));
      files.set(key(statusPath), JSON.stringify({ state: "done", ok: true, output: settings.badOutput && request.phase === "prepare" ? "C:\\outside.json" : output }));
  }
  files.set(key("C:\\addon\\cad\\WpsCadBridge.exe"), "binary");
  files.set(key("C:\\addon\\checksums.json"), JSON.stringify({"cad/WpsCadBridge.exe":"a".repeat(64)}));
  const env = { wps: app, crypto: crypto.webcrypto, console, Event: class { constructor(type) { this.type = type; } },
    setTimeout: (callback, ms) => setTimeout(callback, Math.min(ms, 3)), setInterval() {},
    addEventListener(name, callback) { listeners[name] = callback; }, dispatchEvent(event) { if (listeners[event.type]) listeners[event.type](event); },
    WpsPictureReplace: { clearDeckImageCache() {} }, WpsCadText: { add() {
      calls.text++;
      if (settings.groupFailure && calls.text === 2) throw new Error("font fit rejected");
      return shape(17);
    } }
  };
  env.window = env;
  vm.runInNewContext(script, env, { filename: "cad-engine.js" });
  const api = env.WpsCadBridge;
  function begin(ocr) { return api.BeginDwgImport(JSON.stringify({ path: "C:\\图纸 & 方案.dwg", documentKey: api.documentKey(), ocr: !!ocr })); }
  return { api, env, app, windowObject, presentation, events, listeners, requests, calls, original, shapes: () => shapes, begin };
}

async function waitFor(read, condition) {
  for (let n = 0; n < 1500; n++) {
    const result = read();
    if (condition(result)) return result;
    await delay(2);
  }
  throw new Error("Test timed out");
}
async function converted(f, id) {
  return waitFor(() => JSON.parse(f.api.PollDwgImport(id)), value => value.state === "done");
}
async function inserted(f, id) {
  return waitFor(() => JSON.parse(f.api.PollDwgInsertion(id)), value => value.state === "done");
}

(async function () {
  {
    const f = fixture(), id = f.begin(true);
    assert((await converted(f, id)).ok);
    f.api.BeginDwgInsertion(id, JSON.stringify({ font: "Microsoft YaHei", ocrSelections: [{ index: 0, text: "校对文字" }] }));
    assert((await inserted(f, id)).ok);
    assert.equal(f.shapes().length, 2);
    assert.equal(f.shapes()[0], f.original);
    assert.equal(f.requests[0].ocr, true);
    assert.equal(f.requests.find(value => value.phase === "prepare").ocrSelections[0].text, "校对文字");
    console.log("PASS normal entry reuses one worker and preserves existing objects");
  }
  {
    const f = fixture({ groupFailure: true, rows: 3 }), id = f.begin();
    await converted(f, id); f.api.BeginDwgInsertion(id, "{}");
    const result = await inserted(f, id);
    assert(result.ok); assert.equal(result.textCount, 1); assert.equal(result.fallbackCount, 1);
    assert.deepEqual(f.requests.find(value => value.phase === "finalize").fallbackIds, ["paragraph"]);
    console.log("PASS a failed line restores its complete source paragraph");
  }
  for (const mode of ["cancel", "switch", "unload"]) {
    const f = fixture({ rows: 80 }), id = f.begin();
    await converted(f, id); f.api.BeginDwgInsertion(id, "{}");
    await waitFor(() => f.calls.text, value => value >= 8);
    if (mode === "cancel") f.api.CancelDwgImport(id);
    if (mode === "switch") { f.windowObject.Active = 0; f.app.ActivePresentation = { FullName: "C:\\other.pptx" }; }
    if (mode === "unload") {
      f.listeners.beforeunload();
      assert.deepEqual(f.shapes(), [f.original], "Unload must synchronously remove partial objects");
    }
    const result = await inserted(f, id);
    assert.equal(result.ok, false); assert.deepEqual(f.shapes(), [f.original]);
    console.log("PASS " + mode + " rolls back all partial objects");
  }
  {
    const f = fixture(), id = f.begin(); await converted(f, id);
    f.events.PresentationClose(f.presentation);
    assert.throws(() => f.api.BeginDwgInsertion(id, "{}"), /切换|关闭|取消/);
    assert.equal(f.calls.text, 0);
    console.log("PASS closing and reopening the same filename invalidates the old request");
  }
  for (const mode of ["changePageSize", "badOutput"]) {
    const f = fixture({ [mode]: true }), id = f.begin(); await converted(f, id);
    f.api.BeginDwgInsertion(id, "{}");
    const result = await inserted(f, id);
    assert.equal(result.ok, false); assert.deepEqual(f.shapes(), [f.original]); assert.equal(f.calls.picture, 0);
    console.log("PASS " + mode + " rejects before embedding geometry");
  }
  {
    const f = fixture(), id = f.begin(); await converted(f, id);
    assert.throws(() => f.api.BeginDwgInsertion(id, "{"));
    assert.throws(() => f.api.BeginDwgInsertion(id, "null"));
    f.api.BeginDwgInsertion(id, "{}"); assert((await inserted(f, id)).ok);
    console.log("PASS malformed insertion options do not leave the busy flag set");
  }
  {
    const f = fixture({ launchError: true }), id = f.begin();
    assert.equal((await converted(f, id)).ok, false);
    assert.equal(f.requests.length, 0); assert.deepEqual(f.shapes(), [f.original]);
    console.log("PASS a rejected server launch never retries a different command");
  }
}()).catch(error => { console.error(error); process.exitCode = 1; });
