
// ============================================================
// WPS JSAPI mock harness for Picture Replace Tools v1.2.0
// ============================================================
// Run: node mock_e2e.js   (loads main.js with a mocked wps JSAPI)
// Verifies: deck image collection/grouping, thumbnails, link lifecycle,
// replace/update/unlink/rename/goto, docKey guard, batch regression,
// and same-slide index-shift safety.
// Expected exit code 0: ALL TESTS PASSED.
"use strict";
const fs = require("fs");
const path = require("path");

// ---------- deterministic image model ----------
const IMAGES = {
  "A": { w: 400, h: 300, seed: "IMG_A_400x300" },
  "B": { w: 500, h: 350, seed: "IMG_B_500x350" },
  "C": { w: 640, h: 360, seed: "IMG_C_640x360" }
};

function pngBytes(imgId, w, h) {
  // deterministic pseudo-PNG bytes per image id + size
  const b = Buffer.from("PNG" + IMAGES[imgId].seed + ":" + w + "x" + h + ":" + (w*h));
  return b;
}

// ---------- mock FileSystem ----------
class MockFS {
  constructor() { this.files = new Map(); this.tmp = "C:/mock/tmp/"; }
  tmpdir() { return this.tmp; }
  writeAsBinaryString(p, data) { this.files.set(this._norm(p), String(data)); }
  WriteFile(p, data) { this.files.set(this._norm(p), String(data)); }
  readAsBinaryString(p) {
    const k = this._norm(p);
    if (!this.files.has(k)) throw new Error("mock fs: no such file " + p);
    return this.files.get(k);
  }
  Exists(p) { return this.files.has(this._norm(p)); }
  existsSync(p) { return this.files.has(this._norm(p)); }
  unlinkSync(p) { this.files.delete(this._norm(p)); }
  Remove(p) { this.files.delete(this._norm(p)); }
  _norm(p) { return String(p).replace(/\\/g, "/"); }
}

// ---------- mock Shape ----------
let SHAPE_SEQ = 100;
class MockShape {
  constructor(deck, slide, imageId) {
    this.deck = deck; this.slide = slide;
    this.imageId = imageId;
    this.id = SHAPE_SEQ++;
    this.name = "图片 " + this.id;
    this.alternativeText = "";
    const img = IMAGES[imageId];
    this.width = img.w; this.height = img.h;      // natural size at insert
    this.left = 0; this.top = 0; this.rotation = 0;
    this.hFlip = false; this.vFlip = false;
    this.lockAspectRatio = -1; // msoTrue
    this._cropLeft = 0; this._cropRight = 0; this._cropTop = 0; this._cropBottom = 0;
    this.deleted = false;
    this.z = 0; // assigned by slide
  }
  get Parent() { return this.slide; }
  get PictureFormat() {
    if (this.deleted) throw new Error("shape deleted");
    return this;
  }
  get Width() { return this.width; } set Width(v) { this.width = Number(v); }
  get Height() { return this.height; } set Height(v) { this.height = Number(v); }
  get Left() { return this.left; } set Left(v) { this.left = Number(v); }
  get Top() { return this.top; } set Top(v) { this.top = Number(v); }
  get Rotation() { return this.rotation; } set Rotation(v) { this.rotation = Number(v); }
  get HorizontalFlip() { return this.hFlip; }
  get VerticalFlip() { return this.vFlip; }
  get LockAspectRatio() { return this.lockAspectRatio; } set LockAspectRatio(v) { this.lockAspectRatio = v; }
  get CropLeft() { return this._cropLeft; } set CropLeft(v) { this._cropLeft = Number(v); }
  get CropRight() { return this._cropRight; } set CropRight(v) { this._cropRight = Number(v); }
  get CropTop() { return this._cropTop; } set CropTop(v) { this._cropTop = Number(v); }
  get CropBottom() { return this._cropBottom; } set CropBottom(v) { this._cropBottom = Number(v); }
  get ZOrderPosition() { return this.slide.shapes.indexOf(this) + 1; }
  get Name() { return this.name; } set Name(v) { this.name = String(v); }
  get AlternativeText() { return this.alternativeText; } set AlternativeText(v) { this.alternativeText = String(v); }
  Flip(dir) { if (dir === 0) this.hFlip = !this.hFlip; else this.vFlip = !this.vFlip; }
  ZOrder(dir) {
    const arr = this.slide.shapes;
    const i = arr.indexOf(this);
    if (i < 0) return;
    if (dir === 3) { if (i > 0) { arr[i-1]._zi = arr[i-1]._zi === undefined ? i-1 : arr[i-1]._zi; [arr[i], arr[i-1]] = [arr[i-1], arr[i]]; } }
  }
  Copy() { this.deck.clipboard = this; }
  Duplicate() {
    const d = new MockShape(this.deck, this.slide, this.imageId);
    d.name = this.name + " 副本";
    d.alternativeText = this.alternativeText;
    d.width = this.width; d.height = this.height; d.left = this.left; d.top = this.top;
    d._cropLeft = this._cropLeft; d._cropRight = this._cropRight; d._cropTop = this._cropTop; d._cropBottom = this._cropBottom;
    d.rotation = this.rotation;
    this.slide.shapes.push(d);
    return d;
  }
  Delete() {
    const arr = this.slide.shapes;
    const i = arr.indexOf(this);
    if (i >= 0) arr.splice(i, 1);
    this.deleted = true;
  }
  Select() { this.deck.selectedShape = this; }
  ScaleWidth(w, ignoreAspect) { this.width = w; }
  get Line() { return { Visible: true, Weight: 0, ForeColor: { RGB: 0 } }; }
  get Shadow() { return { Visible: false }; }
  get SoftEdge() { return { Radius: 0 }; }
}

// ---------- mock Slide ----------
class MockSlide {
  constructor(deck, index) {
    this.deck = deck; this.index = index;
    this.shapes = [];
    this.presentation = null;
    this.pageSetup = { slideWidth: 960, slideHeight: 540 };
  }
  get Parent() { return this.presentation; }
  get Shapes() { return this; }
  get Count() { return this.shapes.length; }
  Item(i) {
    const idx = Number(i) - 1;
    if (idx < 0 || idx >= this.shapes.length) throw new Error("shape index out of range " + i);
    return this.shapes[idx];
  }
  AddPicture(file, link, save, left, top, w, h) {
    const id = path.basename(String(file)).replace(/\.[^.]+$/, "").split("_")[0].toUpperCase();
    const sh = new MockShape(this.deck, this, id);
    sh.left = Number(left) || 0; sh.top = Number(top) || 0;
    if (w !== undefined) sh.width = Number(w);
    if (h !== undefined) sh.height = Number(h);
    this.shapes.push(sh);
    return sh;
  }
  Paste() {
    const c = this.deck.clipboard;
    if (!c) throw new Error("mock: clipboard empty");
    const sh = new MockShape(this.deck, this, c.imageId);
    sh._cropLeft = c._cropLeft; sh._cropRight = c._cropRight; sh._cropTop = c._cropTop; sh._cropBottom = c._cropBottom;
    sh.width = IMAGES[c.imageId].w; sh.height = IMAGES[c.imageId].h; sh.left = c.left; sh.top = c.top; sh.rotation = c.rotation;
    sh.hFlip = c.hFlip; sh.vFlip = c.vFlip; sh.lockAspectRatio = c.lockAspectRatio;
    sh.name = c.name; sh.alternativeText = c.alternativeText;
    this.shapes.push(sh);
    return sh;
  }
  PasteSpecial(fmt) { return this.Paste(); }
  Export(p, fmt, w, h) {
    // export the FIRST picture shape (or last) as deterministic bytes
    const pic = this.shapes.find(s => !s.deleted && s.imageId);
    if (!pic) throw new Error("mock export: no picture");
    const ew = Number(w) || 160;
    const eh = Number(h) || 160;
    const bytes = pngBytes(pic.imageId, Math.round(ew), Math.round(eh));
    this.deck.fs.writeAsBinaryString(p, bytes.toString("binary"));
  }
}

// ---------- mock Presentation ----------
class MockPresentation {
  constructor(deck, slidesArray) {
    this.deck = deck;
    this.slidesArray = slidesArray || [];
    this.saved = false;
    this.fullName = deck.path || "";
  }
  get FullName() { return this.fullName; }
  get Slides() {
    const arr = this.slidesArray;
    return {
      get Count() { return arr.length; },
      Item: (i) => { const idx = Number(i) - 1; if (idx < 0 || idx >= arr.length) throw new Error("slide index " + i); return arr[idx]; },
      Add: (i, layout) => { const s = new MockSlide(this.deck, arr.length + 1); s.presentation = this; arr.push(s); return s; }
    };
  }
  get Saved() { return this.saved; } set Saved(v) { this.saved = v; }
  SaveAs(p) { this.fullName = p; }
  Close() { this.closed = true; }
}
let deckRef = null;
const presentationSlides = {};

// ---------- mock Application ----------
function buildApp(deck) {
  deckRef = deck;
  const app = {
    Version: "12.1.0.28043",
    FileSystem: deck.fs,
    CurrentWPSAddIn: { Path: "C:/mock/addin/", Name: "picture-replace-tools-wps" },
    ActivePresentation: null,
    ActiveWindow: { Selection: { ShapeRange: null }, View: { current: 0, GotoSlide(n) { this.current = Number(n); }, get Slide() { return { SlideIndex: this.current }; } } },
    alert(msg) { app._alerts.push(String(msg)); },
    _alerts: [],
    Presentations: {
      Add(win) { return new MockPresentation(deck); },
      Open(p) { const d2 = deck; return new MockPresentation(d2); }
    },
    CreateTaskPane(url, title) { return { Visible: false }; },
    FileDialog() { return { Title: "", AllowMultiSelect: false, Filters: { Clear() {}, Add() {} }, Show() { return -1; }, SelectedItems: { Count: 0, Item() { return ""; } } }; }
  };
  app.ActivePresentation = new MockPresentation(deck, deck.slides);
  deck.slides.forEach(function (s) { s.presentation = app.ActivePresentation; });
  app.Presentations.Add = function () { return new MockPresentation(deck, []); };
  app.Presentations.Open = function (p) { const d3 = deck; return new MockPresentation(d3, deck.slides); };
  return app;
}

// ---------- run tests ----------
async function main() {
  const results = [];
  function check(name, cond, detail) {
    results.push({ name, ok: !!cond, detail: detail || "" });
    console.log((cond ? "PASS " : "FAIL ") + name + (detail ? " | " + detail : ""));
  }

  // build deck: slides 1-3 image A with crops, slide 4 image B
  const deck = {
    fs: new MockFS(),
    slides: [],
    path: "C:/mock/deck.pptx",
    clipboard: null,
    selectedShape: null
  };
  deckRef = deck;
  const s1 = new MockSlide(deck, 1); deck.slides.push(s1);
  const s2 = new MockSlide(deck, 2); deck.slides.push(s2);
  const s3 = new MockSlide(deck, 3); deck.slides.push(s3);
  const s4 = new MockSlide(deck, 4); deck.slides.push(s4);

  const a1 = s1.AddPicture("C:/img/A.png", 0, -1, 30, 50, 240, 150);
  a1.name = "总平面-主图"; a1.lockAspectRatio = 0;
  a1.CropLeft = 25; a1.CropRight = 35; a1.CropTop = 15; a1.CropBottom = 45; a1.rotation = 20;
  const a2 = s2.AddPicture("C:/img/A.png", 0, -1, 100, 200, 200, 120);
  a2.lockAspectRatio = 0;
  a2.CropLeft = 0; a2.CropRight = 40; a2.CropTop = 10; a2.CropBottom = 10;
  const a3 = s3.AddPicture("C:/img/A.png", 0, -1, 400, 60, 180, 140);
  a3.lockAspectRatio = 0;
  a3.CropLeft = 10; a3.CropRight = 10; a3.CropTop = 30; a3.CropBottom = 10; a3.rotation = -15;
  const b1 = s4.AddPicture("C:/img/B.png", 0, -1, 500, 300, 220, 130);
  b1.lockAspectRatio = 0;
  b1.CropLeft = 20; b1.CropRight = 0; b1.CropTop = 0; b1.CropBottom = 20;

  deck.fs.writeAsBinaryString("C:/img/C.png", "PNGIMG_C_640x360:614400");
  deck.fs.writeAsBinaryString("C:/img/B.png", "PNGIMG_B_500x350:175000");
  const app = buildApp(deck);
  global.wps = app;
  global.btoa = (s) => Buffer.from(s, "binary").toString("base64");
  global.setTimeout = setTimeout;

  const src = fs.readFileSync(path.join(__dirname, "..", "main.js"), "utf8");
  eval(src);
  const W = global.WpsPictureReplace;

  // ---- test 1: collectDeckImages ----
  const collect = await W.collectDeckImages();
  check("collect groups count", collect.groups.length === 2, JSON.stringify(collect.groups.map(g => g.name + ":" + g.instances.length + ":" + g.linkState)));
  const groupA = collect.groups.find(g => g.name === "总平面-主图");
  check("group A has 3 instances", groupA && groupA.instances.length === 3, groupA ? String(groupA.instances.length) : "missing");
  check("group A linkState none", groupA && groupA.linkState === "none");
  check("group B linkState none", collect.groups.find(g => g.instances.length === 1) !== undefined);
  check("instances have thumbs", groupA.instances.every(i => !!i.thumb && i.thumb.startsWith("data:image/png;base64,")));
  check("instances have zones", groupA.instances.map(i => i.zone).join(","), groupA.instances.map(i => i.zone).join(","));
  check("hasCrop flags", groupA.instances.every(i => i.hasCrop === true));
  check("docKey present", !!collect.docKey);
  check("slideCount 4", collect.slideCount === 4);

  // ---- test 2: replaceInstances on two A instances ----
  const deck2 = deck; // same deck continues
  const targets = [groupA.instances[0], groupA.instances[1]];
  const rep = await W.replaceInstances(targets, "C:/img/C.png", collect.docKey);
  check("replaceInstances replaced 2", rep.replaced === 2 && rep.failed === 0, JSON.stringify(rep));
  const collect2 = await W.collectDeckImages();
  const groupC = collect2.groups.find(g => g.instances.length === 2);
  check("new group C has 2 instances (linked)", groupC && groupC.instances.length === 2 && groupC.linkState === "linked", groupC ? groupC.linkState : "missing");
  const hasSlide3 = collect2.groups.some(g => g.instances.some(i => i.uid === "3:1"));
  check("group A left 1 instance (slide3 uid retained)", hasSlide3);

  // ---- test 3: docKey guard ----
  let threw = false;
  try { await W.replaceInstances(targets, "C:/img/C.png", "wrong-doc-key"); } catch (e) { threw = true; }
  check("docKey mismatch throws", threw);

  // ---- test 4: source modified -> linkState modified -> update ----
  // mutate the C image file bytes in mock fs
  const cPath = "C:/img/C.png";
  const cBinary = W.baseName(cPath) + "-CHANGED-" + Date.now();
  deck.fs.writeAsBinaryString(cPath, cBinary);
  const collect3 = await W.collectDeckImages();
  const groupCmod = collect3.groups.find(g => g.instances.length === 2);
  check("modified state detected", groupCmod && groupCmod.linkState === "modified", groupCmod ? groupCmod.linkState : "missing");
  const upd = await W.updateLinkedInstances(groupCmod.instances, collect3.docKey);
  check("updateLinkedInstances updated 2", upd.updated === 2, JSON.stringify(upd));
  const collect4 = await W.collectDeckImages();
  const groupCok = collect4.groups.find(g => g.instances.length === 2);
  check("linkState linked after update", groupCok && groupCok.linkState === "linked", groupCok ? groupCok.linkState : "missing");

  // ---- test 5: rename + goto + unlink ----
  const inst5 = groupCok.instances[0];
  const renamed = W.renameShape(inst5.shape, "新版夜景鸟瞰图");
  check("renameShape ok", renamed === true);
  const altAfter = String(inst5.shape.AlternativeText);
  check("link JSON name synced", altAfter.includes("新版夜景鸟瞰图"), altAfter);
  const gone = W.gotoSlide(3);
  check("gotoSlide 3", gone === true && app.ActiveWindow.View.current === 3);
  const un = await W.unlinkInstances([groupCok.instances[1]], collect4.docKey);
  check("unlink 1", un === 1);
  const collect5 = await W.collectDeckImages();
  const still2 = collect5.groups.find(g => g.instances.length === 2);
  check("unlinked instance still grouped (linked state, row shows unlinked)", !!still2 && still2.linkState === "linked", still2 ? still2.linkState : "missing");

  // ---- test 6: batch replaceAllMatching regression ----
  const anyShape = collect5.groups.find(g => g.instances.length === 2).instances[0].shape;
  anyShape.Select();
  const batch = await W.replaceAllMatching(anyShape, "C:/img/B.png");
  check("batch replace matched 2", batch.matched === 2 && batch.success === 2, JSON.stringify(batch));

  // ---- test 7: index-shift safety on same slide ----
  // add two shapes on a fresh slide and replace both: pre-resolved refs must hit the right shapes
  const s5 = new MockSlide(deck, 5); deck.slides.push(s5);
  const x1 = s5.AddPicture("C:/img/A.png", 0, -1, 10, 10, 100, 80);
  x1.name = "X1";
  const x2 = s5.AddPicture("C:/img/A.png", 0, -1, 300, 10, 100, 80);
  x2.name = "X2";
  const collect6 = await W.collectDeckImages();
  const gX = collect6.groups.find(g => g.instances.some(i => i.uid === "5:1") && g.instances.some(i => i.uid === "5:2"));
  check("same-slide group found", !!gX, gX ? "in group " + gX.name : "missing");
  const xInsts = [gX.instances.find(i => i.uid === "5:1"), gX.instances.find(i => i.uid === "5:2")];
  const rX = await W.replaceInstances(xInsts, "C:/img/C.png", collect6.docKey);
  check("same-slide both replaced", rX.replaced === 2, JSON.stringify(rX));
  const namesAfter = s5.shapes.map(sh => sh.name).sort();
  check("right shapes replaced (names preserved X1/X2)", namesAfter.join(",") === "X1,X2", namesAfter.join(","));
  check("both now image C", s5.shapes.every(sh => sh.imageId === "C"), JSON.stringify(s5.shapes.map(sh => sh.imageId)));

  const failed = results.filter(r => !r.ok);
  console.log("\n===== " + (failed.length ? failed.length + " FAILURES" : "ALL TESTS PASSED") + " (" + results.length + " checks) =====");
  process.exit(failed.length ? 1 : 0);
}

main().catch(e => { console.error("HARNESS ERROR:", e); process.exit(2); });
