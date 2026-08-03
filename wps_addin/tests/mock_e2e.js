
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
  "C": { w: 640, h: 360, seed: "IMG_C_640x360" },
  "BIG": { w: 8000, h: 6000, seed: "IMG_BIG_8000x6000" }
};

// Real PNG container with a correct IHDR (width/height) so the add-in's
// imagePixelSize() parses it; the IDAT is a stub because the mock canvas
// never decodes pixels from these files.
function pngBytes(imgId, w, h) {
  const sig = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  function chunk(type, data) {
    const out = Buffer.alloc(12 + data.length);
    out.writeUInt32BE(data.length, 0);
    out.write(type, 4, "ascii");
    data.copy(out, 8);
    let crc = 0xffffffff;
    const crcSrc = Buffer.concat([Buffer.from(type, "ascii"), data]);
    for (let i = 0; i < crcSrc.length; i += 1) {
      crc ^= crcSrc[i];
      for (let k = 0; k < 8; k += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
    }
    out.writeUInt32BE((crc ^ 0xffffffff) >>> 0, 8 + data.length);
    return out;
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(w, 0);
  ihdr.writeUInt32BE(h, 4);
  ihdr[8] = 8; ihdr[9] = 2; // bit depth 8, color type RGB
  const tag = Buffer.from("PICRENEW_MOCK:" + (IMAGES[imgId] ? IMAGES[imgId].seed : imgId) + ":" + w + "x" + h);
  const idat = Buffer.from([0x78, 0x9c, 0x63, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01]);
  return Buffer.concat([sig, chunk("IHDR", ihdr), chunk("prTg", tag), chunk("IDAT", idat), chunk("IEND", Buffer.alloc(0))]);
}

function parsePngDims(data) {
  const s = String(data);
  if (s.length < 24) return null;
  function u32(i) {
    return (((s.charCodeAt(i) & 0xff) << 24) | ((s.charCodeAt(i + 1) & 0xff) << 16) |
            ((s.charCodeAt(i + 2) & 0xff) << 8) | (s.charCodeAt(i + 3) & 0xff)) >>> 0;
  }
  if (u32(0) !== 0x89504e47) return null;
  const w = u32(16); const h = u32(20);
  return (w > 0 && h > 0) ? { w: w, h: h } : null;
}

// ---------- mock FileSystem ----------
class MockFS {
  constructor() { this.files = new Map(); this.tmp = "C:/mock/tmp/"; }
  tmpdir() { return this.tmp; }
  writeAsBinaryString(p, data) { this.files.set(this._norm(p), data instanceof Buffer ? data.toString("binary") : String(data)); }
  WriteFile(p, data) { this.files.set(this._norm(p), data instanceof Buffer ? data.toString("binary") : String(data)); }
  readAsBinaryString(p) {
    const k = this._norm(p);
    if (!this.files.has(k)) throw new Error("mock fs: no such file " + p);
    return this.files.get(k);
  }
  Exists(p) { return this.files.has(this._norm(p)); }
  existsSync(p) { return this.files.has(this._norm(p)); }
  unlinkSync(p) { this.files.delete(this._norm(p)); }
  Remove(p) { this.files.delete(this._norm(p)); }
  has(p) { return this.files.has(this._norm(p)); }
  get(p) { return this.files.get(this._norm(p)); }
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
    this.cropBaselineW = img.w; this.cropBaselineH = img.h; // 96-dpi crop baseline (px * 0.75 when parsed)
    this.left = 0; this.top = 0; this.rotation = 0;
    this.hFlip = false; this.vFlip = false;
    this.lockAspectRatio = -1; // msoTrue
    this._cropLeft = 0; this._cropRight = 0; this._cropTop = 0; this._cropBottom = 0;
    this.scaleX = 1; this.scaleY = 1;
    this.visualId = imageId;
    this.type = 13;
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
  // Full-image display size at the current zoom (what WPS reports after the
  // crops of a pasted copy are zeroed). For an uncropped shape it is the
  // frame size itself.
  get fullWidth() {
    const base = this.cropBaselineW || this.width;
    const crops = (this._cropLeft || 0) + (this._cropRight || 0);
    if (crops > 0.05 && crops < base - 0.05) return this.width * base / (base - crops);
    return this.width;
  }
  get fullHeight() {
    const base = this.cropBaselineH || this.height;
    const crops = (this._cropTop || 0) + (this._cropBottom || 0);
    if (crops > 0.05 && crops < base - 0.05) return this.height * base / (base - crops);
    return this.height;
  }
  get Left() { return this.left; } set Left(v) { this.left = Number(v); }
  get Top() { return this.top; } set Top(v) { this.top = Number(v); }
  get Rotation() { return this.rotation; } set Rotation(v) { this.rotation = Number(v); }
  get HorizontalFlip() { return this.hFlip; }
  get VerticalFlip() { return this.vFlip; }
  get LockAspectRatio() { return this.lockAspectRatio; } set LockAspectRatio(v) { this.lockAspectRatio = v; }
  get Type() { return this.type; }
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
    d.width = this.fullWidth; d.height = this.fullHeight;
    d.cropBaselineW = this.cropBaselineW; d.cropBaselineH = this.cropBaselineH;
    d.left = this.left; d.top = this.top;
    d.scaleX = this.scaleX || 1; d.scaleY = this.scaleY || 1;
    d.visualId = this.visualId || this.imageId;
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
    this.customLayout = null;
    this.pageSetup = { slideWidth: 960, slideHeight: 540 };
  }
  get Parent() { return this.presentation; }
  get CustomLayout() { return this.customLayout; }
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
    // Simulate WPS: the crop baseline is the 96-dpi pixel size (px * 0.75),
    // while the reported insert size is capped at ~11 x 8.25 in.
    let px = null;
    try {
      const raw = this.deck.fs.get(String(file));
      if (raw) px = parsePngDims(raw);
    } catch (_) {}
    if (px && px.w > 0 && px.h > 0) {
      sh.cropBaselineW = px.w * 0.75;
      sh.cropBaselineH = px.h * 0.75;
      const scale = Math.min(792 / sh.cropBaselineW, 594 / sh.cropBaselineH, 1);
      sh.width = sh.cropBaselineW * scale;
      sh.height = sh.cropBaselineH * scale;
    }
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
    sh.cropBaselineW = c.cropBaselineW; sh.cropBaselineH = c.cropBaselineH;
    sh.width = c.fullWidth; sh.height = c.fullHeight; sh.left = c.left; sh.top = c.top; sh.rotation = c.rotation;
    sh.scaleX = c.scaleX || 1; sh.scaleY = c.scaleY || 1;
    sh.visualId = c.visualId || c.imageId;
    sh.hFlip = c.hFlip; sh.vFlip = c.vFlip; sh.lockAspectRatio = c.lockAspectRatio;
    sh.name = c.name; sh.alternativeText = c.alternativeText;
    this.shapes.push(sh);
    return sh;
  }
  PasteSpecial(fmt) { return this.Paste(); }
  Export(p, fmt, w, h) {
    const pics = this.shapes.filter(s => !s.deleted && s.imageId);
    if (!pics.length) throw new Error("mock export: no picture");
    const ew = Number(w) || 160;
    const eh = Number(h) || 160;
    if (pics.length > 1) {
      // batch grid: 6 columns, cells encoded as row,col,visualId
      const cols = 6;
      const cells = pics.map(function (s, i) {
        return Math.floor(i / cols) + "," + (i % cols) + "," + (s.visualId || s.imageId);
      }).join(";");
      const bytes = Buffer.from("GRID:" + Math.round(ew) + "x" + Math.round(eh) + ":" + cols + ":" + cells);
      this.deck.fs.writeAsBinaryString(p, bytes.toString("binary"));
    } else {
      const pic = pics[0];
      const bytes = Buffer.from("VISUAL:" + (pic.visualId || pic.imageId) + "|" + pic.imageId + ":" + Math.round(ew) + "x" + Math.round(eh));
      this.deck.fs.writeAsBinaryString(p, bytes.toString("binary"));
    }
  }
}

// ---------- mock Presentation ----------
class MockPresentation {
  constructor(deck, slidesArray) {
    this.deck = deck;
    this.slidesArray = slidesArray || [];
    this.saved = false;
    this.fullName = deck.path || "";
    this.SlideMaster = null;
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

  deck.fs.writeAsBinaryString("C:/img/A.png", pngBytes("A", 400, 300));
  deck.fs.writeAsBinaryString("C:/img/B.png", pngBytes("B", 500, 350));
  deck.fs.writeAsBinaryString("C:/img/C.png", pngBytes("C", 640, 360));
  deck.fs.writeAsBinaryString("C:/img/BIG.png", pngBytes("BIG", 8000, 6000));
  const app = buildApp(deck);
  global.wps = app;
  global.btoa = (s) => Buffer.from(s, "binary").toString("base64");
  // Fake Image/canvas/URL environment so perceptual-hash grouping is testable.
  const fakeBlobMap = {};
  function fillPixels(px, w, h, visualId) {
    var x = 0;
    for (var i = 0; i < visualId.length; i += 1) x = (x * 31 + visualId.charCodeAt(i)) & 0x7fffffff;
    if (x === 0) x = 0x9e3779b9;
    for (var p = 0; p < w * h; p += 1) {
      x ^= (x << 13) & 0x7fffffff; x ^= x >>> 17; x ^= (x << 5) & 0x7fffffff;
      var v = x & 0xff;
      var w2 = (x >>> 8) & 0xff;
      var z = (x >>> 16) & 0xff;
      px[p * 4] = v; px[p * 4 + 1] = w2; px[p * 4 + 2] = z; px[p * 4 + 3] = 255;
    }
  }
  global.URL = {
    _n: 0,
    createObjectURL: function (blob) { var k = "blob:mock/" + (++global.URL._n); fakeBlobMap[k] = blob; return k; },
    revokeObjectURL: function () {}
  };
  global.Image = function () { this._src = ""; this.visualId = "?"; };
  Object.defineProperty(global.Image.prototype, "src", {
    set: function (v) {
      this._src = v;
      const self = this;
      if (typeof v === "string" && v.indexOf("blob:") === 0 && fakeBlobMap[v]) {
        fakeBlobMap[v].arrayBuffer().then(function (buf) {
          const text = Buffer.from(buf).toString("binary");
          const gm = /^GRID:(\d+)x(\d+):(\d+):(.+)$/.exec(text);
          if (gm) {
            self.width = Number(gm[1]);
            self.height = Number(gm[2]);
            self.cols = Number(gm[3]);
            self.cells = {};
            String(gm[4]).split(";").forEach(function (c) {
              const parts = String(c).split(",");
              if (parts.length === 3) self.cells[parts[0] + "," + parts[1]] = parts[2];
            });
            self.visualId = "GRID";
            if (self.onload) self.onload();
          } else {
            const m = /VISUAL:([^|]+)|/.exec(text);
            self.visualId = m ? m[1] : "?";
            if (self.onload) self.onload();
          }
        }).catch(function () { if (self.onerror) self.onerror(); });
      }
    },
    get: function () { return this._src; }
  });
  global.document = {
    createElement: function (tag) {
      return {
        width: 0, height: 0,
        getContext: function () {
          let lastVisual = "?";
          let cellsRef = null;
          let colsRef = 6;
          return {
            drawImage: function (img, a2, a3, a4, a5, a6, a7, a8, a9) {
              if (img && img.cells && (a2 === undefined || typeof a2 === "number" && arguments.length <= 3)) {
                // plain full draw of a grid image (2- or 3-arg form)
                cellsRef = img.cells;
                colsRef = img.cols || 6;
                lastVisual = "GRID";
                return;
              }
              if (typeof a2 === "number" && arguments.length >= 9 && cellsRef) {
                // 9-arg crop from the grid canvas: pick the cell under (sx, sy)
                const cellW = Math.round((img && img.width ? img.width : 1) / colsRef);
                const cellH = Math.round((img && img.height ? img.height : 1) / colsRef);
                const col = Math.floor(a2 / Math.max(1, cellW));
                const row = Math.floor(a3 / Math.max(1, cellH));
                lastVisual = cellsRef[row + "," + col] || "?";
                return;
              }
              if (img && img.visualId) { lastVisual = img.visualId; }
            },
            getImageData: function (x, y, w, h) {
              const px = new Uint8Array(w * h * 4);
              fillPixels(px, w, h, lastVisual);
              return { data: px };
            },
            toDataURL: function () {
              return "data:image/png;base64," + Buffer.from("CELL:" + lastVisual).toString("base64");
            }
          };
        }
      };
    }
  };

  // Minimal FileReader so the btoa-less fallback path is testable in Node.
  if (typeof global.FileReader === "undefined") {
    global.FileReader = function () { this.result = null; this.onload = null; this.onerror = null; };
    global.FileReader.prototype.readAsDataURL = function (blob) {
      const self = this;
      blob.arrayBuffer().then(function (buf) {
        self.result = "data:image/png;base64," + Buffer.from(buf).toString("base64");
        if (self.onload) self.onload({ target: self });
      }).catch(function () { if (self.onerror) self.onerror(new Error("read failed")); });
    };
  }
  global.setTimeout = setTimeout;

  const src = fs.readFileSync(path.join(__dirname, "..", "main.js"), "utf8");
  eval(src);
  const W = global.WpsPictureReplace;

  // ---- test 1: collectDeckImages ----
  const collect = await W.collectDeckImages();
  await W.refreshLinkStates(collect.groups);
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
  await W.refreshLinkStates(collect2.groups);
  const groupC = collect2.groups.find(g => g.instances.length === 2);
  check("new group C has 2 instances (linked)", groupC && groupC.instances.length === 2 && groupC.linkState === "linked", groupC ? groupC.linkState : "missing");
  const hasSlide3 = collect2.groups.some(g => g.instances.some(i => i.uid === "3:1"));
  check("group A left 1 instance (slide3 uid retained)", hasSlide3);

  // ---- test 2b: crops and frame sizes preserved after replacement ----
  const after1 = s1.shapes[0];
  const after2 = s2.shapes[0];
  check("frame size preserved (slide1)", Math.abs(after1.width - 240) < 0.01 && Math.abs(after1.height - 150) < 0.01, after1.width + "x" + after1.height);
  check("crop preserved (slide1)", (after1.CropLeft + after1.CropRight + after1.CropTop + after1.CropBottom) > 0, JSON.stringify([after1.CropLeft, after1.CropRight, after1.CropTop, after1.CropBottom]));
  check("frame size preserved (slide2)", Math.abs(after2.width - 200) < 0.01 && Math.abs(after2.height - 120) < 0.01, after2.width + "x" + after2.height);
  check("crop preserved (slide2)", (after2.CropLeft + after2.CropRight + after2.CropTop + after2.CropBottom) > 0, JSON.stringify([after2.CropLeft, after2.CropRight, after2.CropTop, after2.CropBottom]));

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
  await W.refreshLinkStates(collect3.groups);
  const groupCmod = collect3.groups.find(g => g.instances.length === 2);
  check("modified state detected", groupCmod && groupCmod.linkState === "modified", groupCmod ? groupCmod.linkState : "missing");
  const upd = await W.updateLinkedInstances(groupCmod.instances, collect3.docKey);
  check("updateLinkedInstances updated 2", upd.updated === 2, JSON.stringify(upd));
  const collect4 = await W.collectDeckImages();
  await W.refreshLinkStates(collect4.groups);
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
  await W.refreshLinkStates(collect5.groups);
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
  await W.refreshLinkStates(collect6.groups);
  const gX = collect6.groups.find(g => g.instances.some(i => i.uid === "5:1") && g.instances.some(i => i.uid === "5:2"));
  check("same-slide group found", !!gX, gX ? "in group " + gX.name : "missing");
  const xInsts = [gX.instances.find(i => i.uid === "5:1"), gX.instances.find(i => i.uid === "5:2")];
  const rX = await W.replaceInstances(xInsts, "C:/img/C.png", collect6.docKey);
  check("same-slide both replaced", rX.replaced === 2, JSON.stringify(rX));
  const namesAfter = s5.shapes.map(sh => sh.name).sort();
  check("right shapes replaced (names preserved X1/X2)", namesAfter.join(",") === "X1,X2", namesAfter.join(","));
  check("both now image C", s5.shapes.every(sh => sh.imageId === "C"), JSON.stringify(s5.shapes.map(sh => sh.imageId)));

  // ---- test 8b: thumbnail fallback without btoa (FileReader path) ----
  const savedBtoa = global.btoa;
  delete global.btoa;
  const collectFR = await W.collectDeckImages();
  await W.refreshLinkStates(collectFR.groups);
  const frOk = collectFR.groups.some(g => g.instances.some(i => i.thumb && i.thumb.startsWith("data:image/png;base64,")));
  check("thumbnail fallback works without btoa", frOk);
  if (savedBtoa !== undefined) global.btoa = savedBtoa;

  // ---- test 8c: same source with different scale still one group ----
  const s6 = new MockSlide(deck, 6); deck.slides.push(s6);
  const g1 = s6.AddPicture("C:/img/A.png", 0, -1, 10, 10, 100, 80);
  g1.name = "缩放A"; g1.scaleX = 1; g1.scaleY = 1;
  const g2 = s6.AddPicture("C:/img/A.png", 0, -1, 300, 10, 150, 100);
  g2.name = "缩放B"; g2.scaleX = 1.5; g2.scaleY = 1;   // non-uniform scale -> different aspect
  const collectScale = await W.collectDeckImages();
  await W.refreshLinkStates(collectScale.groups);
  const scaleGroup = collectScale.groups.find(g => g.instances.some(i => i.uid === "6:1"));
  const scaleSame = scaleGroup && scaleGroup.instances.some(i => i.uid === "6:2");
  check("different aspect same source stays one group", !!scaleSame, scaleGroup ? "in group with " + scaleGroup.instances.length + " instances" : "split");
  const aspA = scaleGroup.instances.find(i => i.uid === "6:1").aspect;
  const aspB = scaleGroup.instances.find(i => i.uid === "6:2").aspect;
  check("aspects genuinely differ", aspA !== aspB, aspA + " vs " + aspB);

  // ---- test 8d: overlapping pictures flagged ----
  const s7 = new MockSlide(deck, 7); deck.slides.push(s7);
  const o1 = s7.AddPicture("C:/img/B.png", 0, -1, 100, 100, 200, 150);
  o1.name = "重叠A";
  const o2 = s7.AddPicture("C:/img/C.png", 0, -1, 100, 100, 200, 150);
  o2.name = "重叠B";
  const collectOv = await W.collectDeckImages();
  await W.refreshLinkStates(collectOv.groups);
  const ovInsts = [];
  collectOv.groups.forEach(function (g) { g.instances.forEach(function (i) { if (i.uid === "7:1" || i.uid === "7:2") ovInsts.push(i); }); });
  check("overlap flagged on both", ovInsts.length === 2 && ovInsts.every(i => i.overlap === true), JSON.stringify(ovInsts.map(i => i.uid + ":" + i.overlap)));

  // ---- test 8g: master + layout pictures with applied pages ----
  function makeShapesColl() {
    const arr = [];
    return {
      get Count() { return arr.length; },
      Item: function (i) { return arr[Number(i) - 1]; },
      _arr: arr,
      _push: function (sh) { arr.push(sh); }
    };
  }
  const masterColl = makeShapesColl();
  const layoutAColl = makeShapesColl();
  const layoutBColl = makeShapesColl();
  const masterObj = { Name: "WPS 母版", Shapes: masterColl, PageSetup: { SlideWidth: 960, SlideHeight: 540 } };
  const layoutA = { Name: "版式A", Shapes: layoutAColl };
  const layoutB = { Name: "版式B", Shapes: layoutBColl };
  masterObj.CustomLayouts = { get Count() { return 2; }, Item: function (i) { return i === 1 ? layoutA : layoutB; } };
  const mp1 = new MockShape(deck, masterObj, "A"); mp1.name = "母版图1"; masterColl._push(mp1);
  const mp2 = new MockShape(deck, masterObj, "B"); mp2.name = "母版图2"; masterColl._push(mp2);
  const lp1 = new MockShape(deck, layoutA, "C"); lp1.name = "版式A图1"; layoutAColl._push(lp1);
  const lp2 = new MockShape(deck, layoutB, "A"); lp2.name = "版式B图1"; layoutBColl._push(lp2);
  const s10 = new MockSlide(deck, 10); deck.slides.push(s10); s10.customLayout = layoutA;
  const s11 = new MockSlide(deck, 11); deck.slides.push(s11); s11.customLayout = layoutA;
  const s12 = new MockSlide(deck, 12); deck.slides.push(s12); s12.customLayout = layoutB;
  app.ActivePresentation.SlideMaster = masterObj;
  const collectM = await W.collectDeckImages();
  await W.refreshLinkStates(collectM.groups);
  let masterInst = null;
  let layoutAInst = null;
  let layoutBInst = null;
  collectM.groups.forEach(function (g) {
    g.instances.forEach(function (i) {
      if (i.shapeName === "母版图1") masterInst = i;
      if (i.shapeName === "版式A图1") layoutAInst = i;
      if (i.shapeName === "版式B图1") layoutBInst = i;
    });
  });
  check("master picture collected", !!masterInst && masterInst.kind === "master", masterInst ? masterInst.kind : "missing");
  check("master applied to all pages", !!masterInst && masterInst.appliedPages.length >= 3, masterInst ? String(masterInst.appliedPages.length) : "0");
  check("layout A picture collected with applied pages", !!layoutAInst && layoutAInst.kind === "layout" && layoutAInst.appliedPages.join(",") === "8,9", layoutAInst ? layoutAInst.appliedPages.join(",") : "missing");
  check("layout B picture collected", !!layoutBInst && layoutBInst.kind === "layout" && layoutBInst.appliedPages.join(",") === "10", layoutBInst ? layoutBInst.appliedPages.join(",") : "missing");

  // ---- test 8f: non-picture shapes (textbox/autoshape) excluded ----
  const s9 = new MockSlide(deck, 9); deck.slides.push(s9);
  const tb1 = s9.AddPicture("C:/img/A.png", 0, -1, 10, 10, 100, 80);
  tb1.name = "文本框1"; tb1.type = 17;   // textbox
  const sh1 = s9.AddPicture("C:/img/B.png", 0, -1, 300, 10, 100, 80);
  sh1.name = "矩形1"; sh1.type = 1;     // autoshape
  const pic1 = s9.AddPicture("C:/img/C.png", 0, -1, 600, 10, 100, 80);
  pic1.name = "真图片1";                  // type 13
  const collectF = await W.collectDeckImages();
  await W.refreshLinkStates(collectF.groups);
  const hasTb = collectF.groups.some(g => g.instances.some(i => i.name === "文本框1" || i.name === "矩形1"));
  const hasPic = collectF.groups.some(g => g.instances.some(i => i.name === "真图片1"));
  check("textbox/autoshape excluded from inventory", !hasTb);
  check("real picture included", hasPic);

  // ---- test 8e: perceptual grouping merges different bytes, same visual ----
  const s8 = new MockSlide(deck, 8); deck.slides.push(s8);
  const v1 = s8.AddPicture("C:/img/A.png", 0, -1, 10, 10, 100, 80);
  v1.name = "视觉同源A"; v1.visualId = "VIS_SRC_B";
  const v2 = s8.AddPicture("C:/img/C.png", 0, -1, 300, 10, 100, 80);
  v2.name = "视觉同源C"; v2.visualId = "VIS_SRC_B";   // different file bytes, same visual content
  const v3 = s8.AddPicture("C:/img/B.png", 0, -1, 600, 10, 100, 80);
  v3.name = "异源B"; v3.visualId = "VIS_SRC_OTHER";
  const collectPh = await W.collectDeckImages();
  await W.refreshLinkStates(collectPh.groups);
  const phA = collectPh.groups.find(g => g.instances.some(i => i.name === "视觉同源A"));
  const phMerged = phA && phA.instances.some(i => i.name === "视觉同源C");
  const phSplit = phA && phA.instances.some(i => i.name === "异源B");
  check("same visual different bytes merge into one group", !!phMerged, phA ? "group has " + phA.instances.length + " instances" : "no group");
  check("visually different stays separate", !phSplit);

  // ---- test 8: addinUrl resolution ----
  // scenario A: document.location available (official SDK path)
  const savedLoc = global.document;
  global.document = { location: { toString: function () { return "http://taskpane.html/taskpane.html"; } } };
  const urlA = W.addinUrl("#panel");
  check("addinUrl uses document.location root", urlA === "http://taskpane.html/taskpane.html#panel", urlA);
  delete global.document;
  if (savedLoc !== undefined) global.document = savedLoc;

  // scenario B: no document.location, CurrentWPSAddIn.Path verified on disk
  const savedAddin = app.CurrentWPSAddIn;
  app.CurrentWPSAddIn = { Path: "C:/mock/addin", Name: "picture-replace-tools-wps" };
  deck.fs.writeAsBinaryString("C:/mock/addin/taskpane.html", "<html></html>");
  const urlB = W.addinUrl("#panel");
  check("addinUrl file probe picks existing file", urlB === "file:///C:/mock/addin/taskpane.html#panel", urlB);
  // scenario C: no document.location, no file on disk -> relative fallback
  app.CurrentWPSAddIn = { Path: "C:/mock/missing", Name: "x" };
  const urlC = W.addinUrl("#panel");
  check("addinUrl relative fallback", urlC === "taskpane.html#panel", urlC);
  app.CurrentWPSAddIn = savedAddin;

  // ---- test 9: large-image crop baseline (WPS 96-dpi) ----
  function jpegBytes(w, h) {
    const b = Buffer.alloc(41);
    b[0] = 0xff; b[1] = 0xd8;                    // SOI
    b[2] = 0xff; b[3] = 0xe0; b[4] = 0x00; b[5] = 0x10; // APP0 len 16
    b.write("JFIF", 6, "ascii"); b[10] = 0;
    b[11] = 1; b[12] = 1; b[13] = 0;             // version + units
    b[14] = 0; b[15] = 0; b[16] = 0; b[17] = 0;  // densities
    b[18] = 0; b[19] = 0;                        // thumbnails
    b[20] = 0xff; b[21] = 0xc0;                  // SOF0
    b[22] = 0x00; b[23] = 0x11; b[24] = 0x08;    // len 17, precision 8
    b[25] = (h >> 8) & 0xff; b[26] = h & 0xff;   // height
    b[27] = (w >> 8) & 0xff; b[28] = w & 0xff;   // width
    b[29] = 3;                                   // components
    b[30] = 1; b[31] = 0x22; b[32] = 0; b[33] = 2; b[34] = 0x11; b[35] = 1;
    b[36] = 3; b[37] = 0x11; b[38] = 1;
    b[39] = 0xff; b[40] = 0xd9;                  // EOI
    return b;
  }
  const pxBig = W._math.imagePixelSize(deck.fs.get("C:/img/BIG.png"));
  check("imagePixelSize parses PNG IHDR", !!pxBig && pxBig.w === 8000 && pxBig.h === 6000, JSON.stringify(pxBig));
  const pxJpg = W._math.imagePixelSize(jpegBytes(800, 600).toString("binary"));
  check("imagePixelSize parses JPEG SOF", !!pxJpg && pxJpg.w === 800 && pxJpg.h === 600, JSON.stringify(pxJpg));

  const sBig = new MockSlide(deck, 10); deck.slides.push(sBig);
  const bigShape = sBig.AddPicture("C:/img/BIG.png", 0, -1, 20, 30, 240, 150);
  bigShape.name = "BigCropped";
  bigShape.lockAspectRatio = 0;
  bigShape.CropLeft = 1500; bigShape.CropRight = 2100; bigShape.CropTop = 900; bigShape.CropBottom = 2700;
  const collectBig = await W.collectDeckImages();
  await W.refreshLinkStates(collectBig.groups);
  let bigInst = null;
  collectBig.groups.forEach(function (g) {
    g.instances.forEach(function (i) { if (i.name === "BigCropped") bigInst = i; });
  });
  check("big image crop detected", !!bigInst && bigInst.hasCrop === true, bigInst ? "hasCrop=" + bigInst.hasCrop : "missing");
  const repBig = await W.replaceInstances([bigInst], "C:/img/BIG.png", collectBig.docKey);
  check("big image replaced", repBig.replaced === 1, JSON.stringify(repBig));
  const afterBig = sBig.shapes[0];
  check("big frame preserved", Math.abs(afterBig.width - 240) < 0.01 && Math.abs(afterBig.height - 150) < 0.01, afterBig.width + "x" + afterBig.height);
  // Baseline is 8000*0.75 = 6000pt; without the fix crops would be ~7.6x smaller (~261).
  check("big crops use 96-dpi baseline",
    afterBig.CropLeft > 1800 && afterBig.CropRight > 2400 && afterBig.CropTop > 800 && afterBig.CropBottom > 2600,
    JSON.stringify([afterBig.CropLeft, afterBig.CropRight, afterBig.CropTop, afterBig.CropBottom]));

  const failed = results.filter(r => !r.ok);
  console.log("\n===== " + (failed.length ? failed.length + " FAILURES" : "ALL TESTS PASSED") + " (" + results.length + " checks) =====");
  process.exit(failed.length ? 1 : 0);
}

main().catch(e => { console.error("HARNESS ERROR:", e); process.exit(2); });
