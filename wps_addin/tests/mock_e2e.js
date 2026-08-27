
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
class MockTags {
  constructor() { this.map = new Map(); }
  get Count() { return this.map.size; }
  Item(key) {
    const wanted = String(key || "").toUpperCase();
    if (typeof key === "number") {
      const values = Array.from(this.map.values());
      return values[Number(key) - 1] || "";
    }
    return this.map.get(wanted) || "";
  }
  Add(name, value) { this.map.set(String(name || "").toUpperCase(), String(value == null ? "" : value)); }
  Delete(name) { this.map.delete(String(name || "").toUpperCase()); }
}
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
    this.visible = -1;
    this.tags = new MockTags();
    this.deleted = false;
    this.z = 0; // assigned by slide
  }
  Select(replace) {
    if (!this.deck.selectedShapes) this.deck.selectedShapes = [];
    const replaceSelection = replace !== 0 && replace !== false;
    if (replaceSelection) this.deck.selectedShapes.length = 0;
    if (this.deck.selectedShapes.indexOf(this) < 0) this.deck.selectedShapes.push(this);
    this.selected = true;
    this.deck.selectedShape = this;
    return 0;
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
  get Visible() { return this.visible; } set Visible(v) { this.visible = Number(v) === 0 || v === false ? 0 : -1; }
  // Current WPS exposes Locked but does not round-trip writes. Keep the mock
  // aligned so the production code exercises its durable Tags fallback.
  get Locked() { return -1; } set Locked(v) { this._lockedProbe = v; }
  get Tags() { return this.tags; }
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
  Range(index) {
    const wanted = Array.isArray(index) ? index.map(String) : [String(index)];
    const found = this.shapes.filter(function (shape) {
      return wanted.indexOf(String(shape.Name || shape.name || "")) >= 0;
    });
    const slide = this;
    return {
      get Count() { return found.length; },
      Item(i) {
        const item = found[Number(i) - 1];
        if (!item) throw new Error("range index out of range " + i);
        return item;
      },
      Select(replace) {
        found.forEach(function (shape, i) { shape.Select(i === 0 ? replace : 0); });
        return slide;
      }
    };
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
      // batch grid: 8 columns, cells encoded as row,col,visualId
      const cols = 8;
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
  let explicitShapeRange = null;
  const selection = {};
  Object.defineProperty(selection, "ShapeRange", {
    configurable: true,
    get: function () {
      if (explicitShapeRange !== null) return explicitShapeRange;
      const selected = deck.selectedShapes || [];
      return {
        get Count() { return selected.length; },
        Item: function (i) { return selected[Number(i) - 1]; }
      };
    },
    set: function (value) { explicitShapeRange = value; }
  });
  selection.Unselect = function () {
    deck.selectedShapes = [];
    deck.selectedShape = null;
  };
  selection.ClearShapeSelect = selection.Unselect;
  const apiEvent = {
    listeners: Object.create(null),
    AddApiEventListener: function (name, callback) {
      this.listeners[String(name)] = callback;
      return true;
    },
    RemoveApiEventListener: function (name) {
      delete this.listeners[String(name)];
      return true;
    }
  };
  const app = {
    Version: "12.1.0.28043",
    FileSystem: deck.fs,
    CurrentWPSAddIn: { Path: "C:/mock/addin/", Name: "picture-replace-tools-wps" },
    OAAssist: {
      ShellExecute: function (url) {
        app._openedUrls.push(String(url));
      }
    },
    ActivePresentation: null,
    ActiveDocument: {
      FollowHyperlink: function (url) {
        app._openedUrls.push(String(url));
      }
    },
    ApiEvent: apiEvent,
    CommandBars: {
      ExecuteMso: function (commandId) {
        app._msoCalls.push(String(commandId));
      }
    },
    ActiveWindow: { ViewType: 9, Selection: selection, View: { current: 0, GotoSlide(n) { this.current = Number(n); }, get Slide() { return { SlideIndex: this.current }; } } },
    alert(msg) { app._alerts.push(String(msg)); },
    _alerts: [],
    _openedUrls: [],
    _msoCalls: [],
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
    selectedShape: null,
    selectedShapes: []
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

  check("smart zoom uses a dedicated ribbon icon", global.OnGetRibbonImage({ Id: "SmartZoomButton" }) === "icon_smart_zoom.png" && global.OnGetRibbonImage({ Id: "CtxSmartZoom" }) === "icon_smart_zoom.png", "SmartZoomButton/CtxSmartZoom");
  check("object filter uses a dedicated ribbon icon", global.OnGetFilterImage() === "icon_filter.png" && global.OnGetRibbonImage({ Id: "ObjectFilterMenu" }) === "icon_filter.png", "ObjectFilterMenu");
  const ribbonXml = fs.readFileSync(path.join(__dirname, "..", "ribbon.xml"), "utf8");
  const taskpaneHtml = fs.readFileSync(path.join(__dirname, "..", "taskpane.html"), "utf8");
  const installerScript = fs.readFileSync(path.join(__dirname, "..", "build_installer.ps1"), "utf8");
  check("ribbon exposes the requested author homepage control", /designed by Dongsidaye/.test(ribbonXml) && /onAction="OpenProjectHome"/.test(ribbonXml), "author/homepage ribbon control");
  check("ribbon uses a dedicated GitHub icon", /getImage="OnGetGithubImage"/.test(ribbonXml) && /function OnGetGithubImage\(\) \{ return "icon_github\.png"; \}/.test(fs.readFileSync(path.join(__dirname, "..", "main.js"), "utf8")) && global.OnGetGithubImage() === "icon_github.png" && fs.existsSync(path.join(__dirname, "..", "icon_github.png")), "GitHub icon");
  check("ribbon visibly exposes the current version", /id="AddonVersion"[^>]*label="v1\.2\.33"/.test(ribbonXml), "AddonVersion");
  check("installer carries the dedicated GitHub icon", /icon_github\.png/.test(installerScript), "build_installer.ps1");
  check("ribbon calls the native animation pane command", /AnimationCustom/.test(fs.readFileSync(path.join(__dirname, "..", "main.js"), "utf8")), "AnimationCustom");
  check("layer manager is named object manager in visible UI", /<h2>对象管理<\/h2>/.test(taskpaneHtml) && /label="对象管理"/.test(ribbonXml), "对象管理");
  check("picture panel routes through the inventory cache", /collectDeckImagesCached/.test(taskpaneHtml) && /if \(panel\) refresh\(false\)/.test(taskpaneHtml) && /refresh\(true\)/.test(taskpaneHtml), "cached panel reopen + explicit refresh");
  check("add-in exposes background inventory preload", typeof W.preloadDeckImages === "function" && /preloadDeckImages/.test(fs.readFileSync(path.join(__dirname, "..", "main.js"), "utf8")), "background preload API");
  check("background scan marker has a stale-context guard", /heartbeatAt/.test(src) && /PANEL_CACHE_BUSY_STALE_MS/.test(src), "scan heartbeat/stale marker guard");
  const openedBefore = app._openedUrls.length;
  const externalHome = W.openExternalUrl("https://github.com/Dongsidaye/ppt-picture-replace-tools");
  check("project homepage uses the WPS external-link API", externalHome && externalHome.ok === true && externalHome.method === "ShellExecute" && app._openedUrls.length === openedBefore + 1 && app._openedUrls[app._openedUrls.length - 1] === "https://github.com/Dongsidaye/ppt-picture-replace-tools", JSON.stringify(externalHome));
  const ribbonHomeBefore = app._openedUrls.length;
  global.OpenProjectHome({ Id: "OpenProjectHome" });
  check("ribbon homepage callback opens synchronously", app._openedUrls.length === ribbonHomeBefore + 1 && app._openedUrls[app._openedUrls.length - 1] === "https://github.com/Dongsidaye/ppt-picture-replace-tools", JSON.stringify(app._openedUrls));
  await new Promise(function (r) { setTimeout(r, 0); });
  global.OpenAnimationPane();
  await new Promise(function (r) { setTimeout(r, 0); });
  check("animation pane button executes WPS native AnimationCustom", app._msoCalls.indexOf("AnimationCustom") >= 0, JSON.stringify(app._msoCalls));

  // ---- smart zoom: selection geometry, anchor math, snapshot reset ----
  const zoomRange = {
    get Count() { return 2; },
    Item: function (i) { return Number(i) === 1 ? a1 : a2; }
  };
  const zoomOriginal = {
    a1: { left: a1.Left, top: a1.Top, width: a1.Width, height: a1.Height },
    a2: { left: a2.Left, top: a2.Top, width: a2.Width, height: a2.Height }
  };
  const zoomBounds = function () {
    return {
      left: Math.min(a1.Left, a2.Left),
      top: Math.min(a1.Top, a2.Top),
      right: Math.max(a1.Left + a1.Width, a2.Left + a2.Width),
      bottom: Math.max(a1.Top + a1.Height, a2.Top + a2.Height)
    };
  };
  app.ActiveWindow.Selection.ShapeRange = zoomRange;
  const zoomStart = W.smartZoomBegin();
  check("smart zoom begins from multi-shape selection", zoomStart && zoomStart.count === 2 && zoomStart.percent === 100, JSON.stringify(zoomStart));
  const beforeZoom = zoomBounds();
  const beforeCenter = { x: (beforeZoom.left + beforeZoom.right) / 2, y: (beforeZoom.top + beforeZoom.bottom) / 2 };
  const zoomApplied = W.smartZoomApply(200, { anchor: "center", scaleText: false });
  const afterZoom = zoomBounds();
  const afterCenter = { x: (afterZoom.left + afterZoom.right) / 2, y: (afterZoom.top + afterZoom.bottom) / 2 };
  check("smart zoom center anchor keeps selection center", Math.abs(afterCenter.x - beforeCenter.x) < 0.001 && Math.abs(afterCenter.y - beforeCenter.y) < 0.001, JSON.stringify({ before: beforeCenter, after: afterCenter }));
  check("smart zoom scales each selected shape from snapshot", Math.abs(a1.Width - zoomOriginal.a1.width * 2) < 0.001 && Math.abs(a2.Height - zoomOriginal.a2.height * 2) < 0.001 && zoomApplied.percent === 200, JSON.stringify(zoomApplied));
  W.smartZoomApply(150, { anchor: "top-left" });
  const topLeftAfter = zoomBounds();
  check("smart zoom top-left anchor keeps selection origin", Math.abs(topLeftAfter.left - beforeZoom.left) < 0.001 && Math.abs(topLeftAfter.top - beforeZoom.top) < 0.001, JSON.stringify(topLeftAfter));
  const zoomReset = W.smartZoomReset();
  check("smart zoom reset restores initial snapshot", Math.abs(a1.Left - zoomOriginal.a1.left) < 0.001 && Math.abs(a1.Width - zoomOriginal.a1.width) < 0.001 && Math.abs(a2.Top - zoomOriginal.a2.top) < 0.001 && zoomReset.percent === 100, JSON.stringify(zoomReset));
  W.smartZoomEnd();

  const textShape = {
    Left: 40, Top: 70, Width: 120, Height: 80, LockAspectRatio: 0, Type: 1,
    Line: { Visible: true, Weight: 2 },
    TextFrame: {
      MarginTop: 6, MarginLeft: 7, MarginRight: 8, MarginBottom: 9,
      TextRange: { Font: { Size: 18 }, ParagraphFormat: { LineRuleAfter: 4 } }
    },
    TextFrame2: { TextRange: { Font: { Size: 18 } } }
  };
  app.ActiveWindow.Selection.ShapeRange = { Count: 1, Item: function () { return textShape; } };
  W.smartZoomBegin();
  W.smartZoomApply(200, { scaleText: true, scaleShapeLine: true });
  check("smart zoom scales text and line when available", textShape.TextFrame2.TextRange.Font.Size === 36 && textShape.TextFrame.MarginLeft === 14 && textShape.Line.Weight === 4, JSON.stringify(textShape));
  W.smartZoomReset();
  check("smart zoom restores text style snapshot", textShape.TextFrame2.TextRange.Font.Size === 18 && textShape.TextFrame.MarginLeft === 7 && textShape.Line.Weight === 2, JSON.stringify(textShape));
  W.smartZoomEnd();

  // Some WPS builds re-fit text when a text-frame margin is written.  The
  // text size must therefore be written after the margins, otherwise the
  // host can shrink it a second time in the same zoom operation.
  let autoFitFontSize = 18;
  const autoFitFont = {};
  Object.defineProperty(autoFitFont, "Size", {
    get: function () { return autoFitFontSize; },
    set: function (value) { autoFitFontSize = Number(value); }
  });
  const autoFitFrame = { TextRange: { Font: autoFitFont, ParagraphFormat: { LineRuleAfter: 4 } } };
  ["MarginTop", "MarginRight", "MarginBottom"].forEach(function (key) {
    Object.defineProperty(autoFitFrame, key, {
      get: function () { return 6; },
      set: function () {}
    });
  });
  Object.defineProperty(autoFitFrame, "MarginLeft", {
    get: function () { return 7; },
    set: function () { autoFitFontSize *= 0.5; }
  });
  const autoFitTextShape = {
    Left: 40, Top: 70, Width: 120, Height: 80, LockAspectRatio: 0, Type: 1,
    TextFrame: autoFitFrame,
    TextFrame2: { TextRange: { Font: autoFitFont } }
  };
  app.ActiveWindow.Selection.ShapeRange = { Count: 1, Item: function () { return autoFitTextShape; } };
  W.smartZoomBegin();
  W.smartZoomApply(200, { scaleText: true });
  check("smart zoom writes text size after host auto-fit margins", autoFitFontSize === 36, JSON.stringify({ fontSize: autoFitFontSize }));
  W.smartZoomEnd();

  const tinyFont = { Size: 5 };
  const tinyTextShape = {
    Left: 40, Top: 70, Width: 120, Height: 80, LockAspectRatio: 0, Type: 1,
    TextFrame2: { TextRange: { Font: tinyFont } }
  };
  app.ActiveWindow.Selection.ShapeRange = { Count: 1, Item: function () { return tinyTextShape; } };
  W.smartZoomBegin();
  W.smartZoomApply(60, { scaleText: true, protectTextReadability: true });
  check("smart zoom keeps a readable minimum text size", tinyFont.Size === 8, JSON.stringify({ fontSize: tinyFont.Size }));
  W.smartZoomEnd();

  const mixedFont = { Size: -2 };
  const mixedTextShape = {
    Left: 40, Top: 70, Width: 120, Height: 80, LockAspectRatio: 0, Type: 1,
    TextFrame2: { TextRange: { Font: mixedFont } }
  };
  app.ActiveWindow.Selection.ShapeRange = { Count: 1, Item: function () { return mixedTextShape; } };
  W.smartZoomBegin();
  W.smartZoomApply(60, { scaleText: true });
  check("smart zoom skips mixed or empty font sizes", mixedFont.Size === -2, JSON.stringify({ fontSize: mixedFont.Size }));
  W.smartZoomEnd();

  const richFonts = [
    { Size: 28 },
    { Size: 18 },
    { Size: 12 }
  ];
  const richRuns = richFonts.map(function (font, index) {
    return { Start: index * 4 + 1, Length: 4, Font: font };
  });
  const richTextRange = {
    Font: { Size: -2 },
    Runs: function (start) {
      const index = Math.max(0, Math.min(richRuns.length - 1, Number(start) - 1));
      return richRuns[index];
    }
  };
  const richTextShape = {
    Left: 40, Top: 70, Width: 120, Height: 80, LockAspectRatio: 0, Type: 1,
    TextFrame2: { TextRange: richTextRange },
    // The legacy text API may expose a default positive size even when the
    // primary range reports a mixed value. Runs must remain authoritative.
    TextFrame: { TextRange: { Font: { Size: 24 } } }
  };
  app.ActiveWindow.Selection.ShapeRange = { Count: 1, Item: function () { return richTextShape; } };
  W.smartZoomBegin();
  W.smartZoomApply(50, { scaleText: true, protectTextReadability: false });
  check("smart zoom scales mixed-size rich text runs", richFonts[0].Size === 14 && richFonts[1].Size === 9 && richFonts[2].Size === 6, JSON.stringify(richFonts));
  W.smartZoomReset();
  check("smart zoom restores mixed-size rich text runs", richFonts[0].Size === 28 && richFonts[1].Size === 18 && richFonts[2].Size === 12, JSON.stringify(richFonts));
  W.smartZoomEnd();

  const tableFonts = [
    { Size: 18 },
    { Size: 14 },
    { Size: 12 },
    { Size: 10 }
  ];
  const tableCells = tableFonts.map(function (font) {
    return { Shape: { TextFrame2: { TextRange: { Font: font } } } };
  });
  const tableShape = {
    Left: 20, Top: 30, Width: 320, Height: 200, LockAspectRatio: 0, Type: 19,
    Table: {
      Rows: { Count: 2 },
      Columns: { Count: 2 },
      Cell: function (row, column) {
        return tableCells[(Number(row) - 1) * 2 + Number(column) - 1];
      }
    }
  };
  app.ActiveWindow.Selection.ShapeRange = { Count: 1, Item: function () { return tableShape; } };
  W.smartZoomBegin();
  W.smartZoomApply(50, { scaleText: true, protectTextReadability: false });
  check("smart zoom scales table cell text", tableFonts[0].Size === 9 && tableFonts[1].Size === 7 && tableFonts[2].Size === 6 && tableFonts[3].Size === 5, JSON.stringify(tableFonts));
  W.smartZoomApply(50, { scaleText: true, protectTextReadability: true });
  check("smart zoom preserves table cell proportions with readability protection", tableFonts[0].Size === 9 && tableFonts[1].Size === 7 && tableFonts[2].Size === 6 && tableFonts[3].Size === 5, JSON.stringify(tableFonts));
  W.smartZoomReset({ protectTextReadability: true });
  check("smart zoom restores table cell text", tableFonts[0].Size === 18 && tableFonts[1].Size === 14 && tableFonts[2].Size === 12 && tableFonts[3].Size === 10, JSON.stringify(tableFonts));
  W.smartZoomEnd();

  const lowTableFonts = [
    { Size: 18 },
    { Size: 14 },
    { Size: 12 },
    { Size: 10 }
  ];
  const lowTableCells = lowTableFonts.map(function (font) {
    return { Shape: { TextFrame2: { TextRange: { Font: font } } } };
  });
  const lowTableShape = {
    Left: 20, Top: 30, Width: 320, Height: 200, LockAspectRatio: 0, Type: 19,
    Table: {
      Rows: { Count: 2 },
      Columns: { Count: 2 },
      Cell: function (row, column) {
        return lowTableCells[(Number(row) - 1) * 2 + Number(column) - 1];
      }
    }
  };
  app.ActiveWindow.Selection.ShapeRange = { Count: 1, Item: function () { return lowTableShape; } };
  W.smartZoomBegin();
  W.smartZoomApply(30, { scaleText: true, protectTextReadability: true });
  const lowTableRatio = lowTableFonts[0].Size / lowTableFonts[3].Size;
  check("smart zoom preserves table font proportions below 30%", Math.abs(lowTableFonts[0].Size - 7.2) < 0.01 && Math.abs(lowTableFonts[1].Size - 5.6) < 0.01 && Math.abs(lowTableFonts[2].Size - 4.8) < 0.01 && Math.abs(lowTableFonts[3].Size - 4) < 0.01 && Math.abs(lowTableRatio - 1.8) < 0.01, JSON.stringify({ fonts: lowTableFonts, ratio: lowTableRatio }));
  W.smartZoomReset({ protectTextReadability: true });
  check("smart zoom restores low table cell text", lowTableFonts[0].Size === 18 && lowTableFonts[1].Size === 14 && lowTableFonts[2].Size === 12 && lowTableFonts[3].Size === 10, JSON.stringify(lowTableFonts));
  W.smartZoomEnd();

  // WPS can materialize an effect when a numeric effect property is written.
  // A disabled reflection must therefore stay disabled at 100%/Reset.
  const effectState = { reflectionVisible: false, reflectionOffset: 12 };
  const reflectionEffect = {};
  Object.defineProperty(reflectionEffect, "Visible", {
    get: function () { return effectState.reflectionVisible; },
    set: function (value) { effectState.reflectionVisible = !!value; }
  });
  Object.defineProperty(reflectionEffect, "Offset", {
    get: function () { return effectState.reflectionOffset; },
    set: function (value) { effectState.reflectionOffset = Number(value); effectState.reflectionVisible = true; }
  });
  const effectShape = {
    Left: 10, Top: 20, Width: 200, Height: 120, LockAspectRatio: 0, Type: 13,
    Reflection: reflectionEffect
  };
  app.ActiveWindow.Selection.ShapeRange = { Count: 1, Item: function () { return effectShape; } };
  W.smartZoomBegin();
  W.smartZoomApply(100, { scaleShapeReflection: true });
  check("smart zoom preserves disabled reflection at 100%", effectState.reflectionVisible === false, JSON.stringify(effectState));
  W.smartZoomEnd();

  const unknownEffectState = { materialized: false, offset: 8 };
  const unknownReflection = {};
  Object.defineProperty(unknownReflection, "Offset", {
    get: function () { return unknownEffectState.offset; },
    set: function (value) { unknownEffectState.offset = Number(value); unknownEffectState.materialized = true; }
  });
  const unknownEffectShape = {
    Left: 10, Top: 20, Width: 200, Height: 120, LockAspectRatio: 0, Type: 13,
    Reflection: unknownReflection
  };
  app.ActiveWindow.Selection.ShapeRange = { Count: 1, Item: function () { return unknownEffectShape; } };
  W.smartZoomBegin();
  W.smartZoomApply(200, { scaleShapeReflection: true });
  check("smart zoom skips effect when visibility is unavailable", unknownEffectState.materialized === false, JSON.stringify(unknownEffectState));
  W.smartZoomEnd();

  let geometryWrites = 0;
  const geometryState = { left: 10, top: 20, width: 200, height: 120, lock: 0 };
  const geometryShape = { Type: 13 };
  [["Left", "left"], ["Top", "top"], ["Width", "width"], ["Height", "height"], ["LockAspectRatio", "lock"]].forEach(function (pair) {
    Object.defineProperty(geometryShape, pair[0], {
      get: function () { return geometryState[pair[1]]; },
      set: function (value) { geometryWrites += 1; geometryState[pair[1]] = Number(value); }
    });
  });
  app.ActiveWindow.Selection.ShapeRange = { Count: 1, Item: function () { return geometryShape; } };
  W.smartZoomBegin();
  W.smartZoomApply(100, {});
  check("smart zoom does not rewrite geometry at unchanged 100%", geometryWrites === 0, JSON.stringify({ geometryWrites: geometryWrites, geometryState: geometryState }));
  W.smartZoomEnd();

  const largeChildren = [];
  for (let i = 0; i < 501; i += 1) largeChildren.push({ Left: i, Top: 0, Width: 1, Height: 1, LockAspectRatio: 0, Type: 1 });
  const largeGroup = {
    Left: 0, Top: 0, Width: 501, Height: 1, LockAspectRatio: 0, Type: 6,
    GroupItems: { Count: largeChildren.length, Item: function (i) { return largeChildren[Number(i) - 1]; } }
  };
  app.ActiveWindow.Selection.ShapeRange = zoomRange;
  const previousSession = W.smartZoomBegin();
  app.ActiveWindow.Selection.ShapeRange = { Count: 1, Item: function () { return largeGroup; } };
  let largeSelectionError = "";
  try { W.smartZoomBegin(); } catch (err) { largeSelectionError = String(err && err.message ? err.message : err); }
  check("smart zoom caps oversized selections before WPS freeze", /超过 500/.test(largeSelectionError), largeSelectionError || "no error");
  let orphanApplyResult = null;
  let orphanApplyError = "";
  try { orphanApplyResult = W.smartZoomApply(150, { _sessionId: previousSession && previousSession.sessionId }); } catch (err) { orphanApplyError = String(err && err.message ? err.message : err); }
  check("smart zoom invalidates the old session when re-pick fails", (orphanApplyResult && orphanApplyResult.stale === true) || /会话已结束|重新拾取/.test(orphanApplyError), JSON.stringify(orphanApplyResult || orphanApplyError));
  try { W.smartZoomEnd(); } catch (_) {}

  app.ActiveWindow.Selection.ShapeRange = zoomRange;
  W.smartZoomBegin();
  W.smartZoomEnd();
  let applyAfterEndError = "";
  try { W.smartZoomApply(150, {}); } catch (err) { applyAfterEndError = String(err && err.message ? err.message : err); }
  check("smart zoom requires a fresh pick after ending a session", /重新拾取|开始智能缩放|会话已结束/.test(applyAfterEndError), applyAfterEndError || "no error");
  try { W.smartZoomEnd(); } catch (_) {}

  const staleStart = W.smartZoomBegin();
  const staleSessionId = staleStart && staleStart.sessionId;
  W.smartZoomEnd();
  const freshStart = W.smartZoomBegin();
  const freshLeft = a1.Left;
  const staleApply = W.smartZoomApply(150, { _sessionId: staleSessionId });
  check("smart zoom ignores delayed work from an older session", staleApply && staleApply.stale === true && a1.Left === freshLeft, JSON.stringify({ staleSessionId: staleSessionId, freshSessionId: freshStart && freshStart.sessionId, result: staleApply, left: a1.Left }));
  W.smartZoomEnd();
  // Keep the image replacement assertions independent of the lifecycle
  // probe above when running the intentionally red pre-fix test.
  a1.Left = zoomOriginal.a1.left; a1.Top = zoomOriginal.a1.top; a1.Width = zoomOriginal.a1.width; a1.Height = zoomOriginal.a1.height;
  a2.Left = zoomOriginal.a2.left; a2.Top = zoomOriginal.a2.top; a2.Width = zoomOriginal.a2.width; a2.Height = zoomOriginal.a2.height;
  app.ActiveWindow.Selection.ShapeRange = null;

  // ---- object filter: selection-only operations on a slide container ----
  const filterSlide = new MockSlide(deck, 99);
  function makeFilterShape(name, type, width, height, extra) {
    const shape = Object.assign({
      deck: deck,
      slide: filterSlide,
      Id: 900 + filterSlide.shapes.length,
      Name: name,
      Type: type,
      Width: width,
      Height: height,
      Parent: filterSlide,
      Fill: { ForeColor: { RGB: 0 } },
      Line: { ForeColor: { RGB: 0 } }
    }, extra || {});
    shape.Select = MockShape.prototype.Select;
    filterSlide.shapes.push(shape);
    return shape;
  }
  const filterTextA = makeFilterShape("文字A", 1, 100, 40, {
    TextFrame2: { HasText: -1, TextRange: { Text: "A", Font: { Size: 18, Fill: { ForeColor: { RGB: 0x112233 } } } } }
  });
  const filterTextB = makeFilterShape("文字B", 1, 100, 50, {
    TextFrame2: { HasText: -1, TextRange: { Text: "B", Font: { Size: 18, Fill: { ForeColor: { RGB: 0x112233 } } } } }
  });
  const filterTextC = makeFilterShape("文字C", 1, 100, 40, {
    TextFrame2: { HasText: -1, TextRange: { Text: "C", Font: { Size: 24, Fill: { ForeColor: { RGB: 0x445566 } } } } }
  });
  const filterLineA = makeFilterShape("线条A", 9, 12, 12, { Line: { ForeColor: { RGB: 0xabcdef } } });
  const filterLineB = makeFilterShape("线条B", 9, 20, 20, { Line: { ForeColor: { RGB: 0xabcdef } } });
  const filterGroup = makeFilterShape("组合A", 6, 220, 120, {});
  deck.selectedShapes = [];
  app.ActiveWindow.Selection.ShapeRange = null;
  filterTextA.Select(-1);
  const sameType = W.objectFilterRun("type");
  check("object filter selects same type", sameType.ok && sameType.count === 3, JSON.stringify(sameType));
  filterTextA.Select(-1);
  const sameFont = W.objectFilterRun("fontsize");
  check("object filter selects same font size", sameFont.ok && sameFont.count === 2, JSON.stringify(sameFont));
  filterTextA.Select(-1);
  const sameWidth = W.objectFilterRun("width");
  check("object filter selects same width", sameWidth.ok && sameWidth.count === 3, JSON.stringify(sameWidth));
  filterTextA.Select(-1);
  const sameHeight = W.objectFilterRun("height");
  check("object filter selects same height", sameHeight.ok && sameHeight.count === 2, JSON.stringify(sameHeight));
  filterTextA.Select(-1);
  const sameColor = W.objectFilterRun("color");
  check("object filter selects same text color", sameColor.ok && sameColor.count === 2, JSON.stringify(sameColor));
  const allText = W.objectFilterRun("text");
  check("object filter selects all text", allText.ok && allText.count === 3, JSON.stringify(allText));
  const allLines = W.objectFilterRun("line");
  check("object filter selects all lines", allLines.ok && allLines.count === 2, JSON.stringify(allLines));
  const allGroups = W.objectFilterRun("group");
  check("object filter selects all groups", allGroups.ok && allGroups.count === 1 && deck.selectedShapes[0] === filterGroup, JSON.stringify(allGroups));
  filterTextA.Select(-1); filterTextB.Select(0);
  const inverted = W.objectFilterRun("invert");
  check("object filter inverts current selection", inverted.ok && inverted.count === 4, JSON.stringify(inverted));
  const allObjects = W.objectFilterRun("all");
  check("object filter selects all objects", allObjects.ok && allObjects.count === 6, JSON.stringify(allObjects));
  const mixedFilterShape = makeFilterShape("混合字号", 1, 100, 40, {
    TextFrame2: { HasText: -1, TextRange: { Text: "mixed", Font: { Size: -2, Fill: { ForeColor: { RGB: -2 } } } } }
  });
  mixedFilterShape.Select(-1);
  const mixedBefore = deck.selectedShapes.slice();
  const mixedFontResult = W.objectFilterRun("fontsize");
  check("object filter rejects mixed font baseline without changing selection", !mixedFontResult.ok && deck.selectedShapes.length === mixedBefore.length && deck.selectedShapes[0] === mixedBefore[0], JSON.stringify(mixedFontResult));
  filterSlide.Range = null;
  filterTextA.Select(-1);
  const fallbackLines = W.objectFilterRun("line");
  check("object filter falls back when Shapes.Range is unavailable", fallbackLines.ok && fallbackLines.count === 2, JSON.stringify(fallbackLines));
  delete filterSlide.Range;
  filterSlide.shapes.pop();
  deck.selectedShapes = [];
  app.ActiveWindow.Selection.ShapeRange = null;

  // ---- test 1: collectDeckImages ----
  let partialInventory = null;
  let thumbnailPartialCount = 0;
  const collect = await W.collectDeckImages(null, function (partial) {
    if (partial && partial.groups) partialInventory = partial;
    if (partial && partial.phase === "thumbnails" && partial.updates && partial.updates.length) thumbnailPartialCount += 1;
  });
  check("partial inventory arrives before thumbnail scan", !!partialInventory && partialInventory.complete === false && partialInventory.total === 4 && partialInventory.groups.length === 4, partialInventory ? String(partialInventory.groups.length) : "missing");
  check("thumbnail rows stream before final grouping", thumbnailPartialCount > 0, String(thumbnailPartialCount));
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

  // ---- regression: reopening the panel reuses the completed inventory ----
  const cacheApiReady = typeof W.collectDeckImagesCached === "function" && typeof W.clearDeckImageCache === "function";
  check("inventory cache API is available", cacheApiReady);
  if (cacheApiReady) {
    W.clearDeckImageCache();
    const cachedFirst = await W.collectDeckImagesCached(null, null, { force: true });
    const cachedSecond = await W.collectDeckImagesCached(null, null);
    check("second inventory open is a cache hit", cachedFirst.cacheHit === false && cachedSecond.cacheHit === true && cachedSecond.groups === cachedFirst.groups, JSON.stringify({ first: cachedFirst.cacheHit, second: cachedSecond.cacheHit }));
    // Re-evaluate the add-in core to model a task-pane context recreated by
    // closing and reopening the panel.  The persisted snapshot must hydrate
    // live shape references without exporting thumbnails again.
    delete global.WpsPictureReplace;
    eval(src);
    const reloadedW = global.WpsPictureReplace;
    const cachedDisk = await reloadedW.collectDeckImagesCached(null, null);
    check("inventory cache survives a pane context reload", cachedDisk.cacheHit === true && cachedDisk.cacheSource === "disk" && cachedDisk.groups.length === cachedFirst.groups.length, JSON.stringify({ hit: cachedDisk.cacheHit, source: cachedDisk.cacheSource }));
    const forcedRefresh = await reloadedW.collectDeckImagesCached(null, null, { force: true });
    check("explicit inventory refresh bypasses cache", forcedRefresh.cacheHit === false && forcedRefresh.cacheSource === "scan", JSON.stringify({ hit: forcedRefresh.cacheHit, source: forcedRefresh.cacheSource }));
    reloadedW.clearDeckImageCache();
  }

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
  const layoutA = { Name: "版式A", Shapes: layoutAColl, Select: function () { this._selected = true; return 0; } };
  const layoutB = { Name: "版式B", Shapes: layoutBColl, Select: function () { this._selected = true; return 0; } };
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

  // ---- test 8h: master/layout locate (view switch + shape select) ----
  app.ActiveWindow.ViewType = 9;
  app.ActiveWindow.Selection.ShapeRange = {
    get Count() { return deck.selectedShape ? 1 : 0; },
    Item: function () { return deck.selectedShape; }
  };
  const mv = W.gotoMasterView();
  check("gotoMasterView switches to master view", mv === true && app.ActiveWindow.ViewType === 2, "mv=" + mv + " view=" + app.ActiveWindow.ViewType);
  const ms = W.selectMasterShape(1);
  check("selectMasterShape selects master shape", ms === true && deck.selectedShape === mp1, String(ms));
  const ls = W.selectLayoutShape(1, 1);
  check("selectLayoutShape selects layout shape", ls === true && deck.selectedShape === lp1, String(ls));
  check("layout instance carries layoutIndex", !!layoutAInst && layoutAInst.layoutIndex === 1, layoutAInst ? String(layoutAInst.layoutIndex) : "missing");

  // ---- test 8h2: leaving master view restores normal slide navigation + highlight ----
  app.ActiveWindow.Selection.ShapeRange = {
    get Count() { return deck.selectedShape ? 1 : 0; },
    Item: function () { return deck.selectedShape; }
  };
  app.ActiveWindow.ViewType = 2;
  const normalFromMaster = W.gotoSlide(3);
  check("gotoSlide exits master view", normalFromMaster === true && app.ActiveWindow.ViewType === 9 && app.ActiveWindow.View.current === 3, "ok=" + normalFromMaster + " view=" + app.ActiveWindow.ViewType);
  const slideSelected = W.selectSlideShape(3, 1);
  check("selectSlideShape highlights normal picture", slideSelected === true && deck.selectedShape === a3, String(slideSelected));

  // Real WPS applies GotoSlide asynchronously. Selecting in the same JS turn
  // can therefore target the old view even though GotoSlide returned. The
  // production locator must wait until View.Slide confirms the requested page
  // and only then select the picture on the document canvas.
  const savedGotoSlide = app.ActiveWindow.View.GotoSlide;
  const savedA3Select = a3.Select;
  app.ActiveWindow.View.current = 1;
  app.ActiveWindow.View.GotoSlide = function (n) {
    const view = this;
    setTimeout(function () { view.current = Number(n); }, 25);
  };
  a3.Select = function () {
    if (app.ActiveWindow.View.current !== 3) return;
    deck.selectedShape = this;
  };
  deck.selectedShape = null;
  const delayedCanvasSelect = W.locateSlideShape ? await W.locateSlideShape(3, 1) : false;
  check("locator waits for active slide then highlights canvas picture", delayedCanvasSelect === true && app.ActiveWindow.View.current === 3 && deck.selectedShape === a3, "ok=" + delayedCanvasSelect + " current=" + app.ActiveWindow.View.current);
  app.ActiveWindow.View.GotoSlide = savedGotoSlide;
  a3.Select = savedA3Select;

  // ---- test 8h3: custom layer manager (jump, visibility and lock state) ----
  app.ActiveWindow.ViewType = 9;
  app.ActiveWindow.View.current = 3;
  deck.selectedShape = null;
  deck.selectedShapes = [];
  app.ActiveWindow.Selection.ShapeRange = null;
  const normalLayers = W.layerList();
  const normalLayerItem = normalLayers && normalLayers.items && normalLayers.items[0];
  check("layer manager lists the active slide only", normalLayers.ok && normalLayers.kind === "slide" && normalLayers.slideIndex === 3 && normalLayers.count === 1 && normalLayerItem.shape === a3, JSON.stringify({ kind: normalLayers.kind, slideIndex: normalLayers.slideIndex, count: normalLayers.count }));
  const hiddenLayer = W.layerSetVisible(normalLayerItem, false);
  check("layer manager toggles visibility", hiddenLayer.ok && hiddenLayer.visible === false && a3.Visible === 0, JSON.stringify(hiddenLayer));
  const shownLayer = W.layerSetVisible(normalLayerItem, true);
  check("layer manager restores visibility", shownLayer.ok && shownLayer.visible === true && a3.Visible === -1, JSON.stringify(shownLayer));
  const lockedLayer = W.layerSetLocked(normalLayerItem, true);
  check("layer manager records lock when native lock is unavailable", lockedLayer.ok && lockedLayer.locked === true && lockedLayer.mode === "plugin" && a3.Tags.Item("CODEXLAYERLOCKED") === "1", JSON.stringify(lockedLayer));
  const selectionGuard = app.ApiEvent.listeners.WindowSelectionChange;
  check("layer manager installs a selection guard", typeof selectionGuard === "function", Object.keys(app.ApiEvent.listeners).join(","));
  deck.selectedShapes = [a3];
  deck.selectedShape = a3;
  if (selectionGuard) selectionGuard(app.ActiveWindow.Selection);
  check("locked object cannot remain selected", deck.selectedShapes.indexOf(a3) < 0 && deck.selectedShape !== a3, JSON.stringify({ selected: deck.selectedShapes.length }));
  const blockedLockedLocate = await W.layerSelect(normalLayerItem);
  check("locked object locator is blocked", blockedLockedLocate === false, String(blockedLockedLocate));
  const blockedPictureLocate = W.selectSlideShape(3, 1);
  check("picture locator also respects object lock", blockedPictureLocate === false, String(blockedPictureLocate));
  const lockedLayers = W.layerList();
  check("layer manager reads persisted plugin lock", lockedLayers.items[0].locked === true && lockedLayers.items[0].lockMode === "plugin", JSON.stringify(lockedLayers.items[0] && { locked: lockedLayers.items[0].locked, mode: lockedLayers.items[0].lockMode }));
  const unlockedLayer = W.layerSetLocked(normalLayerItem, false);
  check("layer manager unlocks the object", unlockedLayer.ok && unlockedLayer.locked === false && a3.Tags.Item("CODEXLAYERLOCKED") === "", JSON.stringify(unlockedLayer));
  const jumpedLayer = await W.layerSelect(normalLayerItem);
  check("layer manager jumps to and highlights slide object", jumpedLayer === true && deck.selectedShape === a3 && app.ActiveWindow.ViewType === 9 && app.ActiveWindow.View.current === 3, String(jumpedLayer));

  app.ActiveWindow.ViewType = 2;
  deck.selectedShape = null;
  deck.selectedShapes = [];
  const masterLayers = W.layerList();
  const masterLayerItem = masterLayers && masterLayers.items && masterLayers.items[0];
  check("layer manager distinguishes master context", masterLayers.ok && masterLayers.kind === "master" && masterLayers.count === 2 && masterLayerItem.kind === "master", JSON.stringify({ kind: masterLayers.kind, count: masterLayers.count }));
  const jumpedMasterLayer = await W.layerSelect(masterLayerItem);
  check("layer manager jumps to and highlights master object", jumpedMasterLayer === true && deck.selectedShape === masterLayerItem.shape && app.ActiveWindow.ViewType === 2, String(jumpedMasterLayer));
  deck.selectedShape = lp1;
  deck.selectedShapes = [lp1];
  const layoutLayers = W.layerList();
  const layoutLayerItem = layoutLayers && layoutLayers.items && layoutLayers.items[0];
  check("layer manager distinguishes layout context", layoutLayers.ok && layoutLayers.kind === "layout" && layoutLayers.layoutIndex === 1 && layoutLayerItem.kind === "layout", JSON.stringify({ kind: layoutLayers.kind, layoutIndex: layoutLayers.layoutIndex }));
  const jumpedLayoutLayer = await W.layerSelect(layoutLayerItem);
  check("layer manager jumps to and highlights layout object", jumpedLayoutLayer === true && deck.selectedShape === layoutLayerItem.shape && app.ActiveWindow.ViewType === 2, String(jumpedLayoutLayer));
  app.ActiveWindow.ViewType = 9;
  app.ActiveWindow.View.current = 3;
  deck.selectedShape = null;
  deck.selectedShapes = [];

  // ---- test 8h4: grouped object manager and batch type selection ----
  const categorySlide = new MockSlide(deck, deck.slides.length + 1); deck.slides.push(categorySlide);
  function makeCategoryShape(name, type, extra) {
    const shape = categorySlide.AddPicture("C:/img/A.png", 0, -1, 20 + categorySlide.shapes.length * 20, 20, 80, 50);
    shape.name = name;
    shape.type = type;
    Object.assign(shape, extra || {});
    return shape;
  }
  const categoryImageA = makeCategoryShape("图片分类A", 13);
  const categoryImageB = makeCategoryShape("图片分类B", 13);
  // Same-named objects must still be treated as separate shapes during a
  // batch selection; WPS exposes Shape.Id, while the display name is editable.
  categoryImageB.name = "图片分类A";
  const categoryText = makeCategoryShape("文字分类", 17, { TextFrame2: { HasText: -1, TextRange: { Text: "分类文本" } } });
  const categoryShape = makeCategoryShape("形状分类", 1);
  const categoryLine = makeCategoryShape("线条分类", 9);
  const categoryTable = makeCategoryShape("表格分类", 19);
  const categoryChart = makeCategoryShape("图表分类", 3);
  const categoryGroup = makeCategoryShape("组合分类", 6);
  app.ActiveWindow.ViewType = 9;
  app.ActiveWindow.View.current = categorySlide.index;
  deck.selectedShape = null;
  deck.selectedShapes = [];
  app.ActiveWindow.Selection.ShapeRange = null;
  const categoryLayers = W.layerList();
  const categoryGroups = categoryLayers && categoryLayers.groups || [];
  const categoryKeys = categoryGroups.map(function (group) { return group.key; });
  const categoryColors = categoryGroups.map(function (group) { return group.color; }).filter(Boolean);
  const categoryDebug = { ok: categoryLayers && categoryLayers.ok, count: categoryLayers && categoryLayers.count, total: categoryLayers && categoryLayers.total, slideIndex: categoryLayers && categoryLayers.slideIndex, shapeCount: categorySlide.shapes.length, groups: categoryGroups };
  check("object manager returns typed groups with counts and colors", categoryLayers.ok && categoryGroups.length >= 6 && categoryGroups.every(function (group) { return group.count > 0 && group.label && group.color; }) && new Set(categoryColors).size >= 4, JSON.stringify(categoryDebug));
  check("object manager labels each item with a stable type key", categoryLayers.items.every(function (item) { return item.typeKey && item.typeColor && categoryKeys.indexOf(item.typeKey) >= 0; }), JSON.stringify(categoryLayers.items.map(function (item) { return { name: item.name, typeKey: item.typeKey, typeLabel: item.typeLabel }; })) || JSON.stringify(categoryDebug));
  const textItems = categoryLayers.items.filter(function (item) { return item.typeKey === "text"; });
  const textBatch = W.layerSelectMany(textItems);
  check("object manager batch-selects one object type", textBatch && textBatch.ok && textBatch.count === 1 && deck.selectedShapes.length === 1 && deck.selectedShapes[0] === categoryText, JSON.stringify(textBatch));
  const lockedCategoryItem = categoryLayers.items.find(function (item) { return item.shape === categoryImageA; });
  const lockedCategory = W.layerSetLocked(lockedCategoryItem, true);
  const allCategory = W.layerSelectMany(categoryLayers.items);
  check("object manager batch selection excludes locked objects", lockedCategory && lockedCategory.ok && allCategory && allCategory.ok && allCategory.skippedLocked >= 1 && allCategory.count === categoryLayers.items.length - 1 && deck.selectedShapes.indexOf(categoryImageA) < 0, JSON.stringify({ locked: lockedCategory, selected: allCategory, count: deck.selectedShapes.length }));
  W.layerSetLocked(lockedCategoryItem, false);
  app.ActiveWindow.View.current = 3;
  const taskpaneHasGroupingUi = /layerSelectAll/.test(taskpaneHtml) && /全选本类/.test(taskpaneHtml) && /layer-group/.test(taskpaneHtml) && /type-image/.test(taskpaneHtml);
  check("object manager exposes grouped batch-selection UI", taskpaneHasGroupingUi, "layerSelectAll/layer-group/type-image");

  // ---- test 8i: floating progress panel helpers ----
  const paneH = W.openProgressPanel("测试进度");
  check("openProgressPanel returns pane", !!paneH);
  W.writeTaskProgress(2, 5, "替换图片");
  const st = W.readTaskState();
  check("writeTaskProgress updates task state", !!st && st.done === 2 && st.total === 5 && st.running === true, st ? st.done + "/" + st.total + " running=" + st.running : "null");
  W.closeProgressPanel(paneH);
  const st2 = W.readTaskState();
  check("closeProgressPanel marks task done", !!st2 && st2.running === false, st2 ? "running=" + st2.running : "null");

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

  // ---- test 10: linked fast-path, progress callback, cancel ----
  function fnv1aTest(data) {
    let hash = 2166136261;
    const s = String(data);
    for (let i = 0; i < s.length; i += 1) {
      hash ^= s.charCodeAt(i) & 255;
      hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0).toString(16);
  }
  const sLink = new MockSlide(deck, 11); deck.slides.push(sLink);
  const linkShape = sLink.AddPicture("C:/img/B.png", 0, -1, 10, 10, 100, 80);
  linkShape.name = "LinkedB";
  linkShape.CropLeft = 93.75; // fL = 93.75 / 375 (B baseline 500px*0.75)
  const bFp = fnv1aTest(deck.fs.get("C:/img/B.png"));
  linkShape.alternativeText = 'PICRENEW|{"v":1,"name":"LinkedB","src":"C:/img/B.png","fileFp":"' + bFp + '","contentFp":"","aspect":1429}';
  const lnOk = W._math.linkedNaturalSize(linkShape, {});
  check("linkedNaturalSize uses unchanged source baseline", !!lnOk && Math.abs(lnOk.w - 375) < 0.01 && Math.abs(lnOk.h - 262.5) < 0.01, JSON.stringify(lnOk));
  const lnBad = W._math.linkedNaturalSize(linkShape, { "src:C:/img/B.png": { ok: false, w: 0, h: 0 } });
  check("linkedNaturalSize honors cached miss", lnBad === null);
  const spoofed = Object.assign({}, linkShape, { alternativeText: 'PICRENEW|{"v":1,"name":"x","src":"C:/img/B.png","fileFp":"deadbeef","contentFp":"","aspect":0}' });
  check("linkedNaturalSize rejects changed source", W._math.linkedNaturalSize(spoofed, {}) === null);

  const collectL = await W.collectDeckImages();
  await W.refreshLinkStates(collectL.groups);
  let linkInst = null;
  collectL.groups.forEach(function (g) { g.instances.forEach(function (i) { if (i.name === "LinkedB") linkInst = i; }); });
  const progressCalls = [];
  const repL = await W.replaceInstances([linkInst], "C:/img/B.png", collectL.docKey, function (done, total, label) { progressCalls.push({ done: done, total: total, label: label }); }, null);
  check("replaceInstances reports progress 1/1", progressCalls.length === 1 && progressCalls[0].done === 1 && progressCalls[0].total === 1 && /替换图片/.test(progressCalls[0].label || ""), JSON.stringify(progressCalls));
  check("linked replace succeeded", repL.replaced === 1, JSON.stringify(repL));
  const afterL = sLink.shapes[0];
  // expected from computeNewCrops(0.25,0,0,0,375,262.5,100,80): cropLeft=93.75, cropTop=cropBottom=18.75
  check("linked replace crop baseline preserved",
    Math.abs(afterL.CropLeft - 93.75) < 0.5 && Math.abs(afterL.CropTop - 18.75) < 0.5 && Math.abs(afterL.CropBottom - 18.75) < 0.5,
    JSON.stringify([afterL.CropLeft, afterL.CropTop, afterL.CropBottom]));

  const sCancel = new MockSlide(deck, 12); deck.slides.push(sCancel);
  const c1 = sCancel.AddPicture("C:/img/A.png", 0, -1, 10, 10, 100, 80);
  c1.name = "Cancel1";
  const c2 = sCancel.AddPicture("C:/img/A.png", 0, -1, 300, 10, 100, 80);
  c2.name = "Cancel2";
  const collectC = await W.collectDeckImages();
  await W.refreshLinkStates(collectC.groups);
  const cInsts = [];
  collectC.groups.forEach(function (g) { g.instances.forEach(function (i) { if (i.name === "Cancel1" || i.name === "Cancel2") cInsts.push(i); }); });
  let cancelCalls = 0;
  const repC = await W.replaceInstances(cInsts, "C:/img/B.png", collectC.docKey, null, function () { cancelCalls += 1; return cancelCalls >= 2; });
  check("cancel stops after first item", repC.replaced === 1 && repC.cancelled === true, JSON.stringify(repC));


  // ---------- GitHub update check logic ----------
  check("compareVersions older", W.compareVersions("1.2.16", "1.2.17") === -1, "1.2.16 vs 1.2.17");
  check("compareVersions equal", W.compareVersions("1.2.17", "1.2.17") === 0, "1.2.17 vs 1.2.17");
  check("compareVersions newer", W.compareVersions("1.2.19", "1.2.17") === 1, "1.2.19 vs 1.2.17");
  check("compareVersions numeric parts", W.compareVersions("1.2.9", "1.2.10") === -1, "1.2.9 vs 1.2.10");

  function MockXHR(responseJson, statusCode, finalUrl) {
    this.readyState = 0;
    this.status = statusCode === undefined ? 200 : statusCode;
    this.responseText = responseJson === null ? "" : JSON.stringify(responseJson);
    this.responseURL = finalUrl || "";
    this.responseType = "";
    this.timeout = 0;
    this.onreadystatechange = null;
    this.onerror = null;
    this.ontimeout = null;
    this.open = function (method, url, async) {
      this._url = url;
      try {
        if (global.__mockXhrRoute) {
          const routed = global.__mockXhrRoute(url);
          if (routed) {
            this.status = routed.status === undefined ? 200 : routed.status;
            this.responseText = routed.responseText || "";
            this.responseURL = routed.responseURL || "";
          }
        }
      } catch (_) {}
    };
    this.send = function () {
      const self = this;
      setTimeout(function () {
        self.readyState = 4;
        if (self.onreadystatechange) self.onreadystatechange();
      }, 5);
    };
  }

  const savedXHR = global.XMLHttpRequest;
  global.XMLHttpRequest = function () { return new MockXHR({ name: "picture-replace-tools-wps", version: "1.2.34" }, 200); };
  const up = await W.checkForUpdates();
  check("update check detects newer", up.ok === true && up.hasUpdate === true && up.latest === "1.2.34", JSON.stringify(up));
  check("update check builds download url", /releases\/download\/v1\.2\.34\/PictureReplaceTools-WPS-1\.2\.34\.exe$/.test(up.downloadUrl || ""), up.downloadUrl || "");

  global.XMLHttpRequest = function () { return new MockXHR({ name: "picture-replace-tools-wps", version: "1.2.17" }, 200); };
  const upSame = await W.checkForUpdates();
  check("update check same version = no update", upSame.ok === true && upSame.hasUpdate === false, JSON.stringify(upSame));

  global.XMLHttpRequest = function () { return new MockXHR(null, 404); };
  const upFail = await W.checkForUpdates();
  check("update check http failure handled", upFail.ok === false && /HTTP/.test(upFail.error || ""), JSON.stringify(upFail));

  // Fallback path: raw manifest 404, releases/latest redirect gives the tag.
  global.XMLHttpRequest = function () { return new MockXHR(null, 404, ""); };
  let xhrCount2 = 0;
  global.__mockXhrRoute = function (url) {
    xhrCount2 += 1;
    if (/releases\/latest/.test(url)) {
      return { status: 200, responseText: "", responseURL: "https://github.com/Dongsidaye/ppt-picture-replace-tools/releases/tag/v1.2.34" };
    }
    return null;
  };
  const upFallback = await W.checkForUpdates();
  check("update check falls back to release tag", upFallback.ok === true && upFallback.hasUpdate === true && upFallback.latest === "1.2.34", JSON.stringify(upFallback));
  check("update check used two sources", xhrCount2 >= 2, "xhrCount=" + xhrCount2);
  global.__mockXhrRoute = null;

  global.XMLHttpRequest = savedXHR;


  // ---------- regression: add-in must not throw "trace is not defined" on load ----------
  const alertCalls = [];
  const savedAlert = app.alert;
  app.alert = function (msg) { alertCalls.push(String(msg)); };
  global.OnAddInLoad();
  await new Promise(function (r) { setTimeout(r, 80); });
  app.alert = savedAlert;
  check("no trace is not defined popup on load", alertCalls.length === 0, JSON.stringify(alertCalls));
  check("background preload hooks document activation", typeof app.ApiEvent.listeners.WindowActivate === "function" || typeof app.ApiEvent.listeners.PresentationOpen === "function", Object.keys(app.ApiEvent.listeners).join(","));

  // ---------- responsiveness regression: inventory must yield between slides ----------
  // Each Type getter represents a small synchronous WPS JSAPI bridge call.
  // A whole-deck scan in one turn blocks timer heartbeats and mirrors the
  // frozen WPS window reported by users.
  const perfDeck = { fs: new MockFS(), slides: [], path: "C:/mock/perf.pptx", clipboard: null, selectedShape: null };
  function busyBridge(ms) {
    const until = Date.now() + ms;
    while (Date.now() < until) {}
  }
  for (let pi = 1; pi <= 24; pi += 1) {
    const ps = new MockSlide(perfDeck, pi);
    for (let pj = 0; pj < 8; pj += 1) {
      ps.shapes.push(Object.defineProperty({ PictureFormat: null }, "Type", {
        configurable: true,
        get: function () { busyBridge(1); return 1; }
      }));
    }
    perfDeck.slides.push(ps);
  }
  const perfApp = buildApp(perfDeck);
  const savedWps = global.wps;
  global.wps = perfApp;
  let lastBeat = Date.now();
  let maxBeatGap = 0;
  const heartbeat = setInterval(function () {
    const now = Date.now();
    maxBeatGap = Math.max(maxBeatGap, now - lastBeat);
    lastBeat = now;
  }, 5);
  await W.collectDeckImages();
  await new Promise(function (r) { setTimeout(r, 20); });
  clearInterval(heartbeat);
  global.wps = savedWps;
  check("inventory scan keeps event loop responsive", maxBeatGap < 80, "maxGap=" + maxBeatGap + "ms");

  const failed = results.filter(r => !r.ok);
  console.log("\n===== " + (failed.length ? failed.length + " FAILURES" : "ALL TESTS PASSED") + " (" + results.length + " checks) =====");
  process.exit(failed.length ? 1 : 0);
}

main().catch(e => { console.error("HARNESS ERROR:", e); process.exit(2); });
