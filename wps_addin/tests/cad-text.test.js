"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const source = fs.readFileSync(path.join(__dirname, "../cad-text.js"), "utf8");

function fixture(options = {}) {
  const shapes = [];
  const deleted = [];
  const legacyWrites = [];
  const window = {};
  vm.runInNewContext(source, { window, isFinite, Error });

  function create(kind) {
    const font = {
      Name: "Arial", NameFarEast: "Arial", Size: 20, Spacing: 0,
      Fill: { Visible: -1, Solid() {}, ForeColor: { RGB: 0 } },
      Line: { Visible: -1, Weight: 1, Transparency: 0, ForeColor: { RGB: 123 } }
    };
    let size = 20;
    Object.defineProperty(font, "Size", {
      get() { return size; },
      set(value) { size = options.badSize ? value + 1 : options.roundSize ? Math.round(value * 10) / 10 : value; }
    });
    const shape = {
      Left: 0, Top: 0, Width: 120, Height: 50, Rotation: 0, LockAspectRatio: 0,
      kind, TextEffect: { PresetShape: 1, NormalizedHeight: 0 },
      Line: { Visible: -1, Weight: 1, Transparency: 0, ForeColor: { RGB: 123 } },
      Delete() {
        if (options.cleanupFailure) throw new Error("cleanup failed");
        deleted.push(shape);
      }
    };
    let width = shape.Width;
    Object.defineProperty(shape, "Width", {
      get() { return width; },
      set(value) {
        width = value;
        // Some hosts reapply preset outlines when their geometry changes.
        font.Line.Transparency = 0;
        font.Line.Visible = -1;
      }
    });
    const range = { Text: "", Font: font, ParagraphFormat: {} };
    Object.defineProperties(range, {
      BoundWidth: { get() { return 100 * size / 20 + font.Spacing * 9; } },
      BoundHeight: { get() { return size * 1.2; } },
      BoundLeft: { get() { return shape.Left; } },
      BoundTop: { get() { return shape.Top; } }
    });
    range.Characters = () => range;
    const frame = { TextRange: range };
    if (options.legacyFrameOnly) {
      let wrap = -1;
      let bottom = 3.6;
      Object.defineProperty(frame, "WordWrap", { get() { return wrap; }, set() {} });
      Object.defineProperty(frame, "MarginBottom", { get() { return bottom; }, set() {} });
      shape.TextFrame = {};
      Object.defineProperty(shape.TextFrame, "WordWrap", {
        get() { return wrap; },
        set(value) { wrap = value; legacyWrites.push("WordWrap"); }
      });
      Object.defineProperty(shape.TextFrame, "MarginBottom", {
        get() { return bottom; },
        set(value) { bottom = value; legacyWrites.push("MarginBottom"); }
      });
    }
    let warp = 0;
    Object.defineProperty(frame, "WarpFormat", {
      get() { return warp; },
      set(value) { if (!options.ignoreWarp) warp = value; }
    });
    if (options.unsupportedCosmetics) {
      Object.defineProperty(shape, "ThreeD", { get() { throw new Error("not implemented"); } });
      Object.defineProperty(font, "Shadow", { get() { throw new Error("not implemented"); } });
    }
    shape.TextFrame2 = frame;
    shapes.push(shape);
    return shape;
  }
  const slide = {
    Shapes: {
      AddTextbox() { return create("textbox"); },
      AddTextEffect() { return create("wordart"); }
    }
  };
  return { api: window.WpsCadText, slide, shapes, deleted, legacyWrites };
}

function plan(overrides = {}) {
  return {
    item: { Text: "中文 DWG 123", Rotation: 0, Color: "#336699", Bold: false, Italic: false },
    fontSize: 20,
    metrics: { x: 2, y: 4, width: 96, height: 18, lineHeight: 25 },
    targetWidth: 97,
    targetHeight: 18,
    centerX: 100,
    centerY: 200,
    ...overrides
  };
}

function close(actual, expected, tolerance = 0.001) {
  assert(Math.abs(actual - expected) <= tolerance, `${actual} differs from ${expected}`);
}

function testOrdinaryTextKeepsLayoutAndContent() {
  const f = fixture();
  const item = plan();
  const shape = f.api.add(f.slide, item, "Microsoft YaHei");
  assert.equal(shape.kind, "textbox");
  assert.equal(shape.TextFrame2.TextRange.Text, item.item.Text);
  assert.equal(shape.TextFrame2.WordWrap, 0);
  assert.equal(shape.TextFrame2.MarginLeft, 0);
  close(shape.TextFrame2.TextRange.Font.Spacing, 1 / 9);
  close(shape.Left, 49.5);
  close(shape.Top, 187.5);
  assert.equal(shape.TextFrame2.TextRange.Font.Fill.ForeColor.RGB, 0x996633);
  assert.equal(shape.TextFrame2.TextRange.Font.Line.Transparency, 1);
  assert.equal(shape.TextFrame2.TextRange.Font.Line.Visible, 0);
  assert.equal(f.deleted.length, 0);
}

function testWordArtForcesActualWpsWarp() {
  const f = fixture();
  const item = plan({ targetWidth: 76.8 });
  item.item.Rotation = -45;
  item.item.Bold = true;
  const shape = f.api.add(f.slide, item, "Microsoft YaHei");
  assert.equal(shape.kind, "wordart");
  assert.equal(shape.TextEffect.PresetShape, 1);
  assert.equal(shape.TextFrame2.WarpFormat, 36);
  assert.equal(shape.TextFrame2.AutoSize, 0);
  assert.equal(shape.TextFrame2.TextRange.Font.Spacing, 0);
  close(shape.Width, 76.8);
  close(shape.Height, 18);
  close(shape.Left + shape.Width / 2, 100);
  close(shape.Top + shape.Height / 2, 200);
  assert.equal(shape.Rotation, 315);
  assert.equal(shape.TextFrame2.TextRange.Font.Bold, -1);
  assert.equal(shape.TextFrame2.TextRange.Font.Line.Transparency, 1);
}

function testOrdinaryRotationKeepsThePaintedCenter() {
  const f = fixture();
  const item = plan();
  item.item.Rotation = 90;
  const shape = f.api.add(f.slide, item, "Arial");
  assert.equal(shape.kind, "textbox");
  assert.equal(shape.Rotation, 90);
  close(shape.Left, 49.9);
  close(shape.Top, 188);
}

function testIgnoredWarpNeverReturnsWrongGlyphs() {
  const f = fixture({ ignoreWarp: true });
  assert.throws(() => f.api.add(f.slide, plan({ targetWidth: 76.8 }), "Arial"), /横向变形/);
  assert.equal(f.shapes.length, 1);
  assert.equal(f.deleted.length, 1);
}

function testUnsupportedAndTinyPlansCreateNothing() {
  const f = fixture();
  assert.throws(() => f.api.add(f.slide, plan({ unsupported: "unsupported source" }), "Arial"), /unsupported source/);
  assert.throws(() => f.api.add(f.slide, plan({ fontSize: 0.9 }), "Arial"), /字号/);
  const blank = plan();
  blank.item.Text = "  \t";
  assert.equal(f.api.add(f.slide, blank, "Arial"), null);
  assert.equal(f.shapes.length, 0);
}

function testRoundedFontMetricsStayUsable() {
  const f = fixture({ roundSize: true });
  const item = plan({ fontSize: 18.89, targetWidth: 96, targetHeight: 18 });
  const shape = f.api.add(f.slide, item, "Microsoft YaHei");
  assert.equal(shape.TextFrame2.TextRange.Font.Size, 18.9);
  const bad = fixture({ badSize: true });
  assert.throws(() => bad.api.add(bad.slide, plan(), "Arial"), /字号/);
  assert.equal(bad.deleted.length, 1);
}

function testFailedCleanupExposesOnlyItsCreatedShape() {
  const f = fixture({ ignoreWarp: true, cleanupFailure: true });
  let error;
  try { f.api.add(f.slide, plan({ targetWidth: 70 }), "Arial"); } catch (caught) { error = caught; }
  assert(error);
  assert.equal(error.createdShape, f.shapes[0]);
  assert.match(error.cleanupError, /cleanup failed/);
  assert.equal(f.shapes.length, 1);
}

function testHeightMismatchFallsBack() {
  const f = fixture();
  assert.throws(() => f.api.add(f.slide, plan({ targetHeight: 40 }), "Arial"), /高度差异/);
  assert.equal(f.deleted.length, 1);
}

function testOptionalCosmeticsDoNotBlockEditableText() {
  const f = fixture({ unsupportedCosmetics: true });
  const shape = f.api.add(f.slide, plan({ targetWidth: 70 }), "Arial");
  assert.equal(shape.TextFrame2.TextRange.Text, "中文 DWG 123");
  assert.equal(shape.TextFrame2.WarpFormat, 36);
}

function testLegacyFrameSettersAreUsedAndVerified() {
  const f = fixture({ legacyFrameOnly: true });
  const shape = f.api.add(f.slide, plan(), "Arial");
  assert.equal(shape.TextFrame2.WordWrap, 0);
  assert.equal(shape.TextFrame2.MarginBottom, 0);
  assert.deepEqual(f.legacyWrites, ["WordWrap", "MarginBottom"]);
  assert.equal(f.deleted.length, 0);
}

for (const test of [
  testOrdinaryTextKeepsLayoutAndContent,
  testWordArtForcesActualWpsWarp,
  testOrdinaryRotationKeepsThePaintedCenter,
  testIgnoredWarpNeverReturnsWrongGlyphs,
  testUnsupportedAndTinyPlansCreateNothing,
  testRoundedFontMetricsStayUsable,
  testFailedCleanupExposesOnlyItsCreatedShape,
  testHeightMismatchFallsBack,
  testOptionalCosmeticsDoNotBlockEditableText,
  testLegacyFrameSettersAreUsedAndVerified
]) {
  test();
  console.log("PASS " + test.name);
}
