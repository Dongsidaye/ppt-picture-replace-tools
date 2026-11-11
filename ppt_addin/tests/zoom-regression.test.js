"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const root = path.resolve(__dirname, "..");
const main = fs.readFileSync(path.join(root, "shared", "main.js"), "utf8");

assert.match(main, /protectTextReadability: false/, "layout fidelity must be default");
assert.match(main, /SMART_ZOOM_HOST_MIN_TEXT_SIZE_PT = 1/, "1pt host lower bound is required");
assert(!main.includes('key: "lineRuleAfter"'), "paragraph enum must not be scaled as a length");
assert.match(main, /textSpacing/, "character spacing must scale with font size");
assert.match(main, /smartZoomSnapshotTextLayout/, "AutoSize\/WordWrap must be snapshotted");
assert.match(main, /smartZoomPrepareTextLayout\(node\)/, "AutoSize must freeze before writes");
assert.match(main, /smartZoomRestoreTextLayout\(smartZoomSession\.nodes\[i\]\)/, "reset restores text-frame state");
assert.match(main, /smartZoomValidateHostTextFloor\(smartZoomSession\.nodes, factor, smartZoomSession\.options\)/,
  "font-floor preflight must happen before any host write");

// For a fixed text frame, proportional frame, margin, glyph and tracking
// scaling preserves the same capacity and hence the same soft-wrap decision.
function capacity(width, margin, font, tracking) {
  return (width - margin * 2) / (font + tracking);
}
const before = capacity(240, 12, 24, 1);
const factor = 0.25;
assert(Math.abs(capacity(240 * factor, 12 * factor, 24 * factor, factor) - before) < 1e-9,
  "proportional text layout preserves line capacity");
assert.notEqual(capacity(240 * factor, 12 * factor, 8, factor), before,
  "a font-only 8pt clamp changes soft wrapping");

// Fixed line spacing is absolute and must scale; proportional spacing is a
// multiplier and must stay untouched. Indents are always absolute lengths.
assert.match(main, /\["LeftMargin", "RightMargin", "FirstLineIndent"\]/,
  "paragraph indents are captured as absolute lengths");
assert.match(main, /beforeRule\.value === 0/, "SpaceBefore only scales when it is points");
assert.match(main, /afterRule\.value === 0/, "SpaceAfter only scales when it is points");
assert.match(main, /lineRule\.value === 0/, "SpaceWithin only scales when it is points");
assert.match(main, /smartZoomPrepareTextLayout\(entry\)/, "nested table text frames are frozen");
assert.match(main, /smartZoomPrepareTextLayout\);/, "group descendants are frozen recursively");
assert.match(main, /smartZoomSession\.options\.scaleText === false\) return/, "paragraph metrics respect disabled text scaling");

function makeParagraph(ruleBefore, ruleAfter, ruleWithin) {
  return {
    Start: 1, Length: 12,
    ParagraphFormat: {
      LineRuleBefore: ruleBefore, LineRuleAfter: ruleAfter, LineRuleWithin: ruleWithin,
      SpaceBefore: 10, SpaceAfter: 8, SpaceWithin: 6,
      LeftMargin: 4, RightMargin: 3, FirstLineIndent: 2
    }
  };
}
function makeShape(fontSize, rules) {
  const paragraph = makeParagraph(...rules);
  const range = {
    Font: { Size: fontSize, Spacing: 2 },
    Paragraphs(i) { return i === 1 ? paragraph : null; }
  };
  return {
    Type: 17, Left: 10, Top: 20, Width: 200, Height: 80, LockAspectRatio: 0,
    TextFrame: { AutoSize: -1, WordWrap: -1, TextRange: range },
    _paragraph: paragraph
  };
}
function loadEngine(selection) {
  const context = { console, setTimeout, clearTimeout, URL, atob, btoa };
  context.window = context;
  context.wps = { ActiveWindow: { Selection: { ShapeRange: {
    Count: selection.length, Item(i) { return selection[i - 1]; }
  } } } };
  const injected = main.replace("  global.WpsPictureReplace = {", "  global.__zoomInternals = { smartZoomPrepareTextLayout: smartZoomPrepareTextLayout, smartZoomRestoreTextLayout: smartZoomRestoreTextLayout };\n  global.WpsPictureReplace = {");
  vm.runInNewContext(injected, context);
  return context;
}

// Executes the generated engine: point line spacing scales, line multipliers
// do not; child and table frames are frozen/restored recursively.
const childShape = makeShape(10, [0, -1, 0]);
const cellShape = makeShape(10, [-1, 0, -1]);
childShape.Table = { Rows: { Count: 1 }, Columns: { Count: 1 }, Cell() { return { Shape: cellShape }; } };
const group = makeShape(10, [0, 0, 0]);
group.Type = 6;
group.GroupItems = { Count: 1, Item() { return childShape; } };
const ctx = loadEngine([group]);
const W = ctx.WpsPictureReplace;
const info = W.smartZoomBegin();
W.smartZoomApply(50, { _sessionId: info.sessionId });
assert.equal(group.TextFrame.AutoSize, 0);
assert.equal(childShape.TextFrame.AutoSize, 0);
assert.equal(cellShape.TextFrame.AutoSize, 0);
assert.equal(group._paragraph.ParagraphFormat.SpaceBefore, 5);
assert.equal(group._paragraph.ParagraphFormat.SpaceAfter, 4);
assert.equal(group._paragraph.ParagraphFormat.SpaceWithin, 3);
assert.equal(childShape._paragraph.ParagraphFormat.SpaceAfter, 8, "line multiplier must not scale");
assert.equal(cellShape._paragraph.ParagraphFormat.SpaceBefore, 10, "table line multiplier must not scale");
W.smartZoomReset({ _sessionId: info.sessionId });
assert.equal(group.TextFrame.AutoSize, -1);
assert.equal(childShape.TextFrame.AutoSize, -1);
assert.equal(cellShape.TextFrame.AutoSize, -1);

// The 1pt validation happens before AutoSize, geometry, font or paragraph
// setters, so a rejected request cannot leave a half-modified selection.
const tiny = makeShape(10, [0, 0, 0]);
let writes = 0;
["Left", "Top", "Width", "Height"].forEach((key) => {
  let value = tiny[key];
  Object.defineProperty(tiny, key, { get: () => value, set: v => { writes++; value = v; }, configurable: true });
});
const tinyCtx = loadEngine([tiny]);
const tinyInfo = tinyCtx.WpsPictureReplace.smartZoomBegin();
assert.throws(() => tinyCtx.WpsPictureReplace.smartZoomApply(5, { _sessionId: tinyInfo.sessionId }), /1pt/);
assert.equal(writes, 0, "failed preflight must not write geometry");

console.log("PASS zoom typography regression contract");
