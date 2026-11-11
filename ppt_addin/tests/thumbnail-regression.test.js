const fs = require('fs');
const vm = require('vm');
const assert = require('node:assert/strict');
const path = require('path');

const main = fs.readFileSync(path.join(__dirname, '../shared/main.js'), 'utf8');

function fixture(options = {}) {
  let clipboard = { color: 'old-clipboard', mode: 'ok' };
  let pasted = options.uncleared ? [{ color: 'old-scratch' }] : [];
  let exported = [];
  const calls = { copies: [], pastes: [], deleted: [], exports: 0, decodes: 0, clears: 0 };
  const shapes = {
    get Count() { return pasted.length; },
    Item(index) { return pasted[index - 1]; },
    Paste() {
      const source = clipboard;
      calls.pastes.push(source.color);
      let width = 1000;
      const shape = {
        color: source.color,
        PictureFormat: {},
        Left: 0, Top: 0, Height: 1000,
        get Width() { return width; },
        set Width(value) {
          if (source.mode === 'partial' || source.mode === 'delete-fails') {
            throw new Error('Paste inserted a shape, but positioning failed');
          }
          width = value;
        },
        Delete() {
          if (source.mode === 'delete-fails') throw new Error('Scratch Delete failed');
          const index = pasted.indexOf(shape);
          assert.notEqual(index, -1, 'Delete must target an existing scratch shape');
          calls.deleted.push(source.color);
          pasted.splice(index, 1);
        }
      };
      pasted.push(shape);
      return shape;
    }
  };
  const slide = {
    Shapes: shapes,
    Export() {
      calls.exports += 1;
      exported = pasted.map(shape => ({
        color: shape.color, left: shape.Left, top: shape.Top,
        width: shape.Width, height: shape.Height
      }));
    }
  };
  const context = { console, setTimeout, clearTimeout, window: null, slide };
  context.window = context;
  context.decode = count => {
    calls.decodes += 1;
    // Model the exported grid, including overlap. Collection order cannot
    // identify a cell once an earlier source fails to paste.
    return Array.from({ length: count }, (_, index) => {
      const x = ((index % context.columns) + 0.5) * context.cell;
      const y = (Math.floor(index / context.columns) + 0.5) * context.cell;
      let color = 'blank';
      for (const shape of exported) {
        if (x >= shape.left && x < shape.left + shape.width &&
            y >= shape.top && y < shape.top + shape.height) color = shape.color;
      }
      return { fp: color, dataUrl: color };
    });
  };
  vm.createContext(context);
  const hooks = `
    ensureBatchSlide = () => global.slide;
    asShape = shape => shape;
    tempPath = () => "thumbnail.png";
    removeFile = () => {};
    fileExists = () => true;
    fileSystem = () => ({ readAsBinaryString: () => "fixture" });
    yieldUI = async () => {};
    cropWindowForShape = () => ({ l: 0, t: 0, r: 1, b: 1 });
    decodeBatchCells = async (binary, count) => global.decode(count);
    global.render = renderThumbnailsBatch;
    global.cell = BATCH_CELL;
    global.columns = BATCH_COLS;
  `;
  const marker = '  global.WpsPictureReplace = {';
  assert.ok(main.includes(marker), 'Engine test injection marker must exist');
  vm.runInContext(main.replace(marker, hooks + '\n' + marker), context);

  return {
    calls,
    async render(sources) {
      const items = sources.map(([color, mode = 'ok']) => ({ shape: {
        Width: 1000, Height: 1000, Rotation: 0, HorizontalFlip: 0, VerticalFlip: 0,
        Copy() {
          calls.copies.push(color);
          if (mode === 'copy-fails') throw new Error('Source Copy failed');
          clipboard = { color, mode };
        }
      } }));
      return context.render(items, {
        clear() {
          calls.clears += 1;
          if (!options.uncleared) pasted = [];
        }
      }, null);
    }
  };
}

function thumbnails(results) {
  return Array.from(results, info => info && info.dataUrl);
}

async function run() {
  const middleFailures = fixture();
  const middleResults = await middleFailures.render([
    ['red'], ['copy-failed', 'copy-fails'], ['partial-failed', 'partial'], ['blue']
  ]);
  assert.deepEqual(thumbnails(middleResults), ['red', null, null, 'blue'],
    'Middle failures must preserve the first and fourth physical grid cells');
  assert.deepEqual(middleFailures.calls.pastes, ['red', 'partial-failed', 'blue'],
    'A failed Copy must never paste the previous clipboard contents');
  assert.deepEqual(middleFailures.calls.deleted, ['partial-failed'],
    'Partial paste cleanup must delete only that failed source');

  const firstFailure = fixture();
  assert.deepEqual(thumbnails(await firstFailure.render([
    ['copy-failed', 'copy-fails'], ['green'], ['yellow'], ['blue']
  ])), [null, 'green', 'yellow', 'blue'],
  'A missing first shape must not shift later thumbnails into earlier rows');

  const cleanupFailure = fixture();
  assert.deepEqual(thumbnails(await cleanupFailure.render([
    ['red'], ['copy-failed', 'copy-fails'], ['partial-failed', 'partial'], ['undeletable', 'delete-fails']
  ])), [null, null, null, null],
  'If partial paste cleanup fails, every thumbnail in that batch must be discarded');
  assert.equal(cleanupFailure.calls.exports, 0, 'A contaminated scratch grid must not be exported');
  assert.equal(cleanupFailure.calls.decodes, 0, 'A contaminated scratch grid must not be decoded');
  assert.ok(!cleanupFailure.calls.deleted.includes('red'), 'Cleanup must preserve previous successful scratch shapes');

  const recovery = fixture();
  assert.deepEqual(thumbnails(await recovery.render([
    ['red'], ['undeletable', 'delete-fails'], ['yellow'], ['blue'],
    ['next-a'], ['next-b'], ['next-c'], ['next-d']
  ])), [null, null, null, null, 'next-a', 'next-b', 'next-c', 'next-d'],
  'A failed batch must not poison the next batch after scratch cleanup succeeds');
  assert.equal(recovery.calls.exports, 1, 'Only the clean second batch should be exported');

  const uncleared = fixture({ uncleared: true });
  assert.deepEqual(thumbnails(await uncleared.render([
    ['red'], ['green'], ['yellow'], ['blue']
  ])), [null, null, null, null],
  'Old scratch content that survives clear must prevent this batch from rendering');
  assert.equal(uncleared.calls.copies.length, 0, 'Do not touch the clipboard when scratch cannot be cleared');
  assert.equal(uncleared.calls.exports, 0, 'Do not export an uncleared scratch slide');

  console.log('PASS thumbnail grid identity, Copy failure, partial cleanup, cleanup failure, and batch recovery');
}

run().catch(error => { console.error(error); process.exitCode = 1; });
