/* Runs only from the add-in's explicit RunSelfTest automation method. */
(async function () {
  "use strict";
  const host = window.__pptHost;
  const W = window.WpsPictureReplace;
  const app = window.wps;
  const tests = [];
  const owned = [];
  const files = [];
  const originalDocuments = [];
  let originalWindow = null;
  let deck = null;
  let slide = null;
  let one = null;
  let two = null;
  let picture = null;
  let beforeCount = -1;
  window.__pptTesting = true;
  function assert(ok, message) { if (!ok) throw new Error(message); }
  function near(a, b, epsilon) { return Math.abs(Number(a) - Number(b)) < (epsilon || 0.2); }
  async function test(name, work) {
    const start = performance.now();
    try { const detail = await work(); tests.push({ name, ok: true, ms: Math.round(performance.now() - start), detail }); }
    catch (error) { tests.push({ name, ok: false, ms: Math.round(performance.now() - start), error: String(error.stack || error) }); }
    host.Log("selftest " + name + ": " + tests[tests.length - 1].ok);
  }
  function activate() {
    deck.Windows.Item(1).Activate();
    app.ActiveWindow.ViewType = 9;
    app.ActiveWindow.View.GotoSlide(1);
  }
  function select(shapes) { activate(); shapes.forEach((shape, i) => shape.Select(i ? 0 : -1)); }
  function png(name, color) {
    const canvas = document.createElement("canvas"); canvas.width = 240; canvas.height = 120;
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = color; ctx.fillRect(0, 0, 240, 120);
    ctx.fillStyle = "#ffffff"; ctx.fillRect(10, 10, 70, 40);
    const path = String(host.TempFolder) + "\\PPTSelfTest_" + Date.now() + "_" + name + ".png";
    host.WriteBinaryBase64(path, canvas.toDataURL("image/png").split(",")[1]); files.push(path); return path;
  }
  async function renderPixels() {
    const path = String(host.TempFolder) + "\\PPTSelfTest_render_" + Date.now() + ".png"; files.push(path);
    slide.Export(path, "PNG", 640, 360);
    const img = new Image();
    await new Promise((resolve, reject) => { img.onload = resolve; img.onerror = reject; img.src = "data:image/png;base64," + host.ReadBinaryBase64(path); });
    const canvas = document.createElement("canvas"); canvas.width = 640; canvas.height = 360;
    const ctx = canvas.getContext("2d"); ctx.drawImage(img, 0, 0, 640, 360);
    return ctx.getImageData(0, 0, 640, 360).data;
  }
  try {
    await test("binary-bridge-preserves-null-and-high-bytes", function () {
      const bytes = Array.from({ length: 256 }, (_, i) => String.fromCharCode(i)).join("");
      const path = String(host.TempFolder) + "\\PPTSelfTest_binary.dat"; files.push(path);
      app.FileSystem.writeAsBinaryString(path, bytes);
      const result = app.FileSystem.readAsBinaryString(path);
      assert(result.length === 256 && result === bytes, "binary transfer truncated or changed bytes");
      return { bytes: result.length };
    });
    await test("bridge-and-owned-document", function () {
      beforeCount = Number(app.Presentations.Count);
      for (let i = 1; i <= beforeCount; i++) {
        const doc = app.Presentations.Item(i);
        if (Number(doc.Windows.Count) > 0) originalDocuments.push({ doc, saved: Number(doc.Saved), slides: Number(doc.Slides.Count), name: String(doc.Name) });
      }
      try { originalWindow = app.ActiveWindow; } catch (_) {}
      const version = String(app.Version);
      deck = app.Presentations.Add(-1); owned.push(deck);
      slide = deck.Slides.Add(1, 12);
      one = slide.Shapes.AddShape(1, 40, 50, 120, 80); one.Name = "PPTSelfTest_one";
      two = slide.Shapes.AddShape(1, 240, 150, 60, 40); two.Name = "PPTSelfTest_two";
      one.TextFrame.TextRange.Text = "Cat cat scatter cat";
      one.TextFrame.TextRange.Font.Size = 24;
      one.Fill.ForeColor.RGB = 255;
      two.TextFrame.TextRange.Text = "second";
      assert(Number(deck.Slides.Count) === 1 && near(one.Width, 120), "COM read/write mismatch");
      return { version, initialPresentations: beforeCount };
    });
    if (!deck || !one) throw new Error("基础 COM 对象桥接失败，后续测试已停止。");
    await test("selection-and-type-filter", function () {
      select([one]); const result = W.objectFilterRun("type");
      assert(Number(app.ActiveWindow.Selection.ShapeRange.Count) === 2, "same-type did not select both shapes");
      return String(result && result.message || "selected");
    });
    await test("smart-zoom-reset-and-stale-session", function () {
      select([one]); const first = W.smartZoomBegin();
      W.smartZoomApply(25, { _sessionId: first.sessionId, anchor: "top-left", protectTextReadability: false });
      assert(near(one.Width, 30) && near(one.TextFrame.TextRange.Font.Size, 6), "zoom geometry/text ratio");
      W.smartZoomReset({ _sessionId: first.sessionId });
      assert(near(one.Width, 120) && near(one.TextFrame.TextRange.Font.Size, 24), "reset mismatch");
      W.smartZoomEnd(); select([two]); const second = W.smartZoomBegin();
      const stale = W.smartZoomApply(10, { _sessionId: first.sessionId });
      assert(stale.stale && near(two.Width, 60), "stale callback changed new selection");
      W.smartZoomEnd(); return { first: first.sessionId, second: second.sessionId };
    });
    await test("table-small-scale-font-hierarchy", function () {
      const table = slide.Shapes.AddTable(2, 2, 30, 300, 300, 100);
      table.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Title";
      table.Table.Cell(1, 1).Shape.TextFrame.TextRange.Font.Size = 24;
      table.Table.Cell(2, 1).Shape.TextFrame.TextRange.Text = "Body";
      table.Table.Cell(2, 1).Shape.TextFrame.TextRange.Font.Size = 12;
      select([table]); const info = W.smartZoomBegin(); W.smartZoomApply(20, { _sessionId: info.sessionId });
      const title = Number(table.Table.Cell(1, 1).Shape.TextFrame.TextRange.Font.Size);
      const body = Number(table.Table.Cell(2, 1).Shape.TextFrame.TextRange.Font.Size);
      assert(body >= 3.9 && near(title / body, 2, 0.1), "table font hierarchy lost: " + title + "/" + body);
      W.smartZoomReset({ _sessionId: info.sessionId }); W.smartZoomEnd(); table.Delete();
      return { title, body };
    });
    await test("literal-whole-word-case-sensitive-text", function () {
      one.TextFrame.TextRange.Text = "Cat cat scatter cat";
      activate(); const result = W.designTextFindReplace("cat", "dog", { scope: "current", matchCase: true, wholeWord: true });
      assert(String(one.TextFrame.TextRange.Text) === "Cat dog scatter dog", "word/case options ignored: " + JSON.stringify(result) + " actual=" + String(one.TextFrame.TextRange.Text));
      assert(Number(result.occurrences) === 2, "occurrence count wrong"); return { occurrences: result.occurrences };
    });
    await test("style-and-ole-color", function () {
      select([one]); const captured = W.designStyleCapture(); select([two]); const applied = W.designStyleApply({ fill: true, text: true });
      assert(Number(two.Fill.ForeColor.RGB) === 255, "style fill wrong: " + JSON.stringify({ captured, applied, source: Number(one.Fill.ForeColor.RGB), target: Number(two.Fill.ForeColor.RGB) }));
      W.designColorReplace("#ff0000", "#0000ff", 0);
      assert(Number(two.Fill.ForeColor.RGB) === 16711680, "RGB/OLE channel conversion wrong");
      W.designColorAdjust(0, 0, 0); return { oleBlue: Number(two.Fill.ForeColor.RGB) };
    });
    await test("layout-uniform-size", function () {
      select([two, one]); W.designAlignRun("uniform-size", {});
      assert(near(two.Width, one.Width) && near(two.Height, one.Height), "uniform-size mismatch");
    });
    await test("object-manager-bulk-hide-and-lock", function () {
      activate(); let items = W.layerList().items.filter(item => item.name.indexOf("PPTSelfTest_") === 0);
      assert(items.length === 2, "object manager missing test shapes");
      const hidden = W.layerSetVisibleMany(items, false); assert(Number(one.Visible) === 0 && Number(two.Visible) === 0, "bulk hide failed: " + JSON.stringify(hidden) + " values=" + one.Visible + "/" + two.Visible);
      W.layerSetVisibleMany(items, true); assert(Number(one.Visible) !== 0 && Number(two.Visible) !== 0, "bulk show failed");
      W.layerSetLockedMany(items, true); items = W.layerList().items.filter(item => item.name.indexOf("PPTSelfTest_") === 0);
      assert(items.every(item => item.locked), "lock state not retained");
      const blocked = W.layerSelectMany(items); assert(!blocked.ok, "locked shapes selectable through manager");
      one.Select(-1);
      window.__pptTesting = false;
      try { window.__pptEmit("WindowSelectionChange"); } finally { window.__pptTesting = true; }
      let selectedLocked = false;
      try { selectedLocked = Number(app.ActiveWindow.Selection.ShapeRange.Item(1).Id) === Number(one.Id); } catch (_) {}
      assert(!selectedLocked, "canvas selection guard failed to unselect the locked object");
      const originalLeft = Number(one.Left);
      one.Left = originalLeft + 30;
      window.__pptTesting = false;
      try { window.__pptEmit("WindowSelectionChange"); } finally { window.__pptTesting = true; }
      assert(near(one.Left, originalLeft), "locked object geometry was not restored");
      W.layerSetLockedMany(items, false);
      return { lockModes: items.map(item => item.lockMode), canvasGuard: true };
    });
    await test("hidden-content-is-not-a-blank-slide", function () {
      const hiddenSlide = deck.Slides.Add(2, 12);
      const hidden = hiddenSlide.Shapes.AddShape(1, 10, 10, 50, 50); hidden.Visible = 0;
      activate(); W.designCleanup("blank-slides", "all");
      assert(Number(deck.Slides.Count) === 2, "hidden content page deleted");
      hiddenSlide.Delete();
    });
    const source = png("source", "#009966"), replacement = png("replacement", "#660099");
    await test("picture-replacement-crop-geometry-and-link", async function () {
      activate(); picture = slide.Shapes.AddPicture(source, 0, -1, 420, 70, 240, 120); picture.Name = "PPTSelfTest_picture";
      picture.PictureFormat.CropLeft = 20; picture.PictureFormat.CropRight = 10;
      const geom = { left: Number(picture.Left), top: Number(picture.Top), width: Number(picture.Width), height: Number(picture.Height) };
      select([picture]); const result = await W.replaceSelectedFromFile(replacement);
      picture = app.ActiveWindow.Selection.ShapeRange.Item(1);
      assert(near(picture.Left, geom.left) && near(picture.Top, geom.top) && near(picture.Width, geom.width) && near(picture.Height, geom.height), "replacement moved picture");
      assert(Number(picture.PictureFormat.CropLeft) > 0, "crop was lost");
      return { name: String(picture.Name), crop: Number(picture.PictureFormat.CropLeft), resultType: typeof result };
    });
    await test("image-inventory-cache-and-location", async function () {
      activate(); const start = performance.now();
      const first = await W.collectDeckImagesCached(null, null, { force: true });
      const elapsed = performance.now() - start;
      const cachedStart = performance.now();
      const second = await W.collectDeckImagesCached();
      const groups = first.groups || first;
      assert(Array.isArray(groups) && groups.length >= 1, "image inventory empty");
      const thumbnail = groups.flatMap(group => group.instances || []).find(instance => instance.thumb);
      assert(thumbnail && thumbnail.thumb.length > 64, "inventory thumbnail is empty or truncated");
      const thumbImage = new Image();
      await new Promise((resolve, reject) => { thumbImage.onload = resolve; thumbImage.onerror = () => reject(new Error("inventory thumbnail cannot be decoded")); thumbImage.src = thumbnail.thumb; });
      assert(thumbImage.naturalWidth > 0 && thumbImage.naturalHeight > 0, "thumbnail has no pixels");
      await W.locateSlideShape(1, Number(picture.ZOrderPosition), picture);
      assert(Number(app.ActiveWindow.Selection.ShapeRange.Item(1).Id) === Number(picture.Id), "did not select picture on slide");
      return { groups: groups.length, firstMs: Math.round(elapsed), cachedMs: Math.round(performance.now() - cachedStart), cacheHit: Boolean(second.cacheHit) };
    });
    await test("batch-file-keeps-user-input-and-clipboard-replacement", async function () {
      activate(); const copy = picture.Duplicate().Item(1); copy.Name = "PPTSelfTest_batch_copy"; copy.Top = 240;
      select([picture]); const batch = await W.replaceAllFromFile(source, null, () => false);
      assert(Number(batch.success) === 2 && Number(batch.failed) === 0, "batch file replacement failed: " + JSON.stringify(batch));
      assert(host.FileExists(source) && host.FileExists(replacement), "input image file was deleted");
      const named = name => { for(let i=1;i<=Number(slide.Shapes.Count);i++){const shape=slide.Shapes.Item(i);if(String(shape.Name)===name)return shape;}throw new Error("missing shape " + name); };
      picture = named("PPTSelfTest_picture");
      const carrier = slide.Shapes.AddPicture(replacement, 0, -1, 700, 10, 120, 60); carrier.Name = "PPTSelfTest_clipboard_carrier";
      select([picture]); carrier.Copy();
      const clipboardBatch = await W.replaceAllFromClipboard(null, () => false);
      assert(Number(clipboardBatch.success) === 2 && Number(clipboardBatch.failed) === 0, "clipboard batch must only replace the two matching pictures: " + JSON.stringify(clipboardBatch));
      picture = named("PPTSelfTest_picture");
      select([picture]); carrier.Copy(); const before = Number(slide.Shapes.Count);
      await W.replaceSelectedFromClipboard();
      picture = app.ActiveWindow.Selection.ShapeRange.Item(1);
      assert(Number(slide.Shapes.Count) === before && Number(picture.PictureFormat.CropLeft) > 0, "clipboard single replacement/crop failed");
      carrier.Delete(); named("PPTSelfTest_batch_copy").Delete();
      return { fileReplaced: batch.success, clipboardReplaced: clipboardBatch.success, inputFilesPreserved: true };
    });
    await test("native-selection-event-enforces-lock", async function () {
      activate(); const rows = W.layerList().items.filter(item => item.name === "PPTSelfTest_one");
      W.layerSetLockedMany(rows, true);
      window.__pptTesting = false;
      try {
        one.Select(-1); await new Promise(resolve => setTimeout(resolve, 300));
        let selected = false;
        try { selected = Number(app.ActiveWindow.Selection.ShapeRange.Item(1).Id) === Number(one.Id); } catch (_) {}
        assert(!selected, "actual PowerPoint selection event did not enforce lock");
      } finally { window.__pptTesting = true; W.layerSetLockedMany(rows, false); }
      return { event: "WindowSelectionChange", enforced: true };
    });
    await test("photoshop-file-roundtrip-and-document-ownership", async function () {
      activate(); select([picture]); const before = await renderPixels();
      const exe = String(host.TempFolder) + "\\PPTSelfTest_Photoshop.exe"; files.push(exe); host.WriteText(exe, "selftest marker; never executed");
      const originalShell = app.OAAssist.ShellExecute; let launchCount = 0; let opened;
      app.OAAssist.ShellExecute = function () { launchCount++; };
      try { opened = W.designPhotoshopOpen(exe); } finally { app.OAAssist.ShellExecute = originalShell; }
      assert(opened.ok && launchCount === 1, "Photoshop handoff failed"); files.push(opened.path);
      W.designPhotoshopReload(); picture = app.ActiveWindow.Selection.ShapeRange.Item(1);
      const after = await renderPixels(); let delta = 0;
      for(let i=0;i<before.length;i++) delta += Math.abs(before[i]-after[i]);
      delta /= before.length;
      assert(delta < 3, "unmodified PS roundtrip changed visible content: mean pixel delta " + delta);
      const other = app.Presentations.Add(-1); owned.push(other); other.Slides.Add(1, 12);
      let rejected = 0;
      for(const action of [W.designPhotoshopReload, W.designPhotoshopUndo]) { try { action(); } catch (e) { if (/其他文稿/.test(String(e))) rejected++; } }
      assert(rejected === 2, "PS job crossed document ownership");
      activate(); W.designPhotoshopUndo();
      picture = null;
      for(let i=1;i<=Number(slide.Shapes.Count);i++){ const shape=slide.Shapes.Item(i); if(String(shape.Name)==="PPTSelfTest_picture") picture=shape; }
      assert(picture, "PS undo did not restore the picture");
      const undone = await renderPixels(); let undoDelta = 0;
      for(let i=0;i<before.length;i++) undoDelta += Math.abs(before[i]-undone[i]);
      undoDelta /= before.length; assert(undoDelta < 3, "PS undo changed visible content");
      return { meanPixelDelta: delta, undoPixelDelta: undoDelta, crossDocumentRejected: rejected, externalApplicationLaunch: "stubbed" };
    });
    await test("page-stamp-and-slide-extraction", async function () {
      activate(); const count = Number(slide.Shapes.Count); await W.designLayerStamp();
      assert(Number(slide.Shapes.Count) === count + 1, "page stamp missing");
      const before = Number(app.Presentations.Count);
      await W.designExtractSlides("current");
      const extracted = app.ActivePresentation;
      if (Number(app.Presentations.Count) > before) owned.push(extracted);
      assert(Number(extracted.Slides.Count) === 1 && Number(extracted.Slides.Item(1).Shapes.Count) === count + 1, "extracted wrong source deck");
      activate();
    });
    await test("panel-routing-and-retained-engine", function () {
      const original = window.WpsPictureReplace;
      const pane = app.CreateTaskPane(location.origin + "/taskpane.html#layers", "对象管理");
      pane.Visible = true; pane.Visible = false;
      const next = app.CreateTaskPane(location.origin + "/taskpane.html#tools/color", "颜色工具"); next.Visible = true;
      assert(window.WpsPictureReplace === original, "opening pane recreated engine");
      return { persistent: true };
    });
  } catch (error) { tests.push({ name: "suite", ok: false, error: String(error.stack || error) }); }
  finally {
    try { W.smartZoomEnd(); } catch (_) {}
    for (let i = owned.length - 1; i >= 0; i--) {
      try { owned[i].Saved = -1; owned[i].Close(); } catch (error) { tests.push({ name: "close-owned-deck", ok: false, error: String(error) }); }
    }
    for (const path of files) { try { host.RemoveFile(path); } catch (_) {} }
    try { if (originalWindow) originalWindow.Activate(); } catch (_) {}
    window.__pptTesting = false;
    try {
      originalDocuments.forEach(entry => assert(Number(entry.doc.Saved) === entry.saved && Number(entry.doc.Slides.Count) === entry.slides && String(entry.doc.Name) === entry.name, "An original document changed during the test"));
      tests.push({ name: "preserve-user-session", ok: true, detail: { visibleDocumentsPreserved: originalDocuments.length, pooledScratchDocuments: Number(app.Presentations.Count) - originalDocuments.length } });
    }
    catch (error) { tests.push({ name: "preserve-user-session", ok: false, error: String(error) }); }
    host.Report(JSON.stringify({ host: "Microsoft PowerPoint", version: W && W.addinVersion, passed: tests.filter(t => t.ok).length, failed: tests.filter(t => !t.ok).length, tests }, null, 2));
  }
}());
