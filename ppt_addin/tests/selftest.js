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
  let regressionFileSequence = 0;
  async function withInventoryDeck(work) {
    const regressionDeck = app.Presentations.Add(-1);
    owned.push(regressionDeck);
    try {
      regressionDeck.Windows.Item(1).Activate();
      app.ActiveWindow.ViewType = 9;
      W.clearDeckImageCache();
      return await work(regressionDeck);
    } finally {
      try { W.clearDeckImageCache(); } catch (_) {}
      // This deck was allocated above; never close ActivePresentation here.
      try {
        regressionDeck.Saved = -1;
        regressionDeck.Close();
        const ownedIndex = owned.indexOf(regressionDeck);
        if (ownedIndex >= 0) owned.splice(ownedIndex, 1);
      } finally { activate(); }
    }
  }
  function inventoryRow(result, ownerSlide, shape) {
    const key = "slide:" + Number(ownerSlide.SlideID) + ":" + Number(shape.Id);
    const rows = result.groups.flatMap(group => group.instances || []);
    const matches = rows.filter(row => row.sourceKey === key);
    assert(matches.length === 1, "Expected exactly one row for " + key + ", got " + matches.length);
    return matches[0];
  }
  async function previewPixels(url) {
    const img = new Image();
    await new Promise((resolve, reject) => {
      img.onload = resolve;
      img.onerror = () => reject(new Error("Regression preview cannot be decoded"));
      img.src = url;
    });
    assert(img.naturalWidth > 0 && img.naturalHeight > 0, "Regression preview is empty");
    const canvas = document.createElement("canvas"); canvas.width = canvas.height = 64;
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#ffffff"; ctx.fillRect(0, 0, 64, 64);
    const scale = 64 / Math.max(img.naturalWidth, img.naturalHeight);
    const width = img.naturalWidth * scale, height = img.naturalHeight * scale;
    ctx.drawImage(img, (64 - width) / 2, (64 - height) / 2, width, height);
    return ctx.getImageData(0, 0, 64, 64).data;
  }
  async function nativePreviewPixels(shape) {
    const path = String(host.TempFolder) + "\\PPTSelfTest_native_" + Date.now() + "_" + (++regressionFileSequence) + ".png";
    files.push(path);
    shape.Export(path, 2);
    assert(host.FileExists(path), "Native shape preview was not exported");
    return previewPixels("data:image/png;base64," + host.ReadBinaryBase64(path));
  }
  function previewDifference(actual, expected) {
    // Patch averages tolerate native resampling while still exposing a moved
    // corner marker, a missing crop, or color/grayscale cache contamination.
    let total = 0, maximumPatch = 0;
    for (let top = 0; top < 64; top += 8) {
      for (let left = 0; left < 64; left += 8) {
        let patch = 0;
        for (let y = top; y < top + 8; y++) for (let x = left; x < left + 8; x++) {
          const offset = (y * 64 + x) * 4;
          for (let channel = 0; channel < 3; channel++) patch += Math.abs(actual[offset + channel] - expected[offset + channel]);
        }
        total += patch;
        maximumPatch = Math.max(maximumPatch, patch / (8 * 8 * 3));
      }
    }
    return { mean: total / (64 * 64 * 3), maximumPatch };
  }
  function assertPreviewMatches(actual, expected, label) {
    const difference = previewDifference(actual, expected);
    assert(difference.mean < 12 && difference.maximumPatch < 45,
      label + " differs from native Shape.Export: " + JSON.stringify(difference));
    return difference;
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
      assert(near(title,4.8) && near(body,2.4) && near(title / body, 2, 0.1), "table font hierarchy lost: " + title + "/" + body);
      W.smartZoomReset({ _sessionId: info.sessionId }); W.smartZoomEnd(); table.Delete();
      return { title, body };
    });
    await test("small-text-zoom-keeps-line-breaks", function () {
      const box=slide.Shapes.AddTextbox(1,20,20,260,130);
      box.TextFrame.AutoSize=0; box.TextFrame.WordWrap=-1;
      box.TextFrame.TextRange.Text="每一行文字数量需要保持一致。文字按原始比例缩放，文本框和字号同步缩小。";
      box.TextFrame.TextRange.Font.Name="Microsoft YaHei";box.TextFrame.TextRange.Font.Size=24;
      function lines(){const r=box.TextFrame.TextRange;const out=[];for(let i=1;i<=Number(r.Lines().Count);i++)out.push(String(r.Lines(i,1).Text));return out;}
      const original=lines();select([box]);const zoom=W.smartZoomBegin();
      for(const percent of [30,20,10]) {W.smartZoomApply(percent,{_sessionId:zoom.sessionId});assert(JSON.stringify(lines())===JSON.stringify(original),"line breaks changed at "+percent+"%: "+JSON.stringify(lines()));}
      W.smartZoomReset({_sessionId:zoom.sessionId});W.smartZoomEnd();box.Delete();return {percentages:[30,20,10],lines:original.length};
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
    await test("thumbnail-colors-match-six-source-objects", async function () {
      activate();const colors=["#ff0000","#00ff00","#0000ff","#ffff00","#00ffff","#ff00ff"], added=[];
      try{
        colors.forEach((color,i)=>{const path=png("identity_"+i,color);const s=slide.Shapes.AddPicture(path,0,-1,20+i*80,400,70,35);s.Name="PPTIdentity_"+i;if(i===0)s.Flip(0);if(i===1)s.Rotation=90;added.push(s);});
        const result=await W.collectDeckImagesCached(null,null,{force:true});const rows=result.groups.flatMap(g=>g.instances);
        const colorGroups=colors.map((_,i)=>result.groups.find(g=>g.instances.some(r=>r.name==='PPTIdentity_'+i)));
        assert(new Set(colorGroups.map(g=>g&&g.key)).size===6,'distinct color sources were incorrectly merged');
        for(let i=0;i<colors.length;i++){
          const row=rows.find(r=>r.name==="PPTIdentity_"+i);assert(row&&row.thumb,"missing identity thumbnail "+i);
          const img=new Image();await new Promise((resolve,reject)=>{img.onload=resolve;img.onerror=reject;img.src=row.thumb;});
          const canvas=document.createElement('canvas');canvas.width=canvas.height=20;const ctx=canvas.getContext('2d');ctx.drawImage(img,0,0,20,20);const pixel=ctx.getImageData(12,12,1,1).data;
          const target=parseInt(colors[i].slice(1),16),rgb=[target>>16,(target>>8)&255,target&255];assert(rgb.every((v,n)=>Math.abs(v-pixel[n])<35),"thumbnail mismatched "+i+": "+Array.from(pixel));
          if(i===0 || i===1){const marker=ctx.getImageData(i===0?16:12,i===0?7:3,1,1).data;assert(marker[0]>220&&marker[1]>220&&marker[2]>220&&marker[3]>220,"thumbnail orientation lost for "+i+": "+Array.from(marker));}
        }
        return {matched:6,batches:2,flipAndRotation:true};
      } finally {added.forEach(s=>{try{s.Delete();}catch(_){}});}
    });
    await test("inventory-same-name-reorder-location-and-unlink", async function () {
      return withInventoryDeck(async function (regressionDeck) {
        const firstSlide = regressionDeck.Slides.Add(1, 12);
        const otherSlide = regressionDeck.Slides.Add(2, 12);
        const pictures = [];
        for (const owner of [firstSlide, otherSlide]) {
          for (let i = 0; i < 2; i++) {
            const shape = owner.Shapes.AddPicture(i ? replacement : source, 0, -1, 40 + i * 260, 40, 240, 120);
            shape.Name = "PPTSelfTest_duplicate_name";
            shape.AlternativeText = "PICRENEW|" + JSON.stringify({
              v: 1, name: "Same display name", src: i ? replacement : source,
              userAlt: "owned-accessibility-" + pictures.length
            });
            pictures.push(shape);
          }
        }
        assert(pictures.every(shape => String(shape.Name) === "PPTSelfTest_duplicate_name"),
          "Host renamed the duplicate-name fixture");
        app.ActiveWindow.View.GotoSlide(1);
        const initial = await W.collectDeckImagesCached(null, null, { force: true });
        const target = pictures[0];
        const originalRow = inventoryRow(initial, firstSlide, target);
        const locateRow = Object.assign({}, originalRow), unlinkRow = Object.assign({}, originalRow);
        const neighbors = pictures.slice(1).map(shape => ({ shape, id: Number(shape.Id), alt: String(shape.AlternativeText) }));
        const targetId = Number(target.Id), targetSlideId = Number(firstSlide.SlideID);
        const previousPosition = Number(target.ZOrderPosition);
        target.ZOrder(0); // msoBringToFront: the old ordinal now points to the other same-name picture.
        otherSlide.MoveTo(1);
        assert(Number(target.ZOrderPosition) !== previousPosition && Number(firstSlide.SlideIndex) === 2,
          "The regression setup did not reorder both shape and slide ordinals");
        assert(typeof W.locatePictureInstance === "function", "Stable inventory location API is missing");
        assert(await W.locatePictureInstance(locateRow), "Old inventory row could not locate its live target after reorder");
        const selected = app.ActiveWindow.Selection.ShapeRange.Item(1);
        assert(Number(selected.Id) === targetId && Number(selected.Parent.SlideID) === targetSlideId &&
          Number(app.ActiveWindow.View.Slide.SlideID) === targetSlideId,
        "Same-name/ordinal fallback selected a different picture or slide");
        const unlinked = await W.unlinkInstances([unlinkRow], initial.docKey);
        assert(Number(unlinked) === 1 && String(target.AlternativeText) === "owned-accessibility-0",
          "Unlink did not target the original picture and preserve its accessibility text");
        neighbors.forEach(entry => assert(Number(entry.shape.Id) === entry.id && String(entry.shape.AlternativeText) === entry.alt,
          "Unlink modified a same-name neighbor"));
        target.Delete();
        app.ActiveWindow.View.GotoSlide(Number(otherSlide.SlideIndex));
        pictures[3].Select(-1);
        let rejected = false;
        try { rejected = !(await W.locatePictureInstance(Object.assign({}, originalRow))); } catch (_) { rejected = true; }
        assert(rejected, "A deleted inventory target was silently redirected to a same-name neighbor");
        const retainedSelection = app.ActiveWindow.Selection.ShapeRange.Item(1);
        assert(Number(retainedSelection.Id) === Number(pictures[3].Id) &&
          Number(retainedSelection.Parent.SlideID) === Number(otherSlide.SlideID),
        "Rejected location still changed the canvas selection");
        return { duplicateNames: pictures.length, shapeOrderChanged: true, slideOrderChanged: true, unlinked, deletedTargetRejected: true };
      });
    });
    await test("inventory-appearance-cache-matches-native-preview", async function () {
      return withInventoryDeck(async function (regressionDeck) {
        const owner = regressionDeck.Slides.Add(1, 12);
        const imagePath = png("appearance", "#e53624");
        const shapes = [0, 1, 2].map(i => {
          const shape = owner.Shapes.AddPicture(imagePath, 0, -1, 30 + i * 260, 100, 240, 120);
          shape.Name = "PPTSelfTest_appearance_" + i;
          return shape;
        });
        app.ActiveWindow.View.GotoSlide(1);
        const initial = await W.collectDeckImagesCached(null, null, { force: true });
        const oldGray = await previewPixels(inventoryRow(initial, owner, shapes[1]).thumb);
        const oldRotated = await previewPixels(inventoryRow(initial, owner, shapes[2]).thumb);
        // All three objects retain the same source media. Appearance belongs to
        // each instance and must never be reused merely because the source matches.
        shapes[1].PictureFormat.ColorType = 2; // msoPictureGrayscale
        shapes[1].PictureFormat.Brightness = 0.4;
        shapes[2].PictureFormat.CropLeft = 35;
        shapes[2].PictureFormat.CropRight = 10;
        shapes[2].Rotation = 90;
        const native = [];
        for (const shape of shapes) native.push(await nativePreviewPixels(shape));
        assert(previewDifference(oldGray, native[1]).mean > 20, "Grayscale fixture did not visibly change");
        assert(previewDifference(oldRotated, native[2]).maximumPatch > 60, "Rotation/crop fixture did not visibly change");
        // Deliberately omit force: this exercises stale cache validation.
        const refreshed = await W.collectDeckImagesCached();
        const reopened = await W.collectDeckImagesCached();
        const differences = [];
        for (let i = 0; i < shapes.length; i++) {
          const row = inventoryRow(refreshed, owner, shapes[i]);
          const cached = inventoryRow(reopened, owner, shapes[i]);
          differences.push(assertPreviewMatches(await previewPixels(row.thumb), native[i], "Instance " + i));
          assertPreviewMatches(await previewPixels(cached.thumb), native[i], "Reopened instance " + i);
        }
        return { sameSourceInstances: shapes.length, grayscale: true, rotation: 90, cropped: true,
          forceRefreshUsedAfterEdit: false, reopenedCacheHit: Boolean(reopened.cacheHit), differences };
      });
    });
    await test("inventory-stale-appearance-rejects-whole-batch", async function () {
      return withInventoryDeck(async function (regressionDeck) {
        const owner = regressionDeck.Slides.Add(1, 12);
        const unchanged = owner.Shapes.AddPicture(source, 0, -1, 40, 60, 240, 120);
        const changed = owner.Shapes.AddPicture(source, 0, -1, 340, 60, 240, 120);
        unchanged.Name = "PPTSelfTest_stale_unchanged"; changed.Name = "PPTSelfTest_stale_changed";
        unchanged.AlternativeText = "unchanged-accessibility"; changed.AlternativeText = "changed-accessibility";
        app.ActiveWindow.View.GotoSlide(1);
        const initial = await W.collectDeckImagesCached(null, null, { force: true });
        const first = inventoryRow(initial, owner, unchanged), second = inventoryRow(initial, owner, changed);
        assert(/^sha256:[a-f0-9]{64}$/i.test(first.sourceIdentity) && first.sourceIdentity === second.sourceIdentity,
          "Stale-batch fixture requires a verified common source identity");
        assert(first.sourceAppearance && second.sourceAppearance, "Appearance identities are missing");
        assert(initial.groups.some(group => group.instances.some(row => row.sourceKey === first.sourceKey) &&
          group.instances.some(row => row.sourceKey === second.sourceKey)), "The fixture was not grouped before editing");
        const ids = [Number(unchanged.Id), Number(changed.Id)];
        const geometry = [Number(changed.Left), Number(changed.Top), Number(changed.Width), Number(changed.Height)];
        changed.PictureFormat.Brightness = 0.25;
        assert(Number(changed.Id) === ids[1] && geometry.every((value, i) =>
          near(value, [changed.Left, changed.Top, changed.Width, changed.Height][i])), "Appearance edit changed identity or geometry");
        const before = [await nativePreviewPixels(unchanged), await nativePreviewPixels(changed)];
        let error = "";
        try {
          // Put the stale target second: preflight must reject the whole group
          // before replacing the first, still-valid picture.
          await W.replaceInstances([first, second], replacement, initial.docKey);
        } catch (failure) { error = String(failure && failure.message || failure); }
        assert(/变化|刷新|失效/.test(error), "A stale appearance was not explicitly rejected: " + error);
        assert(Number(owner.Shapes.Count) === 2 && Number(unchanged.Id) === ids[0] && Number(changed.Id) === ids[1],
          "A rejected batch still replaced or removed a picture");
        assert(String(unchanged.AlternativeText) === "unchanged-accessibility" && String(changed.AlternativeText) === "changed-accessibility",
          "A rejected batch modified accessibility/link metadata");
        for (let i = 0; i < 2; i++) {
          const after = await nativePreviewPixels(i ? changed : unchanged);
          assert(previewDifference(after, before[i]).mean < 0.5, "A rejected batch changed picture pixels: " + i);
        }
        return { unchangedIds: ids, unchangedGeometry: true, rejectedBeforeAnyReplacement: true, reason: error };
      });
    });
    await test("inventory-master-reopen-reuses-verified-live-cache", async function () {
      return withInventoryDeck(async function (regressionDeck) {
        regressionDeck.Slides.Add(1,12);
        const master=regressionDeck.SlideMaster;
        const image=master.Shapes.AddPicture(source,0,-1,20,20,120,60);
        image.Name="PPTSelfTest_master_picture";
        const first=await W.collectDeckImagesCached(null,null,{force:true});
        const second=await W.collectDeckImagesCached();
        const row=second.groups.flatMap(group=>group.instances||[]).find(item=>item.kind==="master" && Number(item.shapeIdentity)===Number(image.Id));
        assert(row && row.containerReference,"master cache lost its live ownership proof");
        assert(second.cacheHit,"reopening a list with master pictures rescanned instead of reusing memory");
        assert(window.__pptSameObject(row.shape,image),"master cache points to a different picture");
        return {total:first.total,cacheHit:true,liveOwnershipVerified:true};
      });
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
    await test("panel-routing-and-retained-engine", async function () {
      const original = window.WpsPictureReplace;
      const pane = app.CreateTaskPane(location.origin + "/taskpane.html#layers", "对象管理");
      pane.Visible = true; pane.Visible = false;
      const next = app.CreateTaskPane(location.origin + "/taskpane.html#tools/color", "颜色工具"); next.Visible = true;
      assert(window.WpsPictureReplace === original, "opening pane recreated engine");
      await window.__pptDispatch('DesignExportToolsOpen');
      assert(location.hash === '#tools/export', 'export route is malformed');
      await window.__pptDispatch('OpenPicturePanelButton');
      assert(location.hash === '#panel', 'image panel cannot reopen after export route');
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
