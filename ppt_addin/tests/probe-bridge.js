(function () {
  const h = window.__pptHost, p = window.chrome.webview.hostObjects.sync.pptApplication;
  const r = {}; let d;
  function probe(name, fn) { try { const value = fn(); r[name] = { type: typeof value, value: String(value) }; } catch (e) { r[name] = { error: String(e) }; } }
  try {
    d = p.Presentations.Add(-1); const s = d.Slides.Add(1, 12), a = s.Shapes.AddShape(1, 40, 40, 200, 60);
    a.TextFrame.TextRange.Text = "bridge sample"; a.Fill.ForeColor.RGB = 255;
    const tr = a.TextFrame.TextRange, color = a.Fill.ForeColor;
    probe("shape-type", () => typeof a);
    probe("text-auto", () => tr.Text);
    probe("text-explicit", () => tr.getHostProperty("Text"));
    probe("text-call", () => tr.Text());
    probe("rgb-auto", () => color.RGB);
    probe("rgb-explicit", () => color.getHostProperty("RGB"));
    probe("rgb-call", () => color.RGB());
    probe("frame-call", () => a.TextFrame());
    probe("textrange-call", () => a.TextFrame.TextRange());
    probe("text2-call", () => a.TextFrame2.TextRange.Text());
    probe("rgb-invoke", () => color.applyHostFunction("RGB"));
    probe("width-explicit", () => a.getHostProperty("Width"));
    probe("frame-explicit", () => a.getHostProperty("TextFrame"));
    probe("method-property", () => a.getHostProperty("Select"));
    probe("unknown-property", () => a.getHostProperty("NoSuchProperty"));
    probe("unknown-auto", () => a.NoSuchProperty);
    probe("item-property", () => s.Shapes.getHostProperty("Item"));
    probe("text-after-set-explicit", () => { tr.setHostProperty("Text", "real text"); return tr.getHostProperty("Text"); });
  } catch (e) { r.fatal = String(e); }
  finally { if (d) { d.Saved = -1; d.Close(); } h.Report(JSON.stringify(r, null, 2)); }
}());
