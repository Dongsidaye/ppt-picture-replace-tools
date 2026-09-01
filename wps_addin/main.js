/* 东四大爷的工具箱 (Picture Replace Tools) - WPS WPP JavaScript add-in.
 * v1.1.6 - WPS JSAPI compatibility fix:
 *  - WPS JSAPI Shape has NO SaveAsPicture method and PictureFormat has no
 *    Crop sub-object (official WPS docs). All raster export now goes through
 *    a scratch presentation + Slide.Export, and crop preservation only uses
 *    the documented scalar properties CropLeft/Right/Top/Bottom.
 */
(function (global) {
  "use strict";





  const PP_PASTE_PNG = 6;
  const PP_PASTE_BITMAP = 1;
  const PP_PASTE_JPG = 5;
  const PP_PASTE_GIF = 3;
  const MsoFalse = 0;
  const MsoTrue = -1;
  const MsoFlipHorizontal = 0;
  const MsoFlipVertical = 1;
  const MsoSendBackward = 3;
  const MAX_BROWSER_FILE_BYTES = 50 * 1024 * 1024;
  const PREVIEW_PX = 320;
  const CLIPBOARD_MAX_PX = 1600;
  const NORMALIZE_PX = 1024;
  const MAX_SCRATCH_SLIDE = 2000;
  const PROJECT_HOME_URL = "https://github.com/Dongsidaye/ppt-picture-replace-tools";

  // The panel displays thumbnails at roughly 84x60 CSS pixels. 96px cells
  // retain enough detail while the responsive 2x2 renderer exports only a
  // 192x192 grid per UI time slice.
  const THUMB_PX = 96;
  const LINK_PREFIX = "PICRENEW|";
  const LINK_VERSION = 1;

  function application() {
    const root = global.wps || global.Application;
    if (!root) throw new Error("WPS JSAPI 尚未就绪。");
    return root;
  }

  function fileSystem() {
    const fs = application().FileSystem;
    if (!fs) throw new Error("当前 WPS 版本没有开放 FileSystem API。");
    return fs;
  }

  function hasMethod(object, name) {
    try { return !!object && typeof object[name] === "function"; } catch (_) { return false; }
  }

  // Diagnostic trace helpers. No-op unless a probe flag explicitly sets a
  // tracePath; they exist so probe code never throws "trace is not defined".
  let TRACE_PATH = "";
  function setTracePath(path) {
    try { TRACE_PATH = String(path || ""); } catch (_) { TRACE_PATH = ""; }
  }
  function trace(message) {
    try {
      if (!TRACE_PATH) return;
      const fs = fileSystem();
      const line = new Date().toISOString() + " " + String(message) + "\r\n";
      try {
        if (fs.Exists && fs.Exists(TRACE_PATH)) {
          const prev = String(fs.readAsBinaryString(TRACE_PATH) || "");
          fs.writeAsBinaryString(TRACE_PATH, prev + line);
        } else if (fs.writeAsBinaryString) {
          fs.writeAsBinaryString(TRACE_PATH, line);
        } else if (fs.WriteFile) {
          fs.WriteFile(TRACE_PATH, line);
        }
      } catch (_) {}
    } catch (_) {}
  }

  function num(value) {
    const n = Number(value);
    return isFinite(n) ? n : 0;
  }

  function isTrue(value) {
    return value === true || value === MsoTrue || value === 1;
  }

  function asShape(value) {
    if (!value) return value;
    try {
      if (value.Item && value.Count !== undefined) {
        const first = value.Item(1);
        if (first) return first;
      }
    } catch (_) {}
    return value;
  }

  function presentationOf(slide) {
    try {
      const p = slide.Parent;
      if (p && p.PageSetup) return p;
    } catch (_) {}
    try { return slide.Presentation; } catch (_) {}
    return null;
  }

  function capabilityProbe() {
    const result = {
      host: "WPS WPP JSAPI",
      application: false,
      version: "",
      activePresentation: false,
      fileSystem: false,
      addPicture: false,
      paste: false,
      pasteSpecial: false,
      slideExport: false,
      copy: false,
      scaleWidth: false,
      crop: false,
      pictureFound: false,
      taskPane: false,
      fileDialog: false,
      currentAddInPath: "",
      currentAddInName: "",
      errors: []
    };
    let app;
    try {
      app = application();
      result.application = true;
      result.version = String(app.Version || app.Build || "");
      result.taskPane = hasMethod(app, "CreateTaskPane");
      result.fileDialog = hasMethod(app, "FileDialog");
      try {
        const addin = app.CurrentWPSAddIn;
        if (addin) {
          result.currentAddInPath = String(addin.Path || "");
          result.currentAddInName = String(addin.Name || "");
        }
      } catch (_) {}
      const fs = app.FileSystem;
      result.fileSystem = !!fs &&
        (hasMethod(fs, "readAsBinaryString") || hasMethod(fs, "ReadFile")) &&
        (hasMethod(fs, "writeAsBinaryString") || hasMethod(fs, "WriteFile"));
      const presentation = app.ActivePresentation;
      result.activePresentation = !!presentation;
      if (presentation && Number(presentation.Slides.Count) > 0) {
        const slideCount = Number(presentation.Slides.Count) || 0;
        for (let slideIndex = 1; slideIndex <= slideCount; slideIndex += 1) {
          const slide = presentation.Slides.Item(slideIndex);
          if (!slide || !slide.Shapes) continue;
          if (!result.addPicture) result.addPicture = hasMethod(slide.Shapes, "AddPicture");
          if (!result.paste) result.paste = hasMethod(slide.Shapes, "Paste");
          if (!result.pasteSpecial) result.pasteSpecial = hasMethod(slide.Shapes, "PasteSpecial");
          if (!result.slideExport) result.slideExport = hasMethod(slide, "Export");
          const shapeCount = Number(slide.Shapes.Count) || 0;
          for (let shapeIndex = 1; shapeIndex <= shapeCount; shapeIndex += 1) {
            try {
              const shape = slide.Shapes.Item(shapeIndex);
              if (!shape || !shape.PictureFormat) continue;
              result.pictureFound = true;
              try {
                const pf = shape.PictureFormat;
                if (pf && pf.CropLeft !== undefined) result.crop = true;
              } catch (_) {}
              if (!result.copy) result.copy = hasMethod(shape, "Copy");
              if (!result.scaleWidth) result.scaleWidth = hasMethod(shape, "ScaleWidth");
            } catch (_) {}
          }
        }
      }
    } catch (error) {
      result.errors.push(error && error.message ? String(error.message) : String(error));
    }
    result.ready = result.application && result.fileSystem && result.addPicture &&
      result.paste && result.slideExport;
    return result;
  }

  function capabilityText() {
    const c = capabilityProbe();
    const yes = value => value ? "可用" : "不可用";
    const cropStatus = c.pictureFound ? yes(c.crop) : "未检测到图片";
    const conclusion = c.pictureFound
      ? (c.ready ? "核心替换 API 已就绪" : "当前环境不满足核心替换 API")
      : (c.activePresentation ? "宿主基础 API 可用，当前文稿未检测到图片" : "请先打开一个 WPS 演示文稿");
    return [
      c.host + (c.version ? " " + c.version : ""),
      "FileSystem: " + yes(c.fileSystem),
      "AddPicture: " + yes(c.addPicture),
      "Shapes.Paste: " + yes(c.paste),
      "Slide.Export: " + yes(c.slideExport),
      "PictureFormat.CropLeft: " + cropStatus,
      "PasteSpecial: " + yes(c.pasteSpecial),
      "CreateTaskPane: " + yes(c.taskPane),
      "FileDialog: " + yes(c.fileDialog),
      c.currentAddInPath ? "AddIn.Path: " + c.currentAddInPath : "",
      c.currentAddInName ? "AddIn.Name: " + c.currentAddInName : "",
      "结论: " + conclusion,
      c.errors.length ? "错误: " + c.errors.join(" | ") : ""
    ].filter(Boolean).join("\n");
  }

  function tell(message, title) {
    try {
      if (application().alert) application().alert(String(message));
      else if (global.alert) global.alert(String(message));
    } catch (_) {
      if (global.console) global.console.error(title || "东四大爷的工具箱", message);
    }
  }

  function activePresentation() {
    const p = application().ActivePresentation;
    if (!p) throw new Error("请先打开一个 WPS 演示文稿。");
    return p;
  }

  function selectedPicture() {
    const windowObject = application().ActiveWindow;
    const selection = windowObject && windowObject.Selection;
    if (!selection || !selection.ShapeRange || selection.ShapeRange.Count < 1) {
      throw new Error("请先选中一张图片。");
    }
    const shape = asShape(selection.ShapeRange);
    if (!isPicture(shape)) throw new Error("当前选中对象不是图片。");
    return shape;
  }

  // =====================================================================
  // Object filtering (WPS)
  // =====================================================================
  // Object filtering is deliberately selection-only: it reads the active
  // slide/master/layout and changes only the native ShapeRange. It never
  // writes shape properties or presentation content.
  const OBJECT_FILTER_EPSILON = 0.01;
  const OBJECT_FILTER_LINE_TYPE = 9;   // msoLine
  const OBJECT_FILTER_GROUP_TYPE = 6;  // msoGroup

  function objectFilterReadPath(object, path) {
    let current = object;
    try {
      for (let i = 0; i < path.length; i += 1) {
        if (current === null || current === undefined) return { ok: false };
        current = current[path[i]];
      }
      return current === null || current === undefined
        ? { ok: false }
        : { ok: true, value: current };
    } catch (_) { return { ok: false }; }
  }

  function objectFilterReadNumber(object, paths) {
    for (let i = 0; i < paths.length; i += 1) {
      const result = objectFilterReadPath(object, paths[i]);
      if (!result.ok) continue;
      const value = Number(result.value);
      if (isFinite(value)) return { ok: true, value: value };
    }
    return { ok: false };
  }

  function objectFilterShapeType(shape) {
    const result = objectFilterReadNumber(shape, [["Type"], ["type"]]);
    return result.ok ? result.value : null;
  }

  function objectFilterShapeHasText(shape) {
    const flagPaths = [
      ["TextFrame2", "HasText"],
      ["TextFrame", "HasText"]
    ];
    for (let i = 0; i < flagPaths.length; i += 1) {
      const result = objectFilterReadPath(shape, flagPaths[i]);
      if (!result.ok) continue;
      if (typeof result.value === "string") {
        if (result.value.trim()) return true;
      } else if (Number(result.value) === 1 || Number(result.value) === -1) {
        return true;
      }
    }
    const textPaths = [
      ["TextFrame2", "TextRange", "Text"],
      ["TextFrame", "TextRange", "Text"]
    ];
    for (let i = 0; i < textPaths.length; i += 1) {
      const result = objectFilterReadPath(shape, textPaths[i]);
      if (result.ok && String(result.value || "").trim()) return true;
    }
    return false;
  }

  function objectFilterFontSize(shape) {
    const paths = [
      ["TextFrame2", "TextRange", "Font", "Size"],
      ["TextFrame", "TextRange", "Font", "Size"],
      ["TextFrame", "Characters", "Font", "Size"]
    ];
    const result = objectFilterReadNumber(shape, paths);
    return result.ok && result.value > 0 ? result.value : null;
  }

  function objectFilterColorValue(shape, paths) {
    const result = objectFilterReadNumber(shape, paths);
    return result.ok && result.value >= 0 && result.value <= 0xffffff ? result.value : null;
  }

  function objectFilterColor(shape) {
    const type = objectFilterShapeType(shape);
    if (objectFilterShapeHasText(shape)) {
      const textColor = objectFilterColorValue(shape, [
        ["TextFrame2", "TextRange", "Font", "Fill", "ForeColor", "RGB"],
        ["TextFrame", "TextRange", "Font", "Color", "RGB"]
      ]);
      if (textColor !== null) return { mode: "text", value: textColor };
    }
    if (type === OBJECT_FILTER_LINE_TYPE) {
      const lineColor = objectFilterColorValue(shape, [["Line", "ForeColor", "RGB"]]);
      if (lineColor !== null) return { mode: "line", value: lineColor };
    }
    const fillColor = objectFilterColorValue(shape, [["Fill", "ForeColor", "RGB"]]);
    if (fillColor !== null) return { mode: "fill", value: fillColor };
    const lineColor = objectFilterColorValue(shape, [["Line", "ForeColor", "RGB"]]);
    return lineColor === null ? null : { mode: "line", value: lineColor };
  }

  function objectFilterSelectedShapes() {
    const windowObject = application().ActiveWindow;
    const range = windowObject && windowObject.Selection && windowObject.Selection.ShapeRange;
    if (!range) return [];
    let count = 0;
    try { count = Number(range.Count) || 0; } catch (_) { count = 0; }
    const result = [];
    for (let i = 1; i <= count; i += 1) {
      let shape = null;
      try {
        if (hasMethod(range, "Item")) shape = range.Item(i);
        else shape = range[i - 1];
      } catch (_) {}
      if (shape) result.push(shape);
    }
    return result;
  }

  function objectFilterShapeKey(shape) {
    let id = "";
    let name = "";
    try { id = String(shape.Id || shape.id || ""); } catch (_) {}
    try { name = String(shape.Name || shape.name || ""); } catch (_) {}
    return id ? "id:" + id : (name ? "name:" + name : "");
  }

  function objectFilterShapeId(shape) {
    try {
      const value = String(shape && (shape.Id || shape.id) || "").trim();
      return value || "";
    } catch (_) { return ""; }
  }

  function objectFilterContainsShape(shapes, target) {
    const targetId = objectFilterShapeId(target);
    for (let i = 0; i < shapes.length; i += 1) {
      if (shapes[i] === target) return true;
      // Names are user-editable and may collide.  Only use the documented
      // Shape.Id fallback across COM proxies; otherwise require the same
      // object reference instead of treating two same-named objects as one.
      if (targetId && objectFilterShapeId(shapes[i]) === targetId) return true;
    }
    return false;
  }

  function objectFilterCurrentContainer(selected) {
    if (selected && selected.length) {
      try {
        const parent = selected[0].Parent;
        if (parent && parent.Shapes) return parent;
      } catch (_) {}
      try {
        const parent = selected[0].Parent;
        if (parent && parent.Parent && parent.Parent.Shapes) return parent.Parent;
      } catch (_) {}
    }
    const app = application();
    const windowObject = app.ActiveWindow;
    const view = windowObject && windowObject.View;
    let viewSlide = null;
    try { viewSlide = view && view.Slide; } catch (_) {}
    if (viewSlide && viewSlide.Shapes) return viewSlide;
    let slideIndex = null;
    try { slideIndex = Number(viewSlide && viewSlide.SlideIndex); } catch (_) {}
    if (!(slideIndex > 0)) {
      try { slideIndex = Number(view && view.SlideIndex); } catch (_) {}
    }
    if (!(slideIndex > 0)) {
      try { slideIndex = Number(view && view.current); } catch (_) {}
    }
    const presentation = activePresentation();
    const slideCount = Number(presentation.Slides.Count) || 0;
    if (!(slideIndex > 0) && slideCount === 1) slideIndex = 1;
    if (slideIndex > 0) return presentation.Slides.Item(slideIndex);
    throw new Error("无法确定当前幻灯片，请先进入普通幻灯片视图。");
  }

  function objectFilterShapes(container) {
    const shapes = container && container.Shapes ? container.Shapes : container;
    if (!shapes) return [];
    let count = 0;
    try { count = Number(shapes.Count) || 0; } catch (_) { count = 0; }
    const result = [];
    for (let i = 1; i <= count; i += 1) {
      try {
        const shape = hasMethod(shapes, "Item") ? shapes.Item(i) : shapes[i - 1];
        if (shape) result.push(shape);
      } catch (_) {}
    }
    return result;
  }

  function objectFilterSelectShapes(container, targets) {
    if (!targets || !targets.length) return false;
    const shapes = container && container.Shapes ? container.Shapes : container;
    const names = targets.map(function (shape) {
      try { return String(shape.Name || shape.name || ""); } catch (_) { return ""; }
    }).filter(Boolean);
    // A Shape.Name is only unique inside a slide when WPS has assigned it
    // correctly.  User-renamed objects can share a name, in which case a
    // names-based Range would silently select the wrong set.  Fall back to
    // the object references for that case so batch selection remains safe.
    const namesUnique = names.length === targets.length && new Set(names).size === names.length;
    let namesUniqueInContainer = namesUnique;
    if (namesUnique && shapes && hasMethod(shapes, "Item")) {
      const wanted = Object.create(null);
      names.forEach(function (name) { wanted[name] = 0; });
      let shapeCount = 0;
      try { shapeCount = Number(shapes.Count) || 0; } catch (_) { shapeCount = 0; }
      for (let i = 1; i <= shapeCount; i += 1) {
        let candidate = null;
        try { candidate = shapes.Item(i); } catch (_) {}
        if (!candidate) continue;
        let candidateName = "";
        try { candidateName = String(candidate.Name || candidate.name || ""); } catch (_) {}
        if (Object.prototype.hasOwnProperty.call(wanted, candidateName)) wanted[candidateName] += 1;
      }
      namesUniqueInContainer = names.every(function (name) { return wanted[name] === 1; });
    }
    if (namesUniqueInContainer && shapes && hasMethod(shapes, "Range")) {
      try {
        const range = shapes.Range(names.length === 1 ? names[0] : names);
        if (range && hasMethod(range, "Select")) {
          range.Select(MsoTrue);
          return true;
        }
      } catch (_) {}
    }
    let selected = false;
    for (let i = 0; i < targets.length; i += 1) {
      const shape = targets[i];
      if (!shape || !hasMethod(shape, "Select")) continue;
      try {
        shape.Select(i === 0 ? MsoTrue : MsoFalse);
        selected = true;
      } catch (_) {
        try {
          if (i === 0) { shape.Select(); selected = true; }
        } catch (__) {}
      }
    }
    return selected;
  }

  function objectFilterSelectionMatches(targets) {
    const windowObject = application().ActiveWindow;
    const range = windowObject && windowObject.Selection && windowObject.Selection.ShapeRange;
    if (!range) return null;
    let count = 0;
    try { count = Number(range.Count) || 0; } catch (_) { return false; }
    if (count !== targets.length) return false;
    const actual = [];
    for (let i = 1; i <= count; i += 1) {
      try {
        const shape = hasMethod(range, "Item") ? range.Item(i) : range[i - 1];
        if (shape) actual.push(shape);
      } catch (_) { return false; }
    }
    return targets.every(function (shape) { return objectFilterContainsShape(actual, shape); });
  }

  function objectFilterRun(mode) {
    const selected = objectFilterSelectedShapes();
    const container = objectFilterCurrentContainer(selected);
    const all = objectFilterShapes(container);
    if (!all.length) return { ok: false, count: 0, message: "当前页没有可筛选的对象。" };

    const needsReference = mode === "type" || mode === "fontsize" || mode === "width" ||
      mode === "height" || mode === "color";
    const reference = needsReference ? selected[0] : null;
    if (needsReference && !reference) {
      return { ok: false, count: 0, message: "请先选中一个对象作为匹配基准。" };
    }
    const referenceType = reference ? objectFilterShapeType(reference) : null;
    const referenceSize = reference ? objectFilterFontSize(reference) : null;
    const referenceColor = reference ? objectFilterColor(reference) : null;
    if (mode === "type" && referenceType === null) return { ok: false, count: 0, message: "无法读取基准对象类型。" };
    if (mode === "fontsize" && referenceSize === null) return { ok: false, count: 0, message: "基准对象没有可读取的单一字号。" };
    if (mode === "color" && !referenceColor) return { ok: false, count: 0, message: "无法读取基准对象颜色。" };

    const matches = [];
    all.forEach(function (shape) {
      let match = false;
      if (mode === "all") {
        match = true;
      } else if (mode === "invert") {
        match = !objectFilterContainsShape(selected, shape);
      } else if (mode === "type") {
        match = objectFilterShapeType(shape) === referenceType;
      } else if (mode === "fontsize") {
        const size = objectFilterFontSize(shape);
        match = size !== null && Math.abs(size - referenceSize) <= OBJECT_FILTER_EPSILON;
      } else if (mode === "width" || mode === "height") {
        const property = mode === "width" ? "Width" : "Height";
        const current = objectFilterReadNumber(shape, [[property], [property.toLowerCase()]]);
        const base = objectFilterReadNumber(reference, [[property], [property.toLowerCase()]]);
        match = current.ok && base.ok && Math.abs(current.value - base.value) <= OBJECT_FILTER_EPSILON;
      } else if (mode === "color") {
        const color = objectFilterColor(shape);
        match = !!color && color.mode === referenceColor.mode && color.value === referenceColor.value;
      } else if (mode === "line") {
        match = objectFilterShapeType(shape) === OBJECT_FILTER_LINE_TYPE;
      } else if (mode === "text") {
        match = objectFilterShapeHasText(shape);
      } else if (mode === "group") {
        match = objectFilterShapeType(shape) === OBJECT_FILTER_GROUP_TYPE;
      }
      if (match) matches.push(shape);
    });

    const selectableMatches = matches.filter(function (shape) { return !layerGuardShapeLocked(shape); });
    if (selectableMatches.length !== matches.length) {
      matches.length = 0;
      selectableMatches.forEach(function (shape) { matches.push(shape); });
    }
    if (!matches.length) {
      return { ok: false, count: 0, message: mode === "invert" ? "反选结果为空。" : "当前页没有符合条件的对象。" };
    }
    if (!objectFilterSelectShapes(container, matches)) {
      return { ok: false, count: 0, message: "宿主没有提供可写入的对象选区。" };
    }
    const verified = objectFilterSelectionMatches(matches);
    if (verified === false) {
      return { ok: false, count: 0, message: "宿主未确认目标对象已全部选中。" };
    }
    return { ok: true, count: matches.length, message: "已选择 " + matches.length + " 个对象。" };
  }

  function objectFilterExecuteMso(commandId) {
    const app = application();
    try {
      if (app.CommandBars && hasMethod(app.CommandBars, "ExecuteMso")) {
        app.CommandBars.ExecuteMso(commandId);
        return true;
      }
    } catch (_) {}
    try {
      if (hasMethod(app, "ExecuteMso")) {
        app.ExecuteMso(commandId);
        return true;
      }
    } catch (_) {}
    return false;
  }

  // WPS names the native animation pane control AnimationCustom.  A few
  // older builds exposed the Office-compatible AnimationPane alias, so keep
  // it as a guarded fallback without making it the primary command.
  function objectFilterExecuteMsoAny(commandIds) {
    const ids = Array.isArray(commandIds) ? commandIds : [commandIds];
    for (let i = 0; i < ids.length; i += 1) {
      if (objectFilterExecuteMso(ids[i])) return String(ids[i]);
    }
    return "";
  }

  function notifyObjectFilter(result) {
    if (!result || !result.ok) tell(result && result.message ? result.message : "对象筛选失败。", "对象筛选");
    else tell(result.message, "对象筛选");
  }

  function OpenAnimationPane() {
    runAsync(function () {
      if (!objectFilterExecuteMsoAny(["AnimationCustom", "AnimationPane"])) tell("当前 WPS 未开放动画窗格命令。", "对象筛选");
    });
  }

  function OpenSelectionPane() {
    runAsync(function () { openPane("#layers", "对象管理"); });
  }

  function SelectAllObjects() { runAsync(function () { notifyObjectFilter(objectFilterRun("all")); }); }
  function InvertSelection() { runAsync(function () { notifyObjectFilter(objectFilterRun("invert")); }); }
  function SelectSameType() { runAsync(function () { notifyObjectFilter(objectFilterRun("type")); }); }
  function SelectSameFontSize() { runAsync(function () { notifyObjectFilter(objectFilterRun("fontsize")); }); }
  function SelectSameWidth() { runAsync(function () { notifyObjectFilter(objectFilterRun("width")); }); }
  function SelectSameHeight() { runAsync(function () { notifyObjectFilter(objectFilterRun("height")); }); }
  function SelectSameColor() { runAsync(function () { notifyObjectFilter(objectFilterRun("color")); }); }
  function SelectAllLines() { runAsync(function () { notifyObjectFilter(objectFilterRun("line")); }); }
  function SelectAllText() { runAsync(function () { notifyObjectFilter(objectFilterRun("text")); }); }
  function SelectAllGroups() { runAsync(function () { notifyObjectFilter(objectFilterRun("group")); }); }

  // =====================================================================
  // Object manager (WPS)
  // =====================================================================
  // WPS exposes Shape.Visible as a real, writable property.  Its current WPP
  // builds expose a Shape.Locked member but do not persist writes to it, so a
  // layer row must never present a silent fake native lock.  We first probe
  // the host's round-trip support when the user presses the lock button.  If
  // that probe fails, the lock state is stored in a namespaced Shape.Tags
  // entry (which survives save/reopen) and the pane labels it as a plugin
  // lock.  This gives users a durable, honest state while remaining safe on
  // hosts that do not implement native object locking.
  const LAYER_LOCK_TAG = "CODEXLAYERLOCKED";
  const LAYER_MAX_OBJECTS = 500;
  // Keep object categories stable across the pane and the batch-selection
  // API.  The colors are UI metadata only; no slide shape formatting is
  // modified when an object is classified.
  const LAYER_TYPE_DEFINITIONS = [
    { key: "image", label: "图片", color: "#1677ff" },
    { key: "text", label: "文字", color: "#8b5cf6" },
    { key: "table", label: "表格", color: "#0f9d58" },
    { key: "chart", label: "图表", color: "#f08c00" },
    { key: "shape", label: "形状", color: "#e11d48" },
    { key: "line", label: "线条", color: "#0891b2" },
    { key: "group", label: "组合", color: "#7c3aed" },
    { key: "media", label: "媒体", color: "#db2777" },
    { key: "smartart", label: "SmartArt", color: "#059669" },
    { key: "ole", label: "嵌入对象", color: "#64748b" },
    { key: "placeholder", label: "占位符", color: "#a16207" },
    { key: "canvas", label: "画布", color: "#475569" },
    { key: "diagram", label: "图示", color: "#0284c7" },
    { key: "ink", label: "墨迹", color: "#be123c" },
    { key: "comment", label: "批注", color: "#ca8a04" },
    { key: "other", label: "其他对象", color: "#6b7280" }
  ];
  const LAYER_TYPE_BY_KEY = Object.create(null);
  LAYER_TYPE_DEFINITIONS.forEach(function (definition) { LAYER_TYPE_BY_KEY[definition.key] = definition; });
  let layerNativeLockSupport = null;
  const layerMemoryLocks = Object.create(null);
  // A Shape.Tags marker survives save/reopen, while this reference list also
  // covers hosts that expose neither writable Shape.Locked nor Shape.Tags.
  // The selection guard below uses both forms so a lock is enforced at the
  // canvas boundary instead of being only a visual state in the pane.
  const layerSessionLockedShapes = [];
  let layerLockGuardBound = false;
  let layerLockGuardSupported = null;
  let layerLockGuardError = "";
  let layerLockGuardBusy = false;

  function layerSessionShapeKey(shape) {
    let parent = null;
    try { parent = shape && shape.Parent; } catch (_) {}
    const parentKey = layerContainerIdentity(parent);
    const shapeKey = objectFilterShapeKey(shape);
    return (parentKey || "") + "|" + (shapeKey || "");
  }

  function layerShapePosition(shape) {
    let left = NaN, top = NaN;
    try { left = Number(shape.Left); } catch (_) {}
    try { top = Number(shape.Top); } catch (_) {}
    if (!isFinite(left) || !isFinite(top)) return null;
    return { left: left, top: top };
  }

  function layerRememberSessionLock(shape, desired) {
    if (!shape) return;
    const key = layerSessionShapeKey(shape);
    let existing = null;
    for (let i = layerSessionLockedShapes.length - 1; i >= 0; i -= 1) {
      const entry = layerSessionLockedShapes[i];
      if (entry.shape === shape || (key && entry.key && key === entry.key)) {
        if (!desired) { layerSessionLockedShapes.splice(i, 1); continue; }
        if (!existing) existing = entry;
        else layerSessionLockedShapes.splice(i, 1);
      }
    }
    if (!desired) return;
    // Keep the original baseline: re-registering an already locked shape must
    // not adopt a position it was dragged to as the new restore target.
    if (existing) {
      if (!existing.pos) existing.pos = layerShapePosition(shape);
      return;
    }
    layerSessionLockedShapes.push({ shape: shape, key: key, pos: layerShapePosition(shape) });
  }

  // Live probes against the installed WPS build show Shape.Locked is a stub
  // (reads -1, writes ignored) and a:spLocks noSelect="1" in the slide XML is
  // not honored either, so a drag that starts on a locked shape still moves
  // it. The selection guard cannot intercept the drag itself; instead the
  // locked shape's position is snapshotted when the lock is recorded and the
  // shape is snapped back on the next selection change, so an accidental
  // drag does not permanently break the layout.
  function layerRestoreLockedGeometry() {
    for (let i = 0; i < layerSessionLockedShapes.length; i += 1) {
      const entry = layerSessionLockedShapes[i];
      if (!entry || !entry.shape || !entry.pos) continue;
      try {
        const left = Number(entry.shape.Left);
        if (isFinite(left) && Math.abs(left - entry.pos.left) > 0.5) entry.shape.Left = entry.pos.left;
      } catch (_) {}
      try {
        const top = Number(entry.shape.Top);
        if (isFinite(top) && Math.abs(top - entry.pos.top) > 0.5) entry.shape.Top = entry.pos.top;
      } catch (_) {}
    }
  }

  function layerIsSessionLocked(shape) {
    if (!shape) return false;
    const key = layerSessionShapeKey(shape);
    for (let i = 0; i < layerSessionLockedShapes.length; i += 1) {
      const entry = layerSessionLockedShapes[i];
      if (entry.shape === shape) return true;
      if (key && entry.key && key === entry.key) return true;
    }
    return false;
  }

  function layerGuardShapeLocked(shape) {
    if (!shape) return false;
    const tag = layerReadTagLock(shape);
    if (tag.locked) return true;
    if (layerIsSessionLocked(shape)) return true;
    if (layerNativeLockSupport === true) {
      try { return isTrue(shape.Locked); } catch (_) {}
    }
    return false;
  }

  function layerSelectionShapes(selection) {
    let source = selection || null;
    if (!source) {
      try {
        const windowObject = application().ActiveWindow;
        source = windowObject && windowObject.Selection;
      } catch (_) { source = null; }
    }
    let range = null;
    try { range = source && source.ShapeRange; } catch (_) { range = null; }
    if (!range) {
      try { range = source && source.Selection && source.Selection.ShapeRange; } catch (_) { range = null; }
    }
    if (!range && source && source.Count !== undefined && hasMethod(source, "Item")) range = source;
    if (!range) return [];
    let count = 0;
    try { count = Number(range.Count) || 0; } catch (_) { count = 0; }
    const shapes = [];
    for (let i = 1; i <= count; i += 1) {
      try {
        const shape = hasMethod(range, "Item") ? range.Item(i) : range[i - 1];
        if (shape) shapes.push(shape);
      } catch (_) {}
    }
    return shapes;
  }

  function layerClearSelection(windowObject) {
    const selection = windowObject && windowObject.Selection;
    const candidates = [selection];
    try { if (selection && selection.ShapeRange) candidates.push(selection.ShapeRange); } catch (_) {}
    for (let i = 0; i < candidates.length; i += 1) {
      const candidate = candidates[i];
      if (!candidate) continue;
      const methods = ["Unselect", "ClearShapeSelect", "ClearSelection"];
      for (let j = 0; j < methods.length; j += 1) {
        const method = methods[j];
        if (!hasMethod(candidate, method)) continue;
        try {
          candidate[method]();
          return true;
        } catch (_) {}
      }
    }
    return false;
  }

  function layerHandleSelectionChange(selection) {
    if (layerLockGuardBusy) return;
    const guardStartedAt = Date.now();
    // Snap drag-moved locked shapes back before evaluating the selection.
    try { layerRestoreLockedGeometry(); } catch (_) {}
    // Selection changes must NOT schedule inventory preloads: the scan walks
    // every slide over the JSAPI bridge and renders thumbnails via clipboard
    // Copy/Paste + Slide.Export, so starting it 650ms after a click caused
    // periodic stutters (and clipboard churn) in the middle of editing.
    // Preloads are driven by load/activate/open events instead, and the pane
    // runs its own foreground scan when it is opened with a cold cache.
    const shapes = layerSelectionShapes(selection);
    if (!shapes.length) { perfTraceTime("guard.empty", guardStartedAt); return; }
    const locked = shapes.some(layerGuardShapeLocked);
    if (!locked) { perfTraceTime("guard", guardStartedAt); return; }
    const allowed = shapes.filter(function (shape) { return !layerGuardShapeLocked(shape); });
    const windowObject = (function () {
      try { return application().ActiveWindow; } catch (_) { return null; }
    }());
    layerLockGuardBusy = true;
    try {
      const cleared = layerClearSelection(windowObject);
      // If the host lacks an explicit Unselect method, selecting the allowed
      // subset with Replace=true still removes locked objects from a mixed
      // selection. With only locked objects, the guard can report a degraded
      // capability rather than pretending the object was cleared.
      if (allowed.length) {
        for (let i = 0; i < allowed.length; i += 1) {
          const shape = allowed[i];
          if (!shape || !hasMethod(shape, "Select")) continue;
          try { shape.Select(i === 0 ? MsoTrue : MsoFalse); } catch (_) {}
        }
      } else if (!cleared) {
        layerLockGuardSupported = false;
        layerLockGuardError = "WPS 选择对象未提供 Unselect/ClearShapeSelect";
      }
    } finally {
      layerLockGuardBusy = false;
      perfTraceTime("guard.locked", guardStartedAt);
    }
  }

  function layerEnsureSelectionGuard() {
    if (layerLockGuardBound) return true;
    let apiEvent = null;
    try { apiEvent = application().ApiEvent; } catch (_) { apiEvent = null; }
    if (!apiEvent || !hasMethod(apiEvent, "AddApiEventListener")) {
      layerLockGuardSupported = false;
      layerLockGuardError = "当前 WPS 未开放 ApiEvent.AddApiEventListener";
      return false;
    }
    try {
      apiEvent.AddApiEventListener("WindowSelectionChange", layerHandleSelectionChange);
      layerLockGuardBound = true;
      layerLockGuardSupported = true;
      layerLockGuardError = "";
      return true;
    } catch (error) {
      layerLockGuardSupported = false;
      layerLockGuardError = String(error && error.message || error);
      return false;
    }
  }

  function layerContainerIdentity(container) {
    let id = "";
    let name = "";
    try { id = String(container && (container.Id || container.id) || ""); } catch (_) {}
    try { name = String(container && (container.Name || container.name) || ""); } catch (_) {}
    return id ? "id:" + id : (name ? "name:" + name : "");
  }

  function layerSameContainer(a, b) {
    if (!a || !b) return false;
    if (a === b) return true;
    const ak = layerContainerIdentity(a);
    const bk = layerContainerIdentity(b);
    if (ak && ak === bk) return true;
    // Some WPS builds return a fresh COM proxy for Shape.Parent on every
    // access and omit a stable Name/Id.  A slide index is still a safe
    // container identity in normal view; without this fallback a valid batch
    // selection would be rejected as an expired context.
    let ai = 0;
    let bi = 0;
    try { ai = Number(a.SlideIndex || a.slideIndex || 0); } catch (_) {}
    try { bi = Number(b.SlideIndex || b.slideIndex || 0); } catch (_) {}
    return ai > 0 && bi > 0 && ai === bi;
  }

  function layerShapeKey(shape, context, shapeIndex) {
    const contextKey = context ? (String(context.kind || "") + ":" + String(context.layoutIndex || context.slideIndex || 0)) : "";
    const shapeKey = objectFilterShapeKey(shape);
    return contextKey + ":" + (shapeKey || "index:" + String(shapeIndex || 0));
  }

  function layerShapeName(shape, shapeIndex) {
    let value = "";
    try { value = String(shape && (shape.Name || shape.name) || "").trim(); } catch (_) {}
    return value || "对象 " + String(shapeIndex || "");
  }

  function layerReadCapability(shape, paths) {
    for (let i = 0; i < paths.length; i += 1) {
      const result = objectFilterReadPath(shape, paths[i]);
      if (!result.ok) continue;
      const value = result.value;
      if (typeof value === "boolean") {
        if (value) return true;
        continue;
      }
      if (typeof value === "number") {
        if (value !== 0) return true;
        continue;
      }
      if (typeof value === "string") {
        if (value.trim()) return true;
        continue;
      }
      // WPS returns COM proxy objects for Table/Chart/MediaFormat/etc.
      if (value !== null && value !== undefined) return true;
    }
    return false;
  }

  function layerTypeDefinition(key) {
    return LAYER_TYPE_BY_KEY[key] || LAYER_TYPE_BY_KEY.other;
  }

  function layerShapeCategory(shape) {
    const type = objectFilterShapeType(shape);
    // Prefer the documented MsoShapeType value first.  This keeps the common
    // path to one cheap read per shape; capability probes are reserved for
    // ambiguous/unknown hosts because each COM property read can be costly.
    if (type === 13 || type === 11) return layerTypeDefinition("image");
    if (type === 19) return layerTypeDefinition("table");
    if (type === 3) return layerTypeDefinition("chart");
    if (type === 16 || type === 26) return layerTypeDefinition("media");
    if (type === 24) return layerTypeDefinition("smartart");
    if (type === 21) return layerTypeDefinition("diagram");
    if (type === 7 || type === 10 || type === 12) {
      return layerReadCapability(shape, [["Table"]]) ? layerTypeDefinition("table") : layerTypeDefinition("ole");
    }
    if (type === OBJECT_FILTER_LINE_TYPE) return layerTypeDefinition("line");
    if (type === OBJECT_FILTER_GROUP_TYPE) return layerTypeDefinition("group");
    if (type === 17) return layerTypeDefinition("text");
    if (type === 14) return layerTypeDefinition("placeholder");
    if (type === 20) return layerTypeDefinition("canvas");
    if (type === 22 || type === 23) return layerTypeDefinition("ink");
    if (type === 4) return layerTypeDefinition("comment");
    const hasText = objectFilterShapeHasText(shape);
    if (hasText) return layerTypeDefinition("text");
    if (type === 1 || type === 2 || type === 5 || type === 8 || type === 15) return layerTypeDefinition("shape");
    // Unknown type values occur on newer WPS builds.  Use capability probes
    // only here so those objects still land in a useful category.
    if (layerReadCapability(shape, [["Table"]])) return layerTypeDefinition("table");
    if (layerReadCapability(shape, [["Chart"], ["ChartData"]])) return layerTypeDefinition("chart");
    if (layerReadCapability(shape, [["MediaFormat"], ["MediaType"]])) return layerTypeDefinition("media");
    if (layerReadCapability(shape, [["SmartArt"]])) return layerTypeDefinition("smartart");
    if (layerReadCapability(shape, [["OLEFormat"]])) return layerTypeDefinition("ole");
    return layerTypeDefinition("other");
  }

  function layerShapeTypeLabel(shape, knownCategory) {
    const category = knownCategory || layerShapeCategory(shape);
    const type = objectFilterShapeType(shape);
    if (category.key === "other" && type !== null) return category.label + " · 类型 " + String(type);
    return category.label;
  }

  function layerReadVisible(shape) {
    try {
      if (shape.Visible === undefined || shape.Visible === null) return { supported: false, visible: true };
      return { supported: true, visible: isTrue(shape.Visible) };
    } catch (_) { return { supported: false, visible: true }; }
  }

  function layerReadTagLock(shape) {
    let tags = null;
    try { tags = shape && shape.Tags; } catch (_) { tags = null; }
    if (!tags || !hasMethod(tags, "Item")) return { supported: false, locked: false };
    try {
      const raw = tags.Item(LAYER_LOCK_TAG);
      const value = String(raw === undefined || raw === null ? "" : raw).trim().toLowerCase();
      return { supported: true, locked: value === "1" || value === "true" || value === "yes" };
    } catch (_) { return { supported: true, locked: false }; }
  }

  function layerLockState(shape, context, shapeIndex) {
    const tag = layerReadTagLock(shape);
    const memoryKey = layerShapeKey(shape, context, shapeIndex);
    if (tag.locked) {
      // Rebuild the session entry (position baseline included) so a persisted
      // plugin lock also gains the selection and geometry guards after the
      // presentation is reopened.
      try { layerRememberSessionLock(shape, true); } catch (_) {}
      return { locked: true, mode: "plugin", native: false, tagSupported: tag.supported };
    }
    if (layerMemoryLocks[memoryKey]) return { locked: true, mode: "session", native: false, tagSupported: tag.supported };
    if (layerNativeLockSupport === true) {
      try {
        const native = isTrue(shape.Locked);
        return { locked: native, mode: native ? "native" : "none", native: true, tagSupported: tag.supported };
      } catch (_) {}
    }
    return { locked: false, mode: "none", native: false, tagSupported: tag.supported };
  }

  function layerCurrentContext() {
    const app = application();
    const windowObject = app.ActiveWindow;
    if (!windowObject) throw new Error("请先打开 WPS 演示文稿窗口。");
    const presentation = activePresentation();
    const viewType = readWindowViewType(windowObject);
    if (viewType === 2) {
      let container = null;
      let kind = "master";
      let layoutIndex = 0;
      let layoutName = "";
      const selected = objectFilterSelectedShapes();
      if (selected.length) {
        try {
          const parent = selected[0].Parent;
          if (parent && parent.Shapes) container = parent;
          else if (parent && parent.Parent && parent.Parent.Shapes) container = parent.Parent;
        } catch (_) {}
      }
      // A stale normal-slide selection can remain exposed while WPS is
      // switching into master view. Do not let that slide become the layer
      // source; accept only the actual master or one of its custom layouts.
      if (container) {
        let validMasterContainer = false;
        try { validMasterContainer = layerSameContainer(container, presentation.SlideMaster); } catch (_) {}
        if (!validMasterContainer) {
          try {
            const layouts = presentation.SlideMaster && presentation.SlideMaster.CustomLayouts;
            const count = Math.min(Number(layouts && layouts.Count) || 0, 200);
            for (let i = 1; i <= count; i += 1) {
              let layout = null;
              try { layout = layouts.Item(i); } catch (_) { continue; }
              if (layerSameContainer(layout, container)) { validMasterContainer = true; break; }
            }
          } catch (_) {}
        }
        if (!validMasterContainer) container = null;
      }
      if (!container) {
        try { container = presentation.SlideMaster; } catch (_) { container = null; }
      }
      if (!container || !container.Shapes) throw new Error("无法读取当前母版或版式的对象列表。");
      try {
        const layouts = presentation.SlideMaster && presentation.SlideMaster.CustomLayouts;
        const count = Math.min(Number(layouts && layouts.Count) || 0, 200);
        for (let i = 1; i <= count; i += 1) {
          let layout = null;
          try { layout = layouts.Item(i); } catch (_) { continue; }
          if (!layerSameContainer(layout, container)) continue;
          kind = "layout";
          layoutIndex = i;
          try { layoutName = String(layout.Name || layout.name || ""); } catch (_) {}
          break;
        }
      } catch (_) {}
      return {
        window: windowObject,
        presentation: presentation,
        viewType: 2,
        kind: kind,
        label: kind === "layout" ? ("版式" + (layoutName ? " · " + layoutName : "")) : "母版",
        container: container,
        slideIndex: 0,
        layoutIndex: layoutIndex,
        layoutName: layoutName
      };
    }

    let slide = null;
    let slideIndex = null;
    try {
      const viewSlide = windowObject.View && windowObject.View.Slide;
      if (viewSlide && viewSlide.Shapes) slide = viewSlide;
      slideIndex = Number(viewSlide && viewSlide.SlideIndex);
    } catch (_) {}
    if (!(slideIndex > 0)) slideIndex = readWindowSlideIndex(windowObject);
    if (!(slideIndex > 0)) {
      try { slideIndex = Number(windowObject.View && windowObject.View.current); } catch (_) {}
    }
    const slideCount = Number(presentation.Slides.Count) || 0;
    if (!(slideIndex > 0) && slideCount === 1) slideIndex = 1;
    if (!slide && slideIndex > 0) {
      try { slide = presentation.Slides.Item(slideIndex); } catch (_) { slide = null; }
    }
    if (!slide || !slide.Shapes) throw new Error("无法确定当前幻灯片，请先进入普通编辑视图。");
    return {
      window: windowObject,
      presentation: presentation,
      viewType: viewType,
      kind: "slide",
      label: "普通页 · 第 " + String(slideIndex) + " 页",
      container: slide,
      slideIndex: Number(slideIndex),
      layoutIndex: 0,
      layoutName: ""
    };
  }

  function layerList() {
    const selectionGuard = layerEnsureSelectionGuard();
    const context = layerCurrentContext();
    const all = objectFilterShapes(context.container);
    const limit = Math.min(all.length, LAYER_MAX_OBJECTS);
    const firstIndex = Math.max(0, all.length - limit);
    const items = [];
    // Native Selection Pane shows the topmost object first. Preserve each
    // original 1-based Shape index for subsequent COM operations.
    for (let i = all.length - 1; i >= firstIndex; i -= 1) {
      const shape = all[i];
      const index = i + 1;
      const visible = layerReadVisible(shape);
      const locked = layerLockState(shape, context, index);
      const category = layerShapeCategory(shape);
      items.push({
        shape: shape,
        shapeIndex: index,
        order: all.length - i,
        id: objectFilterShapeKey(shape),
        name: layerShapeName(shape, index),
        type: objectFilterShapeType(shape),
        typeKey: category.key,
        typeColor: category.color,
        typeLabel: layerShapeTypeLabel(shape, category),
        visible: visible.visible,
        visibleSupported: visible.supported,
        locked: locked.locked,
        lockMode: locked.mode,
        nativeLock: locked.native,
        tagSupported: locked.tagSupported,
        kind: context.kind,
        label: context.label,
        slideIndex: context.slideIndex,
        layoutIndex: context.layoutIndex,
        layoutName: context.layoutName
      });
    }
    const grouped = Object.create(null);
    items.forEach(function (item) {
      const key = item.typeKey || "other";
      if (!grouped[key]) grouped[key] = { count: 0, unlockedCount: 0, lockedCount: 0, visibleCount: 0, hiddenCount: 0 };
      grouped[key].count += 1;
      if (item.locked) grouped[key].lockedCount += 1;
      else grouped[key].unlockedCount += 1;
      if (item.visible) grouped[key].visibleCount += 1;
      else grouped[key].hiddenCount += 1;
    });
    const groups = [];
    LAYER_TYPE_DEFINITIONS.forEach(function (definition) {
      const counts = grouped[definition.key];
      if (!counts) return;
      groups.push({
        key: definition.key,
        label: definition.label,
        color: definition.color,
        count: counts.count,
        unlockedCount: counts.unlockedCount,
        lockedCount: counts.lockedCount,
        visibleCount: counts.visibleCount,
        hiddenCount: counts.hiddenCount
      });
    });
    return {
      ok: true,
      kind: context.kind,
      label: context.label,
      slideIndex: context.slideIndex,
      layoutIndex: context.layoutIndex,
      layoutName: context.layoutName,
      count: items.length,
      total: all.length,
      truncated: all.length > limit,
      nativeLockSupported: layerNativeLockSupport === true,
      nativeLockKnown: layerNativeLockSupport !== null,
      selectionGuardSupported: selectionGuard && layerLockGuardSupported === true,
      selectionGuardKnown: layerLockGuardSupported !== null,
      selectionGuardError: layerLockGuardError,
      tagSupported: items.some(function (item) { return item.tagSupported; }),
      groups: groups,
      items: items
    };
  }

  function layerResolveShape(item) {
    if (item && typeof item === "object") {
      try { if (item.shape) return item.shape; } catch (_) {}
      try { if (item.Shape) return item.Shape; } catch (_) {}
    }
    const context = layerCurrentContext();
    const index = Number(item && (item.shapeIndex || item.index) || item);
    if (!(index > 0)) return null;
    try { return context.container.Shapes.Item(index); } catch (_) { return null; }
  }

  function layerItemContext(item, shape) {
    const context = {
      kind: item && item.kind ? String(item.kind) : "slide",
      slideIndex: Number(item && item.slideIndex) || 0,
      layoutIndex: Number(item && item.layoutIndex) || 0
    };
    if (!item || !item.kind) {
      try {
        const parent = shape && shape.Parent;
        if (parent && parent.CustomLayout) context.kind = "layout";
      } catch (_) {}
    }
    return context;
  }

  function layerSetVisible(item, value) {
    const shape = layerResolveShape(item);
    if (!shape) return { ok: false, message: "对象已不存在，请刷新对象列表。" };
    const desired = !!value;
    return layerApplyVisible(shape, desired);
  }

  function layerApplyVisible(shape, desired) {
    try {
      const before = layerReadVisible(shape);
      shape.Visible = desired ? MsoTrue : MsoFalse;
      const state = layerReadVisible(shape);
      if (!state.supported || state.visible !== desired) return { ok: false, message: "WPS 未确认对象显示状态已改变。" };
      invalidatePanelInventoryCache();
      return { ok: true, visible: state.visible, changed: !before.supported || before.visible !== desired };
    } catch (error) {
      return { ok: false, message: String(error && error.message || error) };
    }
  }

  function layerSetVisibleMany(items, value) {
    const list = Array.isArray(items) ? items : [];
    if (!list.length) return { ok: false, updated: 0, total: 0, message: "当前分类没有可管理的对象。" };
    const context = layerCurrentContext();
    const desired = !!value;
    const seen = [];
    const updatedIndexes = [];
    let updated = 0;
    let changed = 0;
    let skippedMissing = 0;
    let skippedContext = 0;
    let failed = 0;
    let firstError = "";
    for (let i = 0; i < list.length; i += 1) {
      const item = list[i];
      const shape = layerResolveShape(item);
      if (!shape) { skippedMissing += 1; continue; }
      if (!layerItemMatchesContext(item, shape, context)) { skippedContext += 1; continue; }
      if (objectFilterContainsShape(seen, shape)) continue;
      seen.push(shape);
      const result = layerApplyVisible(shape, desired);
      if (result && result.ok) {
        item.visible = result.visible;
        updatedIndexes.push(i);
        updated += 1;
        if (result.changed) changed += 1;
      } else {
        failed += 1;
        if (!firstError) firstError = result && result.message || "WPS 未确认对象显示状态。";
      }
    }
    if (!updated) {
      return {
        ok: false,
        updated: 0,
        changed: 0,
        total: list.length,
        failed: failed,
        skippedMissing: skippedMissing,
        skippedContext: skippedContext,
        message: firstError || "WPS 未确认任何对象显示状态已改变。"
      };
    }
    invalidatePanelInventoryCache();
    const skipped = skippedMissing + skippedContext;
    return {
      ok: true,
      desired: desired,
      updated: updated,
      changed: changed,
      failed: failed,
      total: list.length,
      updatedIndexes: updatedIndexes,
      skippedMissing: skippedMissing,
      skippedContext: skippedContext,
      message: (desired ? "已显示 " : "已隐藏 ") + updated + " 个对象"
        + (changed && changed !== updated ? "（其中 " + changed + " 个状态发生变化）" : "")
        + (failed ? "，" + failed + " 个失败" : "")
        + (skipped ? "，跳过 " + skipped + " 个无效对象" : "") + "。"
    };
  }

  function layerNativeLockRead(shape) {
    try {
      if (shape.Locked === undefined || shape.Locked === null) return { ok: false };
      return { ok: true, value: isTrue(shape.Locked) };
    } catch (_) { return { ok: false }; }
  }

  function layerTryNativeLock(shape, desired) {
    if (layerNativeLockSupport === false) return null;
    const before = layerNativeLockRead(shape);
    if (!before.ok) { layerNativeLockSupport = false; return null; }
    const beforeValue = before.value;
    const probeValue = !beforeValue;
    let probeRead = null;
    let restoreRead = null;
    try {
      shape.Locked = probeValue;
      probeRead = layerNativeLockRead(shape);
      shape.Locked = beforeValue;
      restoreRead = layerNativeLockRead(shape);
    } catch (_) {
      layerNativeLockSupport = false;
      return null;
    }
    if (!probeRead || !probeRead.ok || probeRead.value !== probeValue || !restoreRead || !restoreRead.ok || restoreRead.value !== beforeValue) {
      layerNativeLockSupport = false;
      return null;
    }
    layerNativeLockSupport = true;
    try {
      shape.Locked = !!desired;
      const verify = layerNativeLockRead(shape);
      if (verify.ok && verify.value === !!desired) return { ok: true, mode: "native", locked: verify.value };
    } catch (_) {}
    layerNativeLockSupport = false;
    return null;
  }

  function layerWriteTagLock(shape, desired, context, shapeIndex) {
    let tags = null;
    try { tags = shape && shape.Tags; } catch (_) { tags = null; }
    if (!tags || !hasMethod(tags, "Item")) return { ok: false, supported: false };
    try {
      if (desired) {
        if (hasMethod(tags, "Delete")) { try { tags.Delete(LAYER_LOCK_TAG); } catch (_) {} }
        if (!hasMethod(tags, "Add")) return { ok: false, supported: false };
        tags.Add(LAYER_LOCK_TAG, "1");
      } else if (hasMethod(tags, "Delete")) {
        try { tags.Delete(LAYER_LOCK_TAG); } catch (_) {}
      } else {
        return { ok: false, supported: false };
      }
      const verify = layerReadTagLock(shape);
      if (verify.supported && verify.locked === !!desired) return { ok: true, supported: true };
    } catch (_) {}
    return { ok: false, supported: true };
  }

  function layerSetLocked(item, value) {
    const shape = layerResolveShape(item);
    if (!shape) return { ok: false, message: "对象已不存在，请刷新对象列表。" };
    const desired = !!value;
    return layerApplyLocked(shape, item, desired);
  }

  function layerApplyLocked(shape, item, desired) {
    const selectionGuard = layerEnsureSelectionGuard();
    const native = layerTryNativeLock(shape, desired);
    if (native && native.ok) {
      // Remove a previous plugin marker when a later WPS build supports the
      // real property, otherwise the stale marker would keep the row looking
      // locked after the native object was unlocked.
      try { layerWriteTagLock(shape, false, layerItemContext(item, shape), Number(item && (item.shapeIndex || item.index) || 0)); } catch (_) {}
      layerRememberSessionLock(shape, desired);
      return {
        ok: true,
        locked: native.locked,
        mode: native.mode,
        selectionGuardSupported: selectionGuard && layerLockGuardSupported === true,
        message: desired
          ? (selectionGuard ? "已使用 WPS 原生锁定，并启用选中守卫。" : "已使用 WPS 原生锁定。")
          : "已解除 WPS 原生锁定。"
      };
    }
    const context = layerItemContext(item, shape);
    const index = Number(item && (item.shapeIndex || item.index) || 0);
    const tag = layerWriteTagLock(shape, desired, context, index);
    const memoryKey = layerShapeKey(shape, context, index);
    if (tag.ok) {
      if (desired) layerMemoryLocks[memoryKey] = true;
      else delete layerMemoryLocks[memoryKey];
      layerRememberSessionLock(shape, desired);
      return {
        ok: true,
        locked: desired,
        mode: "plugin",
        selectionGuardSupported: selectionGuard && layerLockGuardSupported === true,
        message: desired
          ? (selectionGuard && layerLockGuardSupported === true
            ? "已锁定对象；WPS 原生锁定不可用，已启用选中守卫，画布不能选中或移动该对象。"
            : "已记录插件锁定标记，但当前 WPS 未提供选中守卫，无法保证阻止画布直接编辑。")
          : "已解除插件锁定标记。"
      };
    }
    if (desired) layerMemoryLocks[memoryKey] = true;
    else delete layerMemoryLocks[memoryKey];
    layerRememberSessionLock(shape, desired);
    return {
      ok: true,
      locked: desired,
      mode: "session",
      selectionGuardSupported: selectionGuard && layerLockGuardSupported === true,
      message: desired
        ? (selectionGuard && layerLockGuardSupported === true
          ? "已锁定对象；当前 WPS 未提供持久化锁定，已启用本次会话选中守卫，画布不能选中或移动该对象。"
          : "已记录本次会话锁定，但当前 WPS 未提供选中守卫，无法保证阻止画布直接编辑。")
        : "已解除本次会话锁定。"
    };
  }

  function layerSetLockedMany(items, value) {
    const list = Array.isArray(items) ? items : [];
    if (!list.length) return { ok: false, updated: 0, total: 0, message: "当前分类没有可管理的对象。" };
    const context = layerCurrentContext();
    const selectionGuard = layerEnsureSelectionGuard();
    const desired = !!value;
    const seen = [];
    const updatedIndexes = [];
    const modeCounts = { native: 0, plugin: 0, session: 0 };
    let updated = 0;
    let skippedMissing = 0;
    let skippedContext = 0;
    let failed = 0;
    let firstError = "";
    for (let i = 0; i < list.length; i += 1) {
      const item = list[i];
      const shape = layerResolveShape(item);
      if (!shape) { skippedMissing += 1; continue; }
      if (!layerItemMatchesContext(item, shape, context)) { skippedContext += 1; continue; }
      if (objectFilterContainsShape(seen, shape)) continue;
      seen.push(shape);
      const result = layerApplyLocked(shape, item, desired);
      if (result && result.ok) {
        item.locked = result.locked;
        item.lockMode = result.mode || "session";
        item.nativeLock = result.mode === "native";
        updatedIndexes.push(i);
        updated += 1;
        modeCounts[result.mode] = (modeCounts[result.mode] || 0) + 1;
      } else {
        failed += 1;
        if (!firstError) firstError = result && result.message || "WPS 未确认对象锁定状态。";
      }
    }
    if (!updated) {
      return {
        ok: false,
        updated: 0,
        total: list.length,
        failed: failed,
        skippedMissing: skippedMissing,
        skippedContext: skippedContext,
        selectionGuardSupported: selectionGuard && layerLockGuardSupported === true,
        message: firstError || "WPS 未确认任何对象锁定状态已改变。"
      };
    }
    invalidatePanelInventoryCache();
    const modes = [];
    if (modeCounts.native) modes.push("原生 " + modeCounts.native);
    if (modeCounts.plugin) modes.push("插件 " + modeCounts.plugin);
    if (modeCounts.session) modes.push("会话 " + modeCounts.session);
    const skipped = skippedMissing + skippedContext;
    return {
      ok: true,
      desired: desired,
      updated: updated,
      failed: failed,
      total: list.length,
      modeCounts: modeCounts,
      updatedIndexes: updatedIndexes,
      skippedMissing: skippedMissing,
      skippedContext: skippedContext,
      selectionGuardSupported: selectionGuard && layerLockGuardSupported === true,
      message: (desired ? "已锁定 " : "已解锁 ") + updated + " 个对象"
        + (modes.length ? "（" + modes.join(" / ") + "）" : "")
        + (failed ? "，" + failed + " 个失败" : "")
        + (skipped ? "，跳过 " + skipped + " 个无效对象" : "") + "。"
    };
  }

  async function layerSelect(item) {
    const shape = layerResolveShape(item);
    const windowObject = application().ActiveWindow;
    if (!shape || !windowObject || !windowObject.View) return false;
    if (layerGuardShapeLocked(shape)) return false;
    const kind = String(item && item.kind || "slide");
    if (kind === "slide") {
      const targetIndex = Number(item && item.slideIndex);
      if (!(targetIndex > 0)) return false;
      if (!await requestWindowViewType(windowObject, NORMAL_VIEW_TYPE)) return false;
      activateWindow(windowObject);
      try { windowObject.View.GotoSlide(targetIndex); } catch (_) { return false; }
      const current = readWindowSlideIndex(windowObject);
      if (current === null) await yieldUI(90);
      else if (!await waitForWindowState(function () { return readWindowSlideIndex(windowObject); }, targetIndex, 1500)) return false;
      return selectCanvasShapeAsync(windowObject, shape, Number(item && item.shapeIndex) || 0);
    }
    if (!await requestWindowViewType(windowObject, 2)) return false;
    if (kind === "layout") {
      try {
        const parent = shape.Parent;
        if (parent && hasMethod(parent, "Select")) parent.Select(MsoTrue);
      } catch (_) {}
    }
    await yieldUI(70);
    return selectCanvasShapeAsync(windowObject, shape, Number(item && item.shapeIndex) || 0);
  }

  function layerItemMatchesContext(item, shape, context) {
    if (item && item.kind && String(item.kind) !== String(context.kind)) return false;
    if (context.kind === "slide" && item && Number(item.slideIndex) > 0 && Number(item.slideIndex) !== Number(context.slideIndex)) return false;
    if (context.kind === "layout" && item && Number(item.layoutIndex) > 0 && Number(item.layoutIndex) !== Number(context.layoutIndex)) return false;
    try {
      const parent = shape && shape.Parent;
      if (parent && context.container && !layerSameContainer(parent, context.container)) return false;
    } catch (_) {}
    return true;
  }

  // Select a set of objects already displayed by the object manager.  The
  // method intentionally works only in the current slide/master/layout so a
  // stale pane cannot accidentally change another document region.
  function layerSelectMany(items) {
    const list = Array.isArray(items) ? items : [];
    if (!list.length) return { ok: false, count: 0, skippedLocked: 0, message: "当前分类没有可选对象。" };
    const context = layerCurrentContext();
    const candidates = [];
    const seen = [];
    let skippedLocked = 0;
    let skippedMissing = 0;
    let skippedContext = 0;
    for (let i = 0; i < list.length; i += 1) {
      const item = list[i];
      const shape = layerResolveShape(item);
      if (!shape) { skippedMissing += 1; continue; }
      if (!layerItemMatchesContext(item, shape, context)) { skippedContext += 1; continue; }
      if (layerGuardShapeLocked(shape)) { skippedLocked += 1; continue; }
      if (objectFilterContainsShape(seen, shape)) continue;
      seen.push(shape);
      candidates.push(shape);
    }
    if (!candidates.length) {
      return {
        ok: false,
        count: 0,
        skippedLocked: skippedLocked,
        skippedMissing: skippedMissing,
        skippedContext: skippedContext,
        message: skippedLocked ? "可选对象为空：其余对象均已锁定。" : "当前分类没有可选对象。"
      };
    }
    if (!objectFilterSelectShapes(context.container, candidates)) {
      return { ok: false, count: 0, skippedLocked: skippedLocked, skippedMissing: skippedMissing, skippedContext: skippedContext, message: "宿主没有提供可写入的对象选区。" };
    }
    const verified = objectFilterSelectionMatches(candidates);
    if (verified === false) {
      return { ok: false, count: 0, skippedLocked: skippedLocked, skippedMissing: skippedMissing, skippedContext: skippedContext, message: "宿主未确认目标对象已全部选中。" };
    }
    const skipped = skippedLocked + skippedMissing + skippedContext;
    return {
      ok: true,
      count: candidates.length,
      skippedLocked: skippedLocked,
      skippedMissing: skippedMissing,
      skippedContext: skippedContext,
      message: "已选择 " + candidates.length + " 个对象" + (skipped ? "，跳过 " + skipped + " 个不可选对象。" : "。")
    };
  }

  function layerSelectAll(items) {
    return layerSelectMany(items);
  }

  // =====================================================================
  // Design productivity tools (WPS)
  // =====================================================================
  // These commands deliberately stay shape-model based: they reuse the same
  // selection and lock rules as object management, return concrete counts,
  // and never launch a destructive bulk action without an explicit command.
  const DESIGN_MAX_SHAPES = 500;
  const designStyleState = { shape: null, snapshot: null, at: 0 };
  const designPhotoshopJobs = [];

  function designReadNumber(object, names, fallback) {
    for (let i = 0; i < names.length; i += 1) {
      try {
        const value = Number(object[names[i]]);
        if (isFinite(value)) return value;
      } catch (_) {}
    }
    return fallback;
  }

  function designReadBool(object, names, fallback) {
    for (let i = 0; i < names.length; i += 1) {
      try {
        const value = object[names[i]];
        if (value !== undefined && value !== null) return isTrue(value);
      } catch (_) {}
    }
    return fallback;
  }

  function designSelectedShapes(required, kindLabel) {
    const app = application();
    const windowObject = app.ActiveWindow;
    if (!windowObject || !windowObject.Selection) throw new Error("请先在正文编辑窗口中选择对象。");
    let range = null;
    try { range = windowObject.Selection.ShapeRange; } catch (_) { range = null; }
    let count = 0;
    try { count = Number(range && range.Count) || 0; } catch (_) { count = 0; }
    if (required && count < required) throw new Error(kindLabel || "请先选择所需对象。");
    const shapes = [];
    for (let i = 1; i <= count; i += 1) {
      let shape = null;
      try { shape = range.Item(i); } catch (_) { shape = null; }
      if (shape) shapes.push(shape);
    }
    return shapes;
  }

  function designCurrentSlide() {
    const app = application();
    const windowObject = app.ActiveWindow;
    if (!windowObject) throw new Error("请先打开演示文稿窗口。");
    try {
      const viewSlide = windowObject.View && windowObject.View.Slide;
      if (viewSlide && viewSlide.Shapes) return viewSlide;
    } catch (_) {}
    const index = readWindowSlideIndex(windowObject);
    if (index > 0) {
      try { return activePresentation().Slides.Item(index); } catch (_) {}
    }
    throw new Error("无法确定当前幻灯片，请进入普通编辑视图。");
  }

  function designSlides(scope) {
    const presentation = activePresentation();
    const count = Number(presentation.Slides.Count) || 0;
    const all = [];
    for (let i = 1; i <= count; i += 1) {
      try {
        const slide = presentation.Slides.Item(i);
        if (slide && slide.Shapes) all.push(slide);
      } catch (_) {}
    }
    if (scope === "all") return all;
    if (scope === "selected") {
      const app = application();
      const selection = app.ActiveWindow && app.ActiveWindow.Selection;
      let type = -1;
      try { type = Number(selection.Type); } catch (_) {}
      if (type === 1) {
        let range = null;
        try { range = selection.SlideRange; } catch (_) { range = null; }
        let slideCount = 0;
        try { slideCount = Number(range && range.Count) || 0; } catch (_) {}
        const selected = [];
        for (let i = 1; i <= slideCount; i += 1) {
          try {
            const slide = range.Item(i);
            if (slide && slide.Shapes) selected.push(slide);
          } catch (_) {}
        }
        if (selected.length) return selected;
      }
      return [designCurrentSlide()];
    }
    return [designCurrentSlide()];
  }

  function designShapeName(shape, index) {
    try {
      const value = String(shape.Name || shape.name || "").trim();
      if (value) return value;
    } catch (_) {}
    return "对象 " + String(index || "");
  }

  function designShapeHasText(shape) {
    try { return isTrue(shape.TextFrame2.HasText) || isTrue(shape.TextFrame.HasText); } catch (_) {}
    try { return String(shape.TextFrame.TextRange.Text || "").length > 0; } catch (_) {}
    return false;
  }

  function designPushShapes(container, output, depth) {
    if (!container || depth > 4) return;
    let count = 0;
    try { count = Number(container.Count) || 0; } catch (_) {}
    for (let i = 1; i <= count && output.length < DESIGN_MAX_SHAPES; i += 1) {
      let shape = null;
      try { shape = container.Item(i); } catch (_) { continue; }
      if (!shape) continue;
      output.push(shape);
      try {
        const type = Number(shape.Type);
        if (type === 6 && shape.GroupItems) designPushShapes(shape.GroupItems, output, depth + 1);
      } catch (_) {}
    }
  }

  function designShapeCollection(scope) {
    const slides = designSlides(scope);
    const rows = [];
    slides.forEach(function (slide) {
      const shapes = [];
      designPushShapes(slide.Shapes, shapes, 0);
      rows.push({ slide: slide, shapes: shapes });
    });
    return rows;
  }

  function designWritePath(object, path, value) {
    let owner = object;
    for (let i = 0; i < path.length - 1; i += 1) {
      try { owner = owner[path[i]]; } catch (_) { return false; }
      if (!owner) return false;
    }
    try { owner[path[path.length - 1]] = value; return true; } catch (_) { return false; }
  }

  function designReadPath(object, path) {
    let value = object;
    for (let i = 0; i < path.length; i += 1) {
      try { value = value[path[i]]; } catch (_) { return undefined; }
    }
    return value;
  }

  function designCaptureEffect(shape, name, keys) {
    let owner = null;
    try { owner = shape[name]; } catch (_) { return null; }
    if (!owner) return null;
    const result = { _available: true };
    keys.forEach(function (key) {
      const value = designReadPath(owner, key);
      if (value !== undefined && value !== null) result[key.join(".")] = value;
    });
    return result;
  }

  function designApplyEffect(shape, name, snapshot) {
    if (!snapshot || !snapshot._available) return false;
    let owner = null;
    try { owner = shape[name]; } catch (_) { return false; }
    if (!owner) return false;
    let changed = false;
    Object.keys(snapshot).forEach(function (key) {
      if (key === "_available") return;
      if (designWritePath(owner, key.split("."), snapshot[key])) changed = true;
    });
    return changed;
  }

  function designCaptureTextStyle(shape) {
    const paths = [
      ["Name"], ["Size"], ["Bold"], ["Italic"], ["Underline"],
      ["Fill", "ForeColor", "RGB"]
    ];
    const result = {};
    let available = false;
    paths.forEach(function (path) {
      ["TextFrame2.TextRange.Font", "TextFrame.TextRange.Font"].forEach(function (ownerPath) {
        const value = designReadPath(shape, ownerPath.split(".").concat(path));
        if (value !== undefined && value !== null) {
          result[path.join(".")] = value;
          available = true;
        }
      });
    });
    result._available = available;
    return result;
  }

  function designApplyTextStyle(shape, snapshot) {
    if (!snapshot || !snapshot._available) return false;
    let changed = false;
    Object.keys(snapshot).forEach(function (key) {
      if (key === "_available") return;
      const path = key.split(".");
      if (designWritePath(shape, ["TextFrame2", "TextRange", "Font"].concat(path), snapshot[key])) changed = true;
      else if (designWritePath(shape, ["TextFrame", "TextRange", "Font"].concat(path), snapshot[key])) changed = true;
    });
    return changed;
  }

  function designStyleCapture() {
    const shapes = designSelectedShapes(1, "请先选择一个样式来源对象。");
    const source = shapes[0];
    if (layerGuardShapeLocked(source)) throw new Error("样式来源对象已锁定。");
    const snapshot = {
      geometry: {
        left: designReadNumber(source, ["Left", "left"]),
        top: designReadNumber(source, ["Top", "top"]),
        width: designReadNumber(source, ["Width", "width"]),
        height: designReadNumber(source, ["Height", "height"]),
        rotation: designReadNumber(source, ["Rotation", "rotation"], 0)
      },
      fill: {
        visible: designReadBool(source.Fill || {}, ["Visible", "visible"], true),
        foreColor: designReadPath(source, ["Fill", "ForeColor", "RGB"]),
        transparency: designReadNumber(source.Fill || {}, ["Transparency", "transparency"])
      },
      line: {
        visible: designReadBool(source.Line || {}, ["Visible", "visible"]),
        weight: designReadNumber(source.Line || {}, ["Weight", "weight"]),
        foreColor: designReadPath(source, ["Line", "ForeColor", "RGB"]),
        dashStyle: designReadNumber(source.Line || {}, ["DashStyle", "dashStyle"])
      },
      text: designCaptureTextStyle(source),
      shadow: designCaptureEffect(source, "Shadow", [["Visible"], ["Transparency"], ["Blur"], ["OffsetX"], ["OffsetY"], ["Style"]]),
      reflection: designCaptureEffect(source, "Reflection", [["Visible"], ["Offset"], ["Blur"], ["Transparency"]]),
      glow: designCaptureEffect(source, "Glow", [["Visible"], ["Radius"], ["Color", "RGB"]]),
      softEdge: designCaptureEffect(source, "SoftEdge", [["Visible"], ["Radius"]]),
      threeD: designCaptureEffect(source, "ThreeD", [["Visible"], ["Depth"], ["BevelTopType"], ["BevelBottomType"]])
    };
    designStyleState.shape = source;
    designStyleState.snapshot = snapshot;
    designStyleState.at = Date.now();
    return {
      ok: true,
      source: designShapeName(source, 1),
      type: designReadNumber(source, ["Type", "type"], 0),
      capturedAt: new Date(designStyleState.at).toLocaleTimeString()
    };
  }

  function designStyleInfo() {
    if (!designStyleState.snapshot) return { ready: false };
    return {
      ready: true,
      source: designShapeName(designStyleState.shape, "来源"),
      capturedAt: new Date(designStyleState.at).toLocaleTimeString()
    };
  }

  function designStyleApply(options) {
    const snapshot = designStyleState.snapshot;
    if (!snapshot) throw new Error("请先点击“拾取样式来源”。");
    const targets = designSelectedShapes(1, "请先选择要应用样式的目标对象。");
    let applied = 0;
    let skippedLocked = 0;
    let propertyCount = 0;
    Object.keys(options || {}).forEach(function (key) { if (options[key]) propertyCount += 1; });
    targets.forEach(function (shape, index) {
      if (layerGuardShapeLocked(shape)) { skippedLocked += 1; return; }
      let changed = false;
      if (options.geometry) {
        ["left", "top", "width", "height", "rotation"].forEach(function (key) {
          const prop = key.charAt(0).toUpperCase() + key.slice(1);
          if (isFinite(snapshot.geometry[key]) && designWritePath(shape, [prop], snapshot.geometry[key])) changed = true;
        });
      }
      if (options.fill) {
        if (snapshot.fill.visible !== undefined && designWritePath(shape, ["Fill", "Visible"], snapshot.fill.visible ? MsoTrue : MsoFalse)) changed = true;
        if (snapshot.fill.foreColor !== undefined && designWritePath(shape, ["Fill", "ForeColor", "RGB"], snapshot.fill.foreColor)) changed = true;
        if (isFinite(snapshot.fill.transparency) && designWritePath(shape, ["Fill", "Transparency"], snapshot.fill.transparency)) changed = true;
      }
      if (options.line) {
        if (snapshot.line.visible !== undefined && designWritePath(shape, ["Line", "Visible"], snapshot.line.visible ? MsoTrue : MsoFalse)) changed = true;
        if (isFinite(snapshot.line.weight) && designWritePath(shape, ["Line", "Weight"], snapshot.line.weight)) changed = true;
        if (snapshot.line.foreColor !== undefined && designWritePath(shape, ["Line", "ForeColor", "RGB"], snapshot.line.foreColor)) changed = true;
        if (isFinite(snapshot.line.dashStyle) && designWritePath(shape, ["Line", "DashStyle"], snapshot.line.dashStyle)) changed = true;
      }
      if (options.text && designApplyTextStyle(shape, snapshot.text)) changed = true;
      if (options.shadow && designApplyEffect(shape, "Shadow", snapshot.shadow)) changed = true;
      if (options.reflection && designApplyEffect(shape, "Reflection", snapshot.reflection)) changed = true;
      if (options.glow && designApplyEffect(shape, "Glow", snapshot.glow)) changed = true;
      if (options.softEdge && designApplyEffect(shape, "SoftEdge", snapshot.softEdge)) changed = true;
      if (options.threeD && designApplyEffect(shape, "ThreeD", snapshot.threeD)) changed = true;
      if (changed) applied += 1;
      void index;
    });
    invalidatePanelInventoryCache();
    if (!applied && skippedLocked === targets.length) throw new Error("目标对象全部处于锁定状态。");
    return {
      ok: true,
      total: targets.length,
      applied: applied,
      skippedLocked: skippedLocked,
      message: "已应用 " + applied + " 个对象" + (skippedLocked ? "，跳过 " + skippedLocked + " 个锁定对象。" : "。")
    };
  }

  function designCountOccurrences(text, needle, matchCase) {
    if (!needle) return 0;
    const hay = String(text == null ? "" : matchCase ? text : String(text).toLowerCase());
    const part = matchCase ? needle : String(needle).toLowerCase();
    let count = 0;
    let offset = hay.indexOf(part);
    while (offset >= 0) {
      count += 1;
      offset = hay.indexOf(part, offset + part.length);
    }
    return count;
  }

  function designReplaceTextRange(range, findText, replaceText, options) {
    let before = "";
    try { before = String(range.Text || range.text || ""); } catch (_) { before = ""; }
    const expected = designCountOccurrences(before, findText, options.matchCase);
    if (!expected) return { count: 0, changed: false, manual: false };
    if (hasMethod(range, "Replace")) {
      const attempts = [
        function () { return range.Replace(findText, replaceText); },
        function () { return range.Replace(findText, replaceText, options.wholeWord ? MsoTrue : MsoFalse); },
        function () { return range.Replace(findText, replaceText, options.wholeWord ? MsoTrue : MsoFalse, options.matchCase ? MsoTrue : MsoFalse); }
      ];
      for (let i = 0; i < attempts.length; i += 1) {
        try {
          attempts[i]();
          const after = String(range.Text || range.text || "");
          const remaining = designCountOccurrences(after, findText, options.matchCase);
          if (remaining < expected) return { count: expected - remaining, changed: true, manual: false };
        } catch (_) {}
      }
    }
    const replaced = options.matchCase
      ? before.split(findText).join(replaceText)
      : before.replace(new RegExp(findText.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi"), replaceText);
    if (replaced === before) return { count: 0, changed: false, manual: true };
    try { range.Text = replaced; return { count: expected, changed: true, manual: true }; }
    catch (_) { return { count: 0, changed: false, manual: true }; }
  }

  function designTextFindReplace(findText, replaceText, options) {
    const find = String(findText == null ? "" : findText);
    if (!find) throw new Error("请输入要查找的文字。");
    const opts = Object.assign({ matchCase: false, wholeWord: false, scope: "all" }, options || {});
    const groups = designShapeCollection(opts.scope);
    let shapeCount = 0;
    let occurrenceCount = 0;
    let manualCount = 0;
    let skippedLocked = 0;
    groups.forEach(function (group) {
      group.shapes.forEach(function (shape) {
        if (!designShapeHasText(shape)) return;
        if (layerGuardShapeLocked(shape)) { skippedLocked += 1; return; }
        const result = designReplaceTextRange(shape.TextFrame2 && shape.TextFrame2.TextRange ? shape.TextFrame2.TextRange : shape.TextFrame.TextRange, find, String(replaceText == null ? "" : replaceText), opts);
        if (result.changed) {
          shapeCount += 1;
          occurrenceCount += result.count;
          if (result.manual) manualCount += 1;
        }
      });
    });
    if (occurrenceCount) invalidatePanelInventoryCache();
    return {
      ok: true,
      shapes: shapeCount,
      occurrences: occurrenceCount,
      manual: manualCount,
      skippedLocked: skippedLocked,
      message: occurrenceCount
        ? "已替换 " + occurrenceCount + " 处文字，涉及 " + shapeCount + " 个对象。" + (manualCount ? " 其中部分对象使用了整段回写。" : "")
        : "没有找到匹配文字。"
    };
  }

  function designTextSwap() {
    const shapes = designSelectedShapes(2, "请选择两个文字对象。");
    const a = shapes[0];
    const b = shapes[1];
    if (layerGuardShapeLocked(a) || layerGuardShapeLocked(b)) throw new Error("选中的文字对象包含锁定状态。");
    if (!designShapeHasText(a) || !designShapeHasText(b)) throw new Error("请选择两个包含文字的对象。");
    const rangeA = a.TextFrame2 && a.TextFrame2.TextRange ? a.TextFrame2.TextRange : a.TextFrame.TextRange;
    const rangeB = b.TextFrame2 && b.TextFrame2.TextRange ? b.TextFrame2.TextRange : b.TextFrame.TextRange;
    const textA = String(rangeA.Text || "");
    const textB = String(rangeB.Text || "");
    rangeA.Text = textB;
    rangeB.Text = textA;
    invalidatePanelInventoryCache();
    return { ok: true, message: "已交换两个文字对象的内容。注意：整段回写可能保留对象级格式。", shapes: 2 };
  }

  function designNotesText(slide) {
    const values = [];
    try {
      const shapes = slide.NotesPage && slide.NotesPage.Shapes;
      let count = 0;
      try { count = Number(shapes.Count) || 0; } catch (_) {}
      for (let i = 1; i <= count; i += 1) {
        try {
          const shape = shapes.Item(i);
          if (isTrue(shape.HasTextFrame) && isTrue(shape.TextFrame.HasText)) values.push(String(shape.TextFrame.TextRange.Text || ""));
        } catch (_) {}
      }
    } catch (_) {}
    return values.join("\n");
  }

  function designTextExtract(scope, includeNotes) {
    const groups = designShapeCollection(scope || "all");
    const items = [];
    groups.forEach(function (group, groupIndex) {
      let slideIndex = groupIndex + 1;
      try { slideIndex = Number(group.slide.SlideIndex || group.slide.slideIndex) || groupIndex + 1; } catch (_) {}
      group.shapes.forEach(function (shape) {
        if (!designShapeHasText(shape)) return;
        const range = shape.TextFrame2 && shape.TextFrame2.TextRange ? shape.TextFrame2.TextRange : shape.TextFrame.TextRange;
        const text = String(range.Text || range.text || "");
        if (!text.trim()) return;
        items.push({ slide: slideIndex, name: designShapeName(shape, items.length + 1), text: text });
      });
      if (includeNotes) {
        const note = designNotesText(group.slide);
        if (note.trim()) items.push({ slide: slideIndex, name: "备注", text: note, note: true });
      }
    });
    return { ok: true, items: items, text: items.map(function (item) { return "P" + item.slide + "\t" + item.name + "\t" + item.text.replace(/\r?\n/g, "\\n"); }).join("\n") };
  }

  function designSelectionGeometry(shapes) {
    let left = Infinity;
    let top = Infinity;
    let right = -Infinity;
    let bottom = -Infinity;
    shapes.forEach(function (shape) {
      left = Math.min(left, designReadNumber(shape, ["Left", "left"], 0));
      top = Math.min(top, designReadNumber(shape, ["Top", "top"], 0));
      right = Math.max(right, designReadNumber(shape, ["Left", "left"], 0) + designReadNumber(shape, ["Width", "width"], 0));
      bottom = Math.max(bottom, designReadNumber(shape, ["Top", "top"], 0) + designReadNumber(shape, ["Height", "height"], 0));
    });
    return { left: left, top: top, right: right, bottom: bottom };
  }

  function designAlignRun(mode, options) {
    const shapes = designSelectedShapes(2, "请至少选择两个对象。").filter(function (shape) { return !layerGuardShapeLocked(shape); });
    if (shapes.length < 2) throw new Error("没有两个可编辑的目标对象。");
    const first = shapes[0];
    const page = activePresentation().PageSetup;
    const pageWidth = designReadNumber(page, ["SlideWidth", "slideWidth"], 720);
    const pageHeight = designReadNumber(page, ["SlideHeight", "slideHeight"], 540);
    let changed = 0;
    const setPos = function (shape, left, top) {
      if (designWritePath(shape, ["Left"], left) | designWritePath(shape, ["Top"], top)) changed += 1;
    };
    if (mode === "align-left") shapes.forEach(function (shape) { setPos(shape, first.Left, shape.Top); });
    else if (mode === "align-hcenter") shapes.forEach(function (shape) { setPos(shape, (pageWidth - shape.Width) / 2, shape.Top); });
    else if (mode === "align-right") shapes.forEach(function (shape) { setPos(shape, pageWidth - shape.Width, shape.Top); });
    else if (mode === "align-top") shapes.forEach(function (shape) { setPos(shape, shape.Left, first.Top); });
    else if (mode === "align-vcenter") shapes.forEach(function (shape) { setPos(shape, shape.Left, (pageHeight - shape.Height) / 2); });
    else if (mode === "align-bottom") shapes.forEach(function (shape) { setPos(shape, shape.Left, pageHeight - shape.Height); });
    else if (mode === "align-page-center") shapes.forEach(function (shape) { setPos(shape, (pageWidth - shape.Width) / 2, (pageHeight - shape.Height) / 2); });
    else if (mode === "swap-position") {
      if (shapes.length !== 2) throw new Error("交换位置请只选择两个对象。");
      const leftA = first.Left, topA = first.Top;
      first.Left = shapes[1].Left; first.Top = shapes[1].Top;
      shapes[1].Left = leftA; shapes[1].Top = topA;
      changed = 2;
    } else if (mode === "distribute-h" || mode === "distribute-v") {
      const horizontal = mode === "distribute-h";
      const ordered = shapes.slice().sort(function (a, b) { return horizontal ? a.Left - b.Left : a.Top - b.Top; });
      const firstValue = horizontal ? ordered[0].Left : ordered[0].Top;
      const lastValue = horizontal ? ordered[ordered.length - 1].Left : ordered[ordered.length - 1].Top;
      const lastSize = horizontal ? ordered[ordered.length - 1].Width : ordered[ordered.length - 1].Height;
      const gap = ordered.length > 1 ? (lastValue - firstValue - lastSize) / (ordered.length - 1) : 0;
      let cursor = firstValue;
      ordered.forEach(function (shape, index) {
        if (index === 0 || index === ordered.length - 1) {
          cursor = horizontal ? shape.Left + shape.Width + gap : shape.Top + shape.Height + gap;
          return;
        }
        if (horizontal) { shape.Left = cursor; cursor += shape.Width + gap; }
        else { shape.Top = cursor; cursor += shape.Height + gap; }
        changed += 1;
      });
    } else if (mode === "matrix") {
      const rows = Math.max(1, parseInt(options.rows, 10) || 2);
      const columns = Math.max(1, parseInt(options.columns, 10) || 2);
      if (rows * columns < shapes.length) throw new Error("行列数量不足容纳当前对象。");
      let maxWidth = 0;
      let maxHeight = 0;
      shapes.forEach(function (shape) { maxWidth = Math.max(maxWidth, shape.Width); maxHeight = Math.max(maxHeight, shape.Height); });
      const originX = isFinite(options.left) ? options.left : first.Left;
      const originY = isFinite(options.top) ? options.top : first.Top;
      shapes.forEach(function (shape, index) {
        const row = Math.floor(index / columns);
        const column = index % columns;
        setPos(shape, originX + column * (maxWidth + Number(options.gapX || 0)), originY + row * (maxHeight + Number(options.gapY || 0)));
      });
    } else if (mode === "ring") {
      const radius = Math.max(1, Number(options.radius) || 150);
      const bounds = designSelectionGeometry(shapes);
      const centerX = isFinite(options.centerX) ? options.centerX : (bounds.left + bounds.right) / 2;
      const centerY = isFinite(options.centerY) ? options.centerY : (bounds.top + bounds.bottom) / 2;
      const start = (Number(options.startAngle) || 0) * Math.PI / 180;
      shapes.forEach(function (shape, index) {
        const angle = start + index * 2 * Math.PI / shapes.length;
        setPos(shape, centerX + Math.cos(angle) * radius - shape.Width / 2, centerY + Math.sin(angle) * radius - shape.Height / 2);
      });
    } else if (mode === "uniform-size" || mode === "uniform-width" || mode === "uniform-height" || mode === "uniform-aspect") {
      const source = shapes[shapes.length - 1];
      shapes.forEach(function (shape) {
        if (shape === source) return;
        if (mode !== "uniform-height") shape.Width = source.Width;
        if (mode !== "uniform-width") shape.Height = mode === "uniform-aspect" ? source.Width * shape.Height / Math.max(0.01, shape.Width) : source.Height;
        changed += 1;
      });
    } else if (mode === "uniform-angle") {
      shapes.forEach(function (shape) { if (designWritePath(shape, ["Rotation"], first.Rotation)) changed += 1; });
    } else {
      throw new Error("不支持的对齐模式。");
    }
    return { ok: true, mode: mode, changed: changed, total: shapes.length, message: "已完成：" + mode + "，处理 " + changed + " 个对象。" };
  }

  function designIsShapeVisible(shape) {
    return designReadBool(shape, ["Visible", "visible"], true);
  }

  function designSlideHasVisibleContent(slide) {
    let count = 0;
    try { count = Number(slide.Shapes.Count) || 0; } catch (_) {}
    for (let i = 1; i <= count; i += 1) {
      try { if (designIsShapeVisible(slide.Shapes.Item(i))) return true; } catch (_) {}
    }
    return false;
  }

  function designClearNotes(slide) {
    let changed = 0;
    try {
      const shapes = slide.NotesPage.Shapes;
      const count = Number(shapes.Count) || 0;
      for (let i = 1; i <= count; i += 1) {
        const shape = shapes.Item(i);
        let hasText = false;
        try { hasText = isTrue(shape.HasTextFrame) && isTrue(shape.TextFrame.HasText); } catch (_) {}
        if (hasText && String(shape.TextFrame.TextRange.Text || "").trim()) {
          shape.TextFrame.TextRange.Text = "";
          changed += 1;
        }
      }
    } catch (_) {}
    return changed;
  }

  function designDeleteAnimations(slide) {
    let deleted = 0;
    try {
      const sequence = slide.TimeLine.MainSequence;
      let count = Number(sequence.Count) || 0;
      for (let i = count; i >= 1; i -= 1) {
        try { sequence.Item(i).Delete(); deleted += 1; } catch (_) {}
        count = Number(sequence.Count) || 0;
      }
      return deleted;
    } catch (_) { return -1; }
  }

  function designCleanup(kind, scope) {
    const slides = designSlides(scope || "all");
    let changed = 0;
    let examined = 0;
    let unsupported = 0;
    for (let slideIndex = slides.length - 1; slideIndex >= 0; slideIndex -= 1) {
      const slide = slides[slideIndex];
      examined += 1;
      if (kind === "blank-slides") {
        if (!designSlideHasVisibleContent(slide)) {
          try { slide.Delete(); changed += 1; } catch (_) {}
        }
      } else if (kind === "notes") {
        changed += designClearNotes(slide);
      } else if (kind === "animations") {
        const result = designDeleteAnimations(slide);
        if (result < 0) unsupported += 1;
        else changed += result;
      } else if (kind === "hidden-shapes" || kind === "outside-shapes") {
        const page = activePresentation().PageSetup;
        const pageWidth = designReadNumber(page, ["SlideWidth", "slideWidth"], 720);
        const pageHeight = designReadNumber(page, ["SlideHeight", "slideHeight"], 540);
        let count = 0;
        try { count = Number(slide.Shapes.Count) || 0; } catch (_) {}
        for (let i = count; i >= 1; i -= 1) {
          try {
            const shape = slide.Shapes.Item(i);
            if (layerGuardShapeLocked(shape)) continue;
            const left = designReadNumber(shape, ["Left", "left"], 0);
            const top = designReadNumber(shape, ["Top", "top"], 0);
            const width = designReadNumber(shape, ["Width", "width"], 0);
            const height = designReadNumber(shape, ["Height", "height"], 0);
            const remove = kind === "hidden-shapes"
              ? !designIsShapeVisible(shape)
              : (left + width < -1 || top + height < -1 || left > pageWidth + 1 || top > pageHeight + 1);
            if (remove) { shape.Delete(); changed += 1; }
          } catch (_) {}
        }
      }
    }
    invalidatePanelInventoryCache();
    const labels = {
      "blank-slides": "空白页",
      "notes": "备注",
      "animations": "动画",
      "hidden-shapes": "隐藏对象",
      "outside-shapes": "画外对象"
    };
    if (kind !== "blank-slides" && changed === 0 && unsupported) throw new Error("当前 WPS 未开放动画时间线 API。");
    return { ok: true, kind: kind, examined: examined, changed: changed, unsupported: unsupported, message: "检查 " + examined + " 页，处理 " + (labels[kind] || kind) + " " + changed + " 项。" };
  }

  function designChooseFolder(title) {
    const app = application();
    if (!hasMethod(app, "FileDialog")) throw new Error("当前 WPS 版本没有提供系统文件夹选择器。");
    const dialog = app.FileDialog(4);
    dialog.Title = title || "选择导出文件夹";
    dialog.AllowMultiSelect = false;
    if (Number(dialog.Show()) !== MsoTrue) return "";
    if (!dialog.SelectedItems || Number(dialog.SelectedItems.Count) < 1) return "";
    return String(dialog.SelectedItems.Item(1));
  }

  function designSafeFileName(value) {
    return String(value || "").replace(/[\\/:*?"<>|]+/g, "_").replace(/\s+/g, " ").trim().slice(0, 80);
  }

  async function designExportSlides(scope, format, dpi) {
    const folder = designChooseFolder("选择图片导出文件夹");
    if (!folder) return { ok: false, cancelled: true, message: "已取消导出。" };
    const cleanFolder = folder.replace(/[\\/]+$/, "");
    const extension = String(format || "png").toLowerCase() === "jpg" ? "jpg" : "png";
    const hostFormat = extension === "jpg" ? "JPG" : "PNG";
    const scale = Math.max(36, Math.min(1200, Number(dpi) || 144)) / 72;
    const page = activePresentation().PageSetup;
    const width = Math.max(32, Math.round(designReadNumber(page, ["SlideWidth", "slideWidth"], 720) * scale));
    const height = Math.max(32, Math.round(designReadNumber(page, ["SlideHeight", "slideHeight"], 540) * scale));
    const slides = designSlides(scope || "all");
    const files = [];
    let failed = 0;
    for (let i = 0; i < slides.length; i += 1) {
      const slide = slides[i];
      let index = i + 1;
      try { index = Number(slide.SlideIndex || slide.slideIndex) || i + 1; } catch (_) {}
      const path = cleanFolder + "\\" + designSafeFileName(activePresentation().FullName || "Slides") + "_P" + String(index).padStart(3, "0") + "." + extension;
      try {
        slide.Export(path, hostFormat, width, height);
        if (fileExists(path)) files.push(path);
        else failed += 1;
      } catch (_) { failed += 1; }
      await yieldUI(15);
    }
    if (!files.length) throw new Error("当前 WPS 未成功导出幻灯片（Slide.Export 不可用或被拒绝）。");
    return { ok: true, count: files.length, failed: failed, folder: cleanFolder, files: files, message: "已导出 " + files.length + " 页到 " + cleanFolder + (failed ? "，失败 " + failed + " 页。" : "。") };
  }

  async function designLayerStamp() {
    const slide = designCurrentSlide();
    const path = tempPath("LayerStamp");
    const page = activePresentation().PageSetup;
    const width = designReadNumber(page, ["SlideWidth", "slideWidth"], 720);
    const height = designReadNumber(page, ["SlideHeight", "slideHeight"], 540);
    slide.Export(path, "PNG", Math.round(width * 2), Math.round(height * 2));
    if (!fileExists(path)) throw new Error("当前 WPS 未提供 Slide.Export，无法盖印图层。");
    const stamp = slide.Shapes.AddPicture(path, MsoFalse, MsoTrue, 0, 0, width, height);
    removeFile(path);
    invalidatePanelInventoryCache();
    return { ok: true, name: designShapeName(stamp, "图层盖印"), message: "已把当前页盖印为整页图片。" };
  }

  async function designExtractSlides(scope) {
    const slides = designSlides(scope || "selected");
    if (!slides.length) throw new Error("请先选择要提取的幻灯片。");
    const app = application();
    const target = app.Presentations.Add();
    if (!target || !target.Slides) throw new Error("当前 WPS 未支持新建演示文稿。");
    const source = activePresentation();
    const indexes = slides.map(function (slide) { return Number(slide.SlideIndex || slide.slideIndex) || 0; }).filter(function (value) { return value > 0; });
    let copied = false;
    try {
      source.Slides.Range(indexes).Copy();
      target.Slides.Paste();
      copied = Number(target.Slides.Count) > 0;
    } catch (_) { copied = false; }
    if (!copied) throw new Error("当前 WPS 未开放跨文稿幻灯片复制接口。");
    return { ok: true, count: Number(target.Slides.Count) || slides.length, message: "已提取 " + slides.length + " 页到新演示文稿，请手动保存。" };
  }

  function designRgbToHsl(rgb) {
    const value = Math.max(0, Math.min(0xffffff, Number(rgb) >>> 0));
    const r = ((value >> 16) & 255) / 255;
    const g = ((value >> 8) & 255) / 255;
    const b = (value & 255) / 255;
    const max = Math.max(r, g, b);
    const min = Math.min(r, g, b);
    let h = 0;
    let s = 0;
    const l = (max + min) / 2;
    if (max !== min) {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
      else if (max === g) h = ((b - r) / d + 2) / 6;
      else h = ((r - g) / d + 4) / 6;
    }
    return { h: h * 360, s: s * 100, l: l * 100 };
  }

  function designHueToRgb(p, q, t) {
    let x = t;
    if (x < 0) x += 1;
    if (x > 1) x -= 1;
    if (x < 1 / 6) return p + (q - p) * 6 * x;
    if (x < 1 / 2) return q;
    if (x < 2 / 3) return p + (q - p) * (2 / 3 - x) * 6;
    return p;
  }

  function designHslToRgb(h, s, l) {
    const hue = ((Number(h) || 0) % 360 + 360) % 360 / 360;
    const sat = Math.max(0, Math.min(100, Number(s) || 0)) / 100;
    const light = Math.max(0, Math.min(100, Number(l) || 0)) / 100;
    if (sat === 0) {
      const gray = Math.round(light * 255);
      return (gray << 16) | (gray << 8) | gray;
    }
    const q = light < 0.5 ? light * (1 + sat) : light + sat - light * sat;
    const p = 2 * light - q;
    const r = designHueToRgb(p, q, hue + 1 / 3);
    const g = designHueToRgb(p, q, hue);
    const b = designHueToRgb(p, q, hue - 1 / 3);
    return (Math.round(r * 255) << 16) | (Math.round(g * 255) << 8) | Math.round(b * 255);
  }

  function designColorTargets(shape) {
    const targets = [];
    targets.push({ owner: shape, path: ["Fill", "ForeColor", "RGB"] });
    targets.push({ owner: shape, path: ["Line", "ForeColor", "RGB"] });
    targets.push({ owner: shape, path: ["TextFrame2", "TextRange", "Font", "Fill", "ForeColor", "RGB"] });
    targets.push({ owner: shape, path: ["TextFrame", "TextRange", "Font", "Color", "RGB"] });
    return targets;
  }

  function designColorAdjust(hueShift, saturationShift, lightnessShift) {
    const shapes = designSelectedShapes(1, "请先选择要调整颜色的对象。").filter(function (shape) { return !layerGuardShapeLocked(shape); });
    let changed = 0;
    shapes.forEach(function (shape) {
      designColorTargets(shape).forEach(function (target) {
        const current = designReadPath(target.owner, target.path);
        if (current === undefined || current === null || !isFinite(Number(current))) return;
        const hsl = designRgbToHsl(Number(current));
        const next = designHslToRgb(hsl.h + Number(hueShift || 0), hsl.s + Number(saturationShift || 0), hsl.l + Number(lightnessShift || 0));
        if (designWritePath(target.owner, target.path, next)) changed += 1;
      });
    });
    return { ok: true, changed: changed, message: "颜色属性已调整 " + changed + " 处。" };
  }

  function designColorReplace(fromHex, toHex, tolerance) {
    const from = parseInt(String(fromHex || "").replace("#", ""), 16);
    const to = parseInt(String(toHex || "").replace("#", ""), 16);
    if (!isFinite(from) || !isFinite(to)) throw new Error("请提供有效的起始色和目标色。");
    const maxDelta = Math.max(0, Math.min(255, Number(tolerance) || 0));
    const shapes = designSelectedShapes(1, "请先选择要替换颜色的对象。").filter(function (shape) { return !layerGuardShapeLocked(shape); });
    const fr = (from >> 16) & 255, fg = (from >> 8) & 255, fb = from & 255;
    let changed = 0;
    shapes.forEach(function (shape) {
      designColorTargets(shape).forEach(function (target) {
        const current = Number(designReadPath(target.owner, target.path));
        if (!isFinite(current)) return;
        const rr = (current >> 16) & 255, gg = (current >> 8) & 255, bb = current & 255;
        if (Math.abs(rr - fr) > maxDelta || Math.abs(gg - fg) > maxDelta || Math.abs(bb - fb) > maxDelta) return;
        if (designWritePath(target.owner, target.path, to)) changed += 1;
      });
    });
    return { ok: true, changed: changed, message: "颜色替换完成，更新 " + changed + " 处。" };
  }

  function designPhotoshopCandidates() {
    const roots = ["C:\\Program Files\\Adobe\\", "C:\\Program Files (x86)\\Adobe\\"];
    const versions = ["2024", "2023", "2022", "2021", "2020", "CC 2019", "CC 2018"];
    const candidates = [];
    roots.forEach(function (root) {
      versions.forEach(function (version) {
        candidates.push(root + "Adobe Photoshop " + version + "\\Photoshop.exe");
      });
      candidates.push(root + "Adobe Photoshop\\Photoshop.exe");
    });
    return candidates;
  }

  function designPhotoshopOpen(explicitPath) {
    const shapes = designSelectedShapes(1, "请先选择一张图片。");
    const shape = shapes[0];
    const type = designReadNumber(shape, ["Type", "type"], 0);
    if (type !== 13 && type !== 11) throw new Error("请选择图片对象。");
    const path = tempPath("PS");
    const width = Math.max(64, Math.round(designReadNumber(shape, ["Width", "width"], 300) * 2));
    const height = Math.max(64, Math.round(designReadNumber(shape, ["Height", "height"], 200) * 2));
    shape.Export(path, "PNG", width, height);
    if (!fileExists(path)) throw new Error("当前 WPS 未提供 Shape.Export，无法导出图片到 Photoshop。");
    let exe = String(explicitPath || "").trim();
    if (!exe) exe = designPhotoshopCandidates().find(fileExists) || "";
    if (!exe) throw new Error("未找到 Photoshop。请在下方填写 Photoshop.exe 完整路径。");
    const launch = shellExecutePath(exe, '"' + path + '"');
    if (!launch || !launch.ok) throw new Error(launch && launch.error ? launch.error : "无法启动 Photoshop。");
    designPhotoshopJobs.unshift({ shape: shape, path: path, at: Date.now() });
    designPhotoshopJobs.length = Math.min(designPhotoshopJobs.length, 20);
    return { ok: true, path: path, exe: exe, message: "已导出并提交给 Photoshop。编辑后保存到原路径，再点击“载回图片”。" };
  }

  function designPhotoshopReload() {
    const job = designPhotoshopJobs[0];
    if (!job) throw new Error("没有可载回的 Photoshop 编辑任务。");
    if (!fileExists(job.path)) throw new Error("编辑文件不存在，请重新导出。");
    replacePictureKeepCrop(job.shape, job.path);
    invalidatePanelInventoryCache();
    return { ok: true, message: "已把 Photoshop 编辑结果原位载回，并保留裁剪和几何状态。" };
  }


  // =====================================================================
  // Smart Zoom (WPS)
  // =====================================================================
  // The reference add-in applies every new value to the state captured when
  // its dialog opens.  Keep the same contract here: slider changes never
  // compound rounding errors, and closing the task pane never writes metadata
  // into the user's shapes.
  const SMART_ZOOM_MIN_PERCENT = 1;
  const SMART_ZOOM_MAX_PERCENT = 300;
  const SMART_ZOOM_PT_PER_CM = 28.3464567;
  const SMART_ZOOM_MIN_TEXT_SIZE_PT = 8;
  // Table cells are one typographic unit: clamping each cell independently
  // destroys the hierarchy of header/body/footnote sizes at small factors.
  // Keep a common scale for the whole table and use a lower practical floor.
  const SMART_ZOOM_MIN_TABLE_TEXT_SIZE_PT = 4;
  const SMART_ZOOM_MAX_OBJECTS = 500;
  const SMART_ZOOM_MAX_GROUP_DEPTH = 64;
  const SMART_ZOOM_GROUP_TYPE = 6;

  let smartZoomSession = null;
  let smartZoomSessionSeq = 0;

  const SMART_ZOOM_STYLE_SPECS = [
    { key: "shapeLine", option: "scaleShapeLine", paths: [["Line", "Weight"]], visibilityPaths: [["Line", "Visible"]] },
    { key: "shapeShadowBlur", option: "scaleShapeShadow", paths: [["Shadow", "Blur"]], visibilityPaths: [["Shadow", "Visible"]] },
    { key: "shapeShadowOffsetX", option: "scaleShapeShadow", paths: [["Shadow", "OffsetX"]], visibilityPaths: [["Shadow", "Visible"]] },
    { key: "shapeShadowOffsetY", option: "scaleShapeShadow", paths: [["Shadow", "OffsetY"]], visibilityPaths: [["Shadow", "Visible"]] },
    { key: "shapeReflection", option: "scaleShapeReflection", paths: [["Reflection", "Offset"]], visibilityPaths: [["Reflection", "Visible"]] },
    { key: "shapeGlow", option: "scaleShapeGlow", paths: [["Glow", "Radius"]], visibilityPaths: [["Glow", "Visible"]] },
    { key: "softEdge", option: "scaleSoftEdge", paths: [["SoftEdge", "Radius"]], disabledWhenZero: true },
    { key: "shapeBevelTopDepth", option: "scaleShapeThreeD", paths: [["ThreeD", "BevelTopDepth"]], visibilityPaths: [["ThreeD", "Visible"]] },
    { key: "shapeBevelTopInset", option: "scaleShapeThreeD", paths: [["ThreeD", "BevelTopInset"]], visibilityPaths: [["ThreeD", "Visible"]] },
    { key: "shapeBevelBottomDepth", option: "scaleShapeThreeD", paths: [["ThreeD", "BevelBottomDepth"]], visibilityPaths: [["ThreeD", "Visible"]] },
    { key: "shapeBevelBottomInset", option: "scaleShapeThreeD", paths: [["ThreeD", "BevelBottomInset"]], visibilityPaths: [["ThreeD", "Visible"]] },
    { key: "shapeDepth", option: "scaleShapeThreeD", paths: [["ThreeD", "Depth"]], visibilityPaths: [["ThreeD", "Visible"]] },
    { key: "shapeContourWidth", option: "scaleShapeThreeD", paths: [["ThreeD", "ContourWidth"]], visibilityPaths: [["ThreeD", "Visible"]] },
    { key: "shapeZ", option: "scaleShapeThreeD", paths: [["ThreeD", "Z"]], visibilityPaths: [["ThreeD", "Visible"]] },
    { key: "textSize", option: "scaleText", paths: [["TextFrame2", "TextRange", "Font", "Size"], ["TextFrame", "TextRange", "Font", "Size"]] },
    { key: "marginTop", option: "scaleText", paths: [["TextFrame", "MarginTop"]] },
    { key: "marginLeft", option: "scaleText", paths: [["TextFrame", "MarginLeft"]] },
    { key: "marginRight", option: "scaleText", paths: [["TextFrame", "MarginRight"]] },
    { key: "marginBottom", option: "scaleText", paths: [["TextFrame", "MarginBottom"]] },
    { key: "lineRuleAfter", option: "scaleText", paths: [["TextFrame", "TextRange", "ParagraphFormat", "LineRuleAfter"]] },
    { key: "textLine", option: "scaleTextLine", paths: [["TextFrame2", "TextRange", "Font", "Line", "Weight"]], visibilityPaths: [["TextFrame2", "TextRange", "Font", "Line", "Visible"]] },
    { key: "textShadowBlur", option: "scaleTextShadow", paths: [["TextFrame2", "TextRange", "Font", "Shadow", "Blur"]], visibilityPaths: [["TextFrame2", "TextRange", "Font", "Shadow", "Visible"]] },
    { key: "textShadowOffsetX", option: "scaleTextShadow", paths: [["TextFrame2", "TextRange", "Font", "Shadow", "OffsetX"]], visibilityPaths: [["TextFrame2", "TextRange", "Font", "Shadow", "Visible"]] },
    { key: "textShadowOffsetY", option: "scaleTextShadow", paths: [["TextFrame2", "TextRange", "Font", "Shadow", "OffsetY"]], visibilityPaths: [["TextFrame2", "TextRange", "Font", "Shadow", "Visible"]] },
    { key: "textReflection", option: "scaleTextReflection", paths: [["TextFrame2", "TextRange", "Font", "Reflection", "Offset"]], visibilityPaths: [["TextFrame2", "TextRange", "Font", "Reflection", "Visible"]] },
    { key: "textGlow", option: "scaleTextGlow", paths: [["TextFrame2", "TextRange", "Font", "Glow", "Radius"]], visibilityPaths: [["TextFrame2", "TextRange", "Font", "Glow", "Visible"]] },
    { key: "textBevelTopDepth", option: "scaleTextThreeD", paths: [["TextFrame2", "ThreeD", "BevelTopDepth"]], visibilityPaths: [["TextFrame2", "ThreeD", "Visible"]] },
    { key: "textBevelTopInset", option: "scaleTextThreeD", paths: [["TextFrame2", "ThreeD", "BevelTopInset"]], visibilityPaths: [["TextFrame2", "ThreeD", "Visible"]] },
    { key: "textBevelBottomDepth", option: "scaleTextThreeD", paths: [["TextFrame2", "ThreeD", "BevelBottomDepth"]], visibilityPaths: [["TextFrame2", "ThreeD", "Visible"]] },
    { key: "textBevelBottomInset", option: "scaleTextThreeD", paths: [["TextFrame2", "ThreeD", "BevelBottomInset"]], visibilityPaths: [["TextFrame2", "ThreeD", "Visible"]] },
    { key: "textDepth", option: "scaleTextThreeD", paths: [["TextFrame2", "ThreeD", "Depth"]], visibilityPaths: [["TextFrame2", "ThreeD", "Visible"]] },
    { key: "textContourWidth", option: "scaleTextThreeD", paths: [["TextFrame2", "ThreeD", "ContourWidth"]], visibilityPaths: [["TextFrame2", "ThreeD", "Visible"]] },
    { key: "textZ", option: "scaleTextThreeD", paths: [["TextFrame2", "ThreeD", "Z"]], visibilityPaths: [["TextFrame2", "ThreeD", "Visible"]] }
  ];

  function smartZoomReadPath(object, path) {
    let current = object;
    try {
      for (let i = 0; i < path.length; i += 1) {
        if (current === null || current === undefined) return { ok: false };
        current = current[path[i]];
      }
      return { ok: current !== null && current !== undefined, value: current, path: path };
    } catch (_) { return { ok: false }; }
  }

  function smartZoomReadNumber(object, path) {
    const result = smartZoomReadPath(object, path);
    if (!result.ok) return { ok: false };
    const value = Number(result.value);
    return isFinite(value) ? { ok: true, value: value, path: result.path } : { ok: false };
  }

  function smartZoomReadFirstNumber(object, paths) {
    for (let i = 0; i < paths.length; i += 1) {
      const result = smartZoomReadNumber(object, paths[i]);
      if (result.ok) return result;
    }
    return { ok: false };
  }

  function smartZoomReadFirstPositiveNumber(object, paths) {
    for (let i = 0; i < paths.length; i += 1) {
      const result = smartZoomReadNumber(object, paths[i]);
      if (result.ok && result.value > 0) return result;
    }
    return { ok: false };
  }

  function smartZoomReadTextSize(object, paths) {
    for (let i = 0; i < paths.length; i += 1) {
      const result = smartZoomReadNumber(object, paths[i]);
      if (!result.ok) continue;
      // A concrete value from the first exposed text API is authoritative.
      // If that API reports a mixed/empty range (0 or -2), do not fall back to
      // another text API that might expose only a default size; enumerate runs
      // instead so rich text keeps its individual sizes.
      return result.value > 0 ? result : { ok: false };
    }
    return { ok: false };
  }

  function smartZoomReadFirst(object, paths) {
    for (let i = 0; i < paths.length; i += 1) {
      const result = smartZoomReadPath(object, paths[i]);
      if (result.ok) return result;
    }
    return { ok: false };
  }

  function smartZoomIsVisible(value) {
    if (value === true || value === MsoTrue || value === 1) return true;
    if (typeof value === "string") return /^(true|msoTrue|-1|1)$/i.test(value.trim());
    return false;
  }

  function smartZoomWritePath(object, path, value) {
    if (!object || !path || !path.length) return false;
    try {
      let parent = object;
      for (let i = 0; i < path.length - 1; i += 1) {
        parent = parent[path[i]];
        if (parent === null || parent === undefined) return false;
      }
      parent[path[path.length - 1]] = value;
      return true;
    } catch (_) { return false; }
  }

  function smartZoomIsGroup(shape) {
    try { if (Number(shape.Type) === SMART_ZOOM_GROUP_TYPE) return true; } catch (_) {}
    try { return !!shape.GroupItems && Number(shape.GroupItems.Count) > 0; } catch (_) { return false; }
  }

  function smartZoomChildren(shape) {
    const result = [];
    if (!smartZoomIsGroup(shape)) return result;
    try {
      const items = shape.GroupItems;
      const count = Number(items.Count) || 0;
      for (let i = 1; i <= count; i += 1) {
        const child = items.Item(i);
        if (child) result.push(child);
      }
    } catch (_) {}
    return result;
  }

  function smartZoomAdjustment(shape) {
    try {
      const adjustments = shape.Adjustments;
      if (!adjustments) return null;
      if (hasMethod(adjustments, "Item")) {
        const value = Number(adjustments.Item(1));
        return isFinite(value) ? value : null;
      }
      const value = Number(adjustments[1] !== undefined ? adjustments[1] : adjustments[0]);
      return isFinite(value) ? value : null;
    } catch (_) { return null; }
  }

  function smartZoomSetAdjustment(shape, value) {
    try {
      const adjustments = shape.Adjustments;
      if (!adjustments) return false;
      if (hasMethod(adjustments, "Item")) {
        try {
          adjustments.Item(1, value);
          return true;
        } catch (_) {
          try {
            const item = adjustments.Item(1);
            if (item && item.Value !== undefined) { item.Value = value; return true; }
            if (item && item.value !== undefined) { item.value = value; return true; }
          } catch (__) {}
        }
      }
      if (adjustments[1] !== undefined) adjustments[1] = value;
      else adjustments[0] = value;
      return true;
    } catch (_) { return false; }
  }

  function smartZoomSnapshotStyles(shape, textOnly) {
    const styles = {};
    for (let i = 0; i < SMART_ZOOM_STYLE_SPECS.length; i += 1) {
      const spec = SMART_ZOOM_STYLE_SPECS[i];
      if (textOnly && String(spec.option || "").indexOf("scaleText") !== 0) continue;
      const result = spec.key === "textSize"
        ? smartZoomReadTextSize(shape, spec.paths)
        : smartZoomReadFirstNumber(shape, spec.paths);
      if (!result.ok) continue;
      // A mixed/empty WPS font size is commonly exposed as 0 or -2.  It is
      // not a writable concrete size; skipping it avoids collapsing text to
      // the 1pt safety clamp during a later zoom.
      if (spec.key === "textSize" && result.value <= 0) continue;
      if (spec.disabledWhenZero && Math.abs(result.value) < 0.000001) continue;
      if (spec.visibilityPaths && spec.visibilityPaths.length) {
        const visibility = smartZoomReadFirst(shape, spec.visibilityPaths);
        // If WPS does not expose the visibility flag, do not write the
        // numeric effect property: some hosts materialize a disabled effect
        // as soon as Offset/Radius/Depth is assigned.
        if (!visibility.ok || !smartZoomIsVisible(visibility.value)) continue;
      }
      styles[spec.key] = { value: result.value, path: result.path, option: spec.option };
    }
    return styles;
  }

  function smartZoomTextRanges(shape) {
    const ranges = [];
    const paths = [["TextFrame2", "TextRange"], ["TextFrame", "TextRange"]];
    for (let i = 0; i < paths.length; i += 1) {
      const result = smartZoomReadPath(shape, paths[i]);
      if (!result.ok || !result.value) continue;
      let duplicate = false;
      for (let j = 0; j < ranges.length; j += 1) {
        if (ranges[j] === result.value) { duplicate = true; break; }
      }
      if (!duplicate) ranges.push(result.value);
    }
    return ranges;
  }

  function smartZoomTextRunAt(textRange, index) {
    let runs = null;
    try { runs = textRange && textRange.Runs; } catch (_) {}
    if (!runs) return null;
    if (typeof runs === "function") {
      try {
        const run = runs.call(textRange, index, 1);
        if (run) return run;
      } catch (_) {}
      try {
        const run = runs.call(textRange, index);
        if (run) return run;
      } catch (_) {}
    }
    if (hasMethod(runs, "Item")) {
      try {
        const run = runs.Item(index);
        if (run) return run;
      } catch (_) {}
    }
    try { return runs[index] || null; } catch (_) { return null; }
  }

  function smartZoomTextRunKey(run) {
    const start = smartZoomReadNumber(run, ["Start"]);
    const length = smartZoomReadNumber(run, ["Length"]);
    if (start.ok && length.ok) return "range:" + start.value + ":" + length.value;
    const text = smartZoomReadPath(run, ["Text"]);
    return text.ok ? "text:" + String(text.value) : "";
  }

  function smartZoomSnapshotTextRangeRuns(textRange, context) {
    const result = [];
    const seen = Object.create(null);
    for (let i = 1; i <= SMART_ZOOM_MAX_OBJECTS; i += 1) {
      const run = smartZoomTextRunAt(textRange, i);
      if (!run) break;
      const key = smartZoomTextRunKey(run);
      // PowerPoint returns the last run when the requested start is past the
      // end. Stop on that repeated range instead of writing the last run over
      // and over; the same guard also protects older WPS JSAPI builds.
      if (key && seen[key]) break;
      if (!key && i > 1) break;
      if (key) seen[key] = true;
      const size = smartZoomReadFirstPositiveNumber(run, [["Font", "Size"]]);
      if (!size.ok) continue;
      context.count += 1;
      if (context.count > SMART_ZOOM_MAX_OBJECTS) {
        throw new Error("选区对象超过 500 个，请分批缩放以保持 WPS 稳定。");
      }
      result.push({
        shape: run,
        styles: { textSize: { value: size.value, path: ["Font", "Size"], option: "scaleText" } }
      });
    }
    return result;
  }

  function smartZoomSnapshotTextRuns(shape, context) {
    const ranges = smartZoomTextRanges(shape);
    for (let i = 0; i < ranges.length; i += 1) {
      const result = smartZoomSnapshotTextRangeRuns(ranges[i], context);
      if (result.length) return result;
    }
    return [];
  }

  function smartZoomSnapshotTableText(shape, context) {
    let table = null;
    try { table = shape && shape.Table; } catch (_) {}
    if (!table) return [];
    const rows = smartZoomReadNumber(table, ["Rows", "Count"]);
    const columns = smartZoomReadNumber(table, ["Columns", "Count"]);
    if (!rows.ok || !columns.ok || rows.value < 1 || columns.value < 1) return [];
    const result = [];
    for (let row = 1; row <= Math.floor(rows.value); row += 1) {
      for (let column = 1; column <= Math.floor(columns.value); column += 1) {
        context.count += 1;
        if (context.count > SMART_ZOOM_MAX_OBJECTS) {
          throw new Error("选区对象超过 500 个，请分批缩放以保持 WPS 稳定。");
        }
        let cell = null;
        try {
          if (hasMethod(table, "Cell")) cell = table.Cell(row, column);
        } catch (_) {}
        if (!cell) continue;
        let cellShape = null;
        try { cellShape = cell.Shape; } catch (_) {}
        if (!cellShape) continue;
        const styles = smartZoomSnapshotStyles(cellShape, true);
        const textRuns = styles.textSize ? [] : smartZoomSnapshotTextRuns(cellShape, context);
        if (Object.keys(styles).length || textRuns.length) {
          result.push({ shape: cellShape, styles: styles, textRuns: textRuns });
        }
      }
    }
    return result;
  }

  function smartZoomSnapshotNode(shape, context, depth) {
    const snapshotContext = context || { count: 0 };
    const level = depth || 0;
    if (level > SMART_ZOOM_MAX_GROUP_DEPTH) throw new Error("选区组合层级过深，已停止智能缩放以避免 WPS 卡死。");
    snapshotContext.count += 1;
    if (snapshotContext.count > SMART_ZOOM_MAX_OBJECTS) throw new Error("选区对象超过 500 个，请分批缩放以保持 WPS 稳定。");
    const width = smartZoomReadNumber(shape, ["Width"]);
    const height = smartZoomReadNumber(shape, ["Height"]);
    const left = smartZoomReadNumber(shape, ["Left"]);
    const top = smartZoomReadNumber(shape, ["Top"]);
    if (!width.ok || !height.ok || !left.ok || !top.ok) return null;

    const styles = smartZoomSnapshotStyles(shape);
    const node = {
      shape: shape,
      geometry: { left: left.value, top: top.value, width: width.value, height: height.value },
      styles: styles,
      textRuns: styles.textSize ? [] : smartZoomSnapshotTextRuns(shape, snapshotContext),
      tableText: smartZoomSnapshotTableText(shape, snapshotContext),
      children: []
    };
    const adjustment = smartZoomAdjustment(shape);
    if (adjustment !== null && Math.min(width.value, height.value) > 0) {
      node.cornerRadius = adjustment * Math.min(width.value, height.value);
    }
    try {
      const lock = Number(shape.LockAspectRatio);
      if (isFinite(lock)) node.lockAspectRatio = lock;
    } catch (_) {}
    const children = smartZoomChildren(shape);
    for (let i = 0; i < children.length; i += 1) {
      const child = smartZoomSnapshotNode(children[i], snapshotContext, level + 1);
      if (child) node.children.push(child);
    }
    return node;
  }

  function smartZoomSelectedShapes() {
    const windowObject = application().ActiveWindow;
    const selection = windowObject && windowObject.Selection;
    const range = selection && selection.ShapeRange;
    if (!range) return [];
    const result = [];
    try {
      const count = Number(range.Count) || 0;
      for (let i = 1; i <= count; i += 1) {
        const shape = range.Item(i);
        if (shape) result.push(shape);
      }
    } catch (_) {}
    if (!result.length) {
      try { const shape = asShape(range); if (shape) result.push(shape); } catch (_) {}
    }
    return result;
  }

  function smartZoomBounds(shapes) {
    if (!shapes || !shapes.length) return null;
    let bounds = null;
    for (let i = 0; i < shapes.length; i += 1) {
      const leftRead = smartZoomReadNumber(shapes[i], ["Left"]);
      const topRead = smartZoomReadNumber(shapes[i], ["Top"]);
      const widthRead = smartZoomReadNumber(shapes[i], ["Width"]);
      const heightRead = smartZoomReadNumber(shapes[i], ["Height"]);
      if (!leftRead.ok || !topRead.ok || !widthRead.ok || !heightRead.ok) continue;
      const left = leftRead.value;
      const top = topRead.value;
      const width = Math.max(0, widthRead.value);
      const height = Math.max(0, heightRead.value);
      if (!bounds) {
        bounds = { left: left, top: top, right: left + width, bottom: top + height };
      } else {
        bounds.left = Math.min(bounds.left, left);
        bounds.top = Math.min(bounds.top, top);
        bounds.right = Math.max(bounds.right, left + width);
        bounds.bottom = Math.max(bounds.bottom, top + height);
      }
    }
    return bounds && bounds.right > bounds.left && bounds.bottom > bounds.top ? bounds : null;
  }

  function smartZoomCountTree(node) {
    if (!node) return 0;
    let count = 1;
    for (let i = 0; i < node.children.length; i += 1) count += smartZoomCountTree(node.children[i]);
    return count;
  }

  function smartZoomDefaultOptions() {
    return {
      scaleText: true,
      protectTextReadability: true,
      scaleShapeLine: true,
      scaleShapeShadow: true,
      scaleShapeReflection: true,
      scaleShapeGlow: true,
      scaleSoftEdge: true,
      scaleShapeCorner: false,
      scaleShapeThreeD: true,
      scaleTextLine: true,
      scaleTextShadow: true,
      scaleTextReflection: true,
      scaleTextGlow: true,
      scaleTextThreeD: true
    };
  }

  function smartZoomOptionsSignature(options) {
    return Object.keys(options || {}).sort().map(function (key) {
      return key + "=" + (options[key] ? "1" : "0");
    }).join("|");
  }

  function smartZoomNormalizeAnchor(value) {
    const key = String(value || "center").toLowerCase().replace(/_/g, "-");
    if (key === "top-left" || key === "topleft" || key === "左上") return "top-left";
    if (key === "top-right" || key === "topright" || key === "右上") return "top-right";
    if (key === "bottom-left" || key === "bottomleft" || key === "左下") return "bottom-left";
    if (key === "bottom-right" || key === "bottomright" || key === "右下") return "bottom-right";
    return "center";
  }

  function smartZoomClampPercent(value) {
    const percent = Number(value);
    if (!isFinite(percent)) return 100;
    return Math.max(SMART_ZOOM_MIN_PERCENT, Math.min(SMART_ZOOM_MAX_PERCENT, percent));
  }

  function smartZoomBegin() {
    // Re-picking is a hard session boundary.  Invalidate the previous
    // session before reading the new selection so a failed pick cannot leave
    // the old selection writable through a delayed callback.
    smartZoomSession = null;
    const shapes = smartZoomSelectedShapes();
    if (!shapes.length) throw new Error("请先选择一个或多个图形。");
    const nodes = [];
    const snapshotContext = { count: 0 };
    for (let i = 0; i < shapes.length; i += 1) {
      const node = smartZoomSnapshotNode(shapes[i], snapshotContext, 0);
      if (node) nodes.push(node);
    }
    if (!nodes.length) throw new Error("无法读取选中图形的尺寸。");
    const bounds = smartZoomBounds(shapes);
    if (!bounds) throw new Error("选中图形的边界无效。");
    smartZoomSession = {
      sessionId: ++smartZoomSessionSeq,
      nodes: nodes,
      shapes: shapes,
      bounds: bounds,
      anchor: "center",
      options: smartZoomDefaultOptions(),
      percent: 100,
      hasApplied: false,
      lastAppliedFactor: 1,
      lastAppliedAnchor: "center",
      lastAppliedOptions: smartZoomOptionsSignature(smartZoomDefaultOptions())
    };
    return smartZoomInfo();
  }

  function smartZoomAnchorPoint(bounds, anchor) {
    let x = (bounds.left + bounds.right) / 2;
    let y = (bounds.top + bounds.bottom) / 2;
    if (anchor === "top-left" || anchor === "bottom-left") x = bounds.left;
    if (anchor === "top-right" || anchor === "bottom-right") x = bounds.right;
    if (anchor === "top-left" || anchor === "top-right") y = bounds.top;
    if (anchor === "bottom-left" || anchor === "bottom-right") y = bounds.bottom;
    return { x: x, y: y };
  }

  function smartZoomApplyGeometry(node, factor) {
    const old = node.geometry;
    const anchor = smartZoomAnchorPoint(smartZoomSession.bounds, smartZoomSession.anchor);
    const newWidth = old.width * factor;
    const newHeight = old.height * factor;
    const newLeft = anchor.x + (old.left - anchor.x) * factor;
    const newTop = anchor.y + (old.top - anchor.y) * factor;
    let lock = null;
    try { lock = Number(node.shape.LockAspectRatio); } catch (_) {}
    try { if (lock !== null && isFinite(lock)) node.shape.LockAspectRatio = MsoFalse; } catch (_) {}
    try { node.shape.Width = newWidth; } catch (_) {}
    try { node.shape.Height = newHeight; } catch (_) {}
    try { node.shape.Left = newLeft; } catch (_) {}
    try { node.shape.Top = newTop; } catch (_) {}
    try { if (lock !== null && isFinite(lock)) node.shape.LockAspectRatio = lock; } catch (_) {}
  }

  function smartZoomTextSizeFromStyles(styles) {
    const style = styles && styles.textSize;
    if (!style) return 0;
    const value = Number(style.value);
    return isFinite(value) && value > 0 ? value : 0;
  }

  function smartZoomMinTableTextSize(tableText) {
    let minimum = 0;
    const entries = tableText || [];
    for (let i = 0; i < entries.length; i += 1) {
      const entry = entries[i] || {};
      const direct = smartZoomTextSizeFromStyles(entry.styles);
      if (direct > 0 && (!minimum || direct < minimum)) minimum = direct;
      const runs = entry.textRuns || [];
      for (let j = 0; j < runs.length; j += 1) {
        const runSize = smartZoomTextSizeFromStyles(runs[j] && runs[j].styles);
        if (runSize > 0 && (!minimum || runSize < minimum)) minimum = runSize;
      }
    }
    return minimum;
  }

  function smartZoomTableTextFactor(tableText, factor) {
    if (!isFinite(factor) || factor >= 1 || !smartZoomSession ||
        smartZoomSession.options.protectTextReadability === false ||
        smartZoomSession.options.scaleText === false) return factor;
    const minimum = smartZoomMinTableTextSize(tableText);
    if (!(minimum > 0)) return factor;
    // Apply one factor to every cell.  If the smallest source size would fall
    // below 4pt, raise the *whole table* just enough to keep that smallest size
    // readable; cap at 1 so an already tiny source font is never enlarged by a
    // zoom-out operation.
    const readableFactor = Math.min(1, SMART_ZOOM_MIN_TABLE_TEXT_SIZE_PT / minimum);
    return Math.max(factor, readableFactor);
  }

  function smartZoomApplyStyleSet(shape, styles, factor, minimumTextSize) {
    styles = styles || {};
    const textFloor = Number(minimumTextSize) > 0 ? Number(minimumTextSize) : SMART_ZOOM_MIN_TEXT_SIZE_PT;
    const styleKeys = Object.keys(styles);
    // WPS may re-fit text when a text-frame margin is assigned.  Apply the
    // final font size after all margins so a host-side auto-fit cannot shrink
    // the requested size a second time.
    styleKeys.sort(function (a, b) {
      if (a === b) return 0;
      if (a === "textSize") return 1;
      if (b === "textSize") return -1;
      return 0;
    });
    styleKeys.forEach(function (key) {
      const style = styles[key];
      const scale = smartZoomSession.options[style.option] === false ? 1 : factor;
      let value = style.value * scale;
      if (key === "textSize") {
        if (smartZoomSession.options.protectTextReadability !== false &&
            smartZoomSession.options[style.option] !== false && factor < 1 &&
            value > 0 && value < textFloor) {
          value = textFloor;
        }
        value = Math.max(1, value);
      }
      smartZoomWritePath(shape, style.path, value);
    });
  }

  function smartZoomApplyTextRuns(textRuns, factor, minimumTextSize) {
    const runs = textRuns || [];
    for (let i = 0; i < runs.length; i += 1) {
      smartZoomApplyStyleSet(runs[i].shape, runs[i].styles, factor, minimumTextSize);
    }
  }

  function smartZoomApplyStyles(node, factor) {
    smartZoomApplyStyleSet(node.shape, node.styles, factor);
    smartZoomApplyTextRuns(node.textRuns, factor);
    const tableText = node.tableText || [];
    const tableFactor = smartZoomTableTextFactor(tableText, factor);
    for (let i = 0; i < tableText.length; i += 1) {
      smartZoomApplyStyleSet(tableText[i].shape, tableText[i].styles, tableFactor, SMART_ZOOM_MIN_TABLE_TEXT_SIZE_PT);
      smartZoomApplyTextRuns(tableText[i].textRuns, tableFactor, SMART_ZOOM_MIN_TABLE_TEXT_SIZE_PT);
    }
    if (node.cornerRadius !== undefined) {
      const minSize = Math.min(node.geometry.width * factor, node.geometry.height * factor);
      if (minSize > 0) {
        const radius = smartZoomSession.options.scaleShapeCorner === false
          ? node.cornerRadius : node.cornerRadius * factor;
        smartZoomSetAdjustment(node.shape, Math.max(0, Math.min(1, radius / minSize)));
      }
    }
    for (let i = 0; i < node.children.length; i += 1) smartZoomApplyStyles(node.children[i], factor);
  }

  function smartZoomApply(percent, options) {
    const config = options && typeof options === "object" ? options : {};
    if (!smartZoomSession) {
      if (config._sessionId !== undefined) {
        const staleInfo = smartZoomInfo();
        staleInfo.stale = true;
        return staleInfo;
      }
      throw new Error("智能缩放会话已结束，请先点击“重新拾取选区”。");
    }
    if (config._sessionId !== undefined && String(config._sessionId) !== String(smartZoomSession.sessionId)) {
      const staleInfo = smartZoomInfo();
      staleInfo.stale = true;
      return staleInfo;
    }
    const force = config._force === true;
    smartZoomSession.percent = smartZoomClampPercent(percent);
    if (config.anchor !== undefined) smartZoomSession.anchor = smartZoomNormalizeAnchor(config.anchor);
    Object.keys(smartZoomSession.options).forEach(function (key) {
      if (config[key] !== undefined) smartZoomSession.options[key] = !!config[key];
    });
    const factor = smartZoomSession.percent / 100;
    const optionsSignature = smartZoomOptionsSignature(smartZoomSession.options);
    const firstUnchangedApply = !smartZoomSession.hasApplied && factor === 1 && !force;
    const geometryDirty = force || (!firstUnchangedApply && (
      factor !== smartZoomSession.lastAppliedFactor ||
      (factor !== 1 && smartZoomSession.anchor !== smartZoomSession.lastAppliedAnchor)
    ));
    const stylesDirty = force || (!firstUnchangedApply && (
      factor !== smartZoomSession.lastAppliedFactor ||
      optionsSignature !== smartZoomSession.lastAppliedOptions
    ));
    if (geometryDirty || stylesDirty) {
      for (let i = 0; i < smartZoomSession.nodes.length; i += 1) {
        if (geometryDirty) smartZoomApplyGeometry(smartZoomSession.nodes[i], factor);
        if (stylesDirty) smartZoomApplyStyles(smartZoomSession.nodes[i], factor);
      }
    }
    smartZoomSession.hasApplied = true;
    smartZoomSession.lastAppliedFactor = factor;
    smartZoomSession.lastAppliedAnchor = smartZoomSession.anchor;
    smartZoomSession.lastAppliedOptions = optionsSignature;
    return smartZoomInfo();
  }

  function smartZoomReset(options) {
    const config = options && typeof options === "object" ? options : {};
    config._force = true;
    return smartZoomApply(100, config);
  }

  function smartZoomCurrentWidth() {
    if (!smartZoomSession) return 0;
    try {
      const bounds = smartZoomBounds(smartZoomSession.shapes);
      return bounds ? bounds.right - bounds.left : 0;
    } catch (_) { return 0; }
  }

  function smartZoomInfo() {
    if (!smartZoomSession) return { ready: false, count: 0, objectCount: 0, percent: 100, widthCm: 0, originalWidthCm: 0, anchor: "center", sessionId: null };
    let objectCount = 0;
    smartZoomSession.nodes.forEach(function (node) { objectCount += smartZoomCountTree(node); });
    const originalWidth = smartZoomSession.bounds.right - smartZoomSession.bounds.left;
    return {
      ready: true,
      count: smartZoomSession.nodes.length,
      objectCount: objectCount,
      percent: smartZoomSession.percent,
      anchor: smartZoomSession.anchor,
      widthCm: Math.round(smartZoomCurrentWidth() / SMART_ZOOM_PT_PER_CM * 100) / 100,
      originalWidthCm: Math.round(originalWidth / SMART_ZOOM_PT_PER_CM * 100) / 100,
      minPercent: SMART_ZOOM_MIN_PERCENT,
      maxPercent: SMART_ZOOM_MAX_PERCENT,
      sessionId: smartZoomSession.sessionId
    };
  }

  function smartZoomPercentForWidth(widthCm) {
    if (!smartZoomSession) return 0;
    const originalWidth = smartZoomSession.bounds.right - smartZoomSession.bounds.left;
    if (!(originalWidth > 0) || !(Number(widthCm) > 0)) return 0;
    return smartZoomClampPercent(Number(widthCm) * SMART_ZOOM_PT_PER_CM / originalWidth * 100);
  }

  function smartZoomEnd() {
    smartZoomSession = null;
    return true;
  }

  function slideOf(shape) {
    let slide = shape && shape.Parent;
    if (!slide || !slide.Shapes) {
      try { slide = shape.Parent.Parent; } catch (_) {}
    }
    if (!slide || !slide.Shapes) throw new Error("无法取得图片所在幻灯片。");
    return slide;
  }

  // WPS exposes PictureFormat on EVERY shape (textboxes, autoshapes, charts
  // included), so the legacy probe is unusable. Judge by shape.Type instead:
  // msoPicture = 13, msoLinkedPicture = 11. When Type is unavailable on the
  // host, fall back to the legacy PictureFormat probe.
  function isPicture(shape) {
    try {
      const type = Number(shape.Type);
      if (type === 13 || type === 11) return true;
      if (type === 0 || type === -9999 || isNaN(type)) {
        try { void shape.PictureFormat; return true; } catch (_) { return false; }
      }
      return false;
    } catch (_) {
      try { void shape.PictureFormat; return true; } catch (__) { return false; }
    }
  }

  function clamp(value, minimum, maximum) {
    if (maximum < minimum) return 0;
    return Math.max(minimum, Math.min(maximum, value));
  }

  // Recover the natural size of the ORIGINAL picture from measurable values:
  // a duplicate with crops zeroed shows the full image at the current zoom
  // (fullW, fullH), the frame is (frameW, frameH) and the crop amounts are in
  // original-image points, so  natural = (CL+CR) * fullW / (fullW - frameW).
  // Pure function, exported for unit tests.
  function recoverNaturalSize(frameW, frameH, cropLeft, cropRight, cropTop, cropBottom, fullW, fullH) {
    let naturalW = 0;
    let naturalH = 0;
    if (!(fullW > 0 && fullH > 0)) return { naturalW: 0, naturalH: 0, aspect: 0 };
    const aspect = fullW / fullH;
    if ((cropLeft + cropRight) > 0.05 && Math.abs(fullW - frameW) > 0.05) {
      naturalW = (cropLeft + cropRight) * fullW / (fullW - frameW);
    }
    if ((cropTop + cropBottom) > 0.05 && Math.abs(fullH - frameH) > 0.05) {
      naturalH = (cropTop + cropBottom) * fullH / (fullH - frameH);
    }
    if (!(naturalW > 0)) naturalW = naturalH > 0 ? naturalH * aspect : 0;
    if (!(naturalH > 0)) naturalH = naturalW > 0 ? naturalW / aspect : 0;
    return { naturalW: naturalW, naturalH: naturalH, aspect: aspect };
  }

  // Pictures that share one embedded media file are not necessarily the same
  // visible picture: WPS can store a vertical/horizontal contact sheet once
  // and expose different, disjoint regions through PictureFormat crops. Keep
  // source identity and visible-region identity as two separate dimensions.
  const CROP_GROUP_MIN_COVERAGE = 0.65;

  function fullCropWindow(known, raw) {
    return { l: 0, t: 0, r: 1, b: 1, known: known !== false, raw: raw || [0, 0, 0, 0] };
  }

  function normalizeCropWindow(windowValue) {
    const value = windowValue || {};
    const l = clamp(num(value.l), 0, 1);
    const t = clamp(num(value.t), 0, 1);
    const r = clamp(value.r === undefined ? 1 : num(value.r), 0, 1);
    const b = clamp(value.b === undefined ? 1 : num(value.b), 0, 1);
    if (r - l <= 0.0001 || b - t <= 0.0001) {
      return fullCropWindow(false, Array.isArray(value.raw) ? value.raw.slice(0, 4) : [0, 0, 0, 0]);
    }
    return {
      l: l,
      t: t,
      r: r,
      b: b,
      known: value.known !== false,
      raw: Array.isArray(value.raw) ? value.raw.slice(0, 4).map(num) : [0, 0, 0, 0]
    };
  }

  function cropWindowForShape(shape, fullW, fullH) {
    let cropLeft = 0;
    let cropRight = 0;
    let cropTop = 0;
    let cropBottom = 0;
    try {
      const pf = shape && shape.PictureFormat;
      cropLeft = num(pf && pf.CropLeft);
      cropRight = num(pf && pf.CropRight);
      cropTop = num(pf && pf.CropTop);
      cropBottom = num(pf && pf.CropBottom);
    } catch (_) {}
    const raw = [cropLeft, cropRight, cropTop, cropBottom];
    if (raw.reduce(function (sum, value) { return sum + Math.abs(value); }, 0) <= 0.05) {
      return fullCropWindow(true, raw);
    }
    const recovered = recoverNaturalSize(
      num(shape && shape.Width), num(shape && shape.Height),
      cropLeft, cropRight, cropTop, cropBottom,
      num(fullW), num(fullH)
    );
    if (!(recovered.naturalW > 0 && recovered.naturalH > 0)) {
      // Unknown is intentionally not treated as a full-image crop. The
      // compatibility fallback below compares the raw crop values instead,
      // which prevents an unmeasurable crop from silently joining everything.
      return fullCropWindow(false, raw);
    }
    return normalizeCropWindow({
      l: cropLeft / recovered.naturalW,
      t: cropTop / recovered.naturalH,
      r: 1 - cropRight / recovered.naturalW,
      b: 1 - cropBottom / recovered.naturalH,
      known: true,
      raw: raw
    });
  }

  function cropWindowCoverage(aValue, bValue) {
    const a = normalizeCropWindow(aValue);
    const b = normalizeCropWindow(bValue);
    if (!a.known || !b.known) {
      const ar = a.raw || [];
      const br = b.raw || [];
      if (ar.length !== 4 || br.length !== 4) return 0;
      for (let i = 0; i < 4; i += 1) {
        if (Math.abs(num(ar[i]) - num(br[i])) > 0.75) return 0;
      }
      return 1;
    }
    const intersectionW = Math.max(0, Math.min(a.r, b.r) - Math.max(a.l, b.l));
    const intersectionH = Math.max(0, Math.min(a.b, b.b) - Math.max(a.t, b.t));
    const areaA = Math.max(0, a.r - a.l) * Math.max(0, a.b - a.t);
    const areaB = Math.max(0, b.r - b.l) * Math.max(0, b.b - b.t);
    // Normalize by the larger visible area. A full composite must not count as
    // equivalent to a narrow crop merely because it contains that crop; both
    // objects need to show substantially the same region.
    const larger = Math.max(areaA, areaB);
    return larger > 0 ? (intersectionW * intersectionH) / larger : 0;
  }

  function cropWindowCompatible(a, b) {
    return cropWindowCoverage(a, b) >= CROP_GROUP_MIN_COVERAGE;
  }

  function cropWindowGroupKey(windowValue) {
    const value = normalizeCropWindow(windowValue);
    if (!value.known) {
      return "raw-" + value.raw.map(function (part) { return Math.round(num(part) * 4) / 4; }).join("-");
    }
    return [value.l, value.t, value.r, value.b].map(function (part) { return Math.round(part * 1000); }).join("-");
  }

  function rawCropPreviewKey(shape) {
    const windowValue = cropWindowForShape(shape, 0, 0);
    return cropWindowGroupKey(windowValue);
  }

  // Fast path: for a shape with valid link metadata whose source file is
  // unchanged on disk (fingerprint match), the crop baseline is exactly
  // px*0.75 of that file. This skips the scratch copy/paste round-trip per
  // shape, which is the dominant cost of batch replacement on linked decks.
  function linkedNaturalSize(shape, cache) {
    try {
      const meta = parseLink(String(shape.AlternativeText || ""));
      if (!meta || !meta.src || !meta.fileFp) return null;
      const key = "src:" + meta.src;
      if (cache) {
        if (Object.prototype.hasOwnProperty.call(cache, key)) return cache[key].ok ? cache[key] : null;
      }
      let entry = { ok: false, w: 0, h: 0 };
      try {
        if (fileExists(meta.src)) {
          const binary = fileSystem().readAsBinaryString(meta.src);
          if (binary) {
            const px = imagePixelSize(binary);
            if (px && px.w > 0 && px.h > 0 && String(fnv1a(binary)) === String(meta.fileFp)) {
              entry = { ok: true, w: px.w * 0.75, h: px.h * 0.75 };
            }
          }
        }
      } catch (_) {}
      if (cache) cache[key] = entry;
      return entry.ok ? entry : null;
    } catch (_) { return null; }
  }

  function captureState(shape, scratch, cache) {
    const pf = shape.PictureFormat;
    const cropLeft = num(pf.CropLeft);
    const cropRight = num(pf.CropRight);
    const cropTop = num(pf.CropTop);
    const cropBottom = num(pf.CropBottom);
    const frameW = num(shape.Width);
    const frameH = num(shape.Height);

    let fullW = 0;
    let fullH = 0;
    if (scratch) {
      // Paste a copy into the scratch slide, zero the crops: WPS reports
      // the full image display size at the current zoom (Duplicate inside
      // the source deck returns 1x1 in WPS, so paste is required).
      try {
        const slide = scratch.ensure();
        scratch.clear();
        const shapes = slide.Shapes;
        const before = Number(shapes.Count) || 0;
        let result = null;
        let after = Number(shapes.Count) || 0;
        for (let attempt = 0; attempt < 5 && after <= before; attempt += 1) {
          try { shape.Copy(); } catch (_) {}
          try { result = shapes.Paste(); } catch (_) {}
          after = Number(shapes.Count) || 0;
        }
        if (after > before) {
          const pasted = asShape(result) || shapes.Item(before + 1);
          const dpf = pasted.PictureFormat;
          try { dpf.CropLeft = 0; } catch (_) {}
          try { dpf.CropRight = 0; } catch (_) {}
          try { dpf.CropTop = 0; } catch (_) {}
          try { dpf.CropBottom = 0; } catch (_) {}
          fullW = num(pasted.Width);
          fullH = num(pasted.Height);
          // WPS occasionally reports a 1x1 placeholder after paste; treat it
          // as a failed capture instead of silently losing the crop.
          if (fullW <= 2 || fullH <= 2) { fullW = 0; fullH = 0; }
        }
      } catch (_) {}
      try { scratch.clear(); } catch (_) {}
    }
    if (!(fullW > 0 && fullH > 0)) {
      // Fallback: duplicate in place (some builds still report correctly).
      let dup = null;
      try {
        dup = shape.Duplicate();
        const dpf = dup.PictureFormat;
        try { dpf.CropLeft = 0; } catch (_) {}
        try { dpf.CropRight = 0; } catch (_) {}
        try { dpf.CropTop = 0; } catch (_) {}
        try { dpf.CropBottom = 0; } catch (_) {}
        fullW = num(dup.Width);
        fullH = num(dup.Height);
        if (fullW <= 2 || fullH <= 2) { fullW = 0; fullH = 0; }
      } catch (_) {}
      try { if (dup) dup.Delete(); } catch (_) {}
    }

    let naturalW = 0;
    let naturalH = 0;
    const linkedNatural = cache ? linkedNaturalSize(shape, cache) : null;
    if (linkedNatural && linkedNatural.w > 0 && linkedNatural.h > 0) {
      naturalW = linkedNatural.w;
      naturalH = linkedNatural.h;
      fullW = linkedNatural.w;
      fullH = linkedNatural.h;
    } else {
      const recovered = recoverNaturalSize(frameW, frameH, cropLeft, cropRight, cropTop, cropBottom, fullW, fullH);
      naturalW = recovered.naturalW;
      naturalH = recovered.naturalH;
    }
    const fL = naturalW > 0 ? clamp(cropLeft / naturalW, 0, 1) : 0;
    const fR = naturalW > 0 ? clamp(cropRight / naturalW, 0, 1) : 0;
    const fT = naturalH > 0 ? clamp(cropTop / naturalH, 0, 1) : 0;
    const fB = naturalH > 0 ? clamp(cropBottom / naturalH, 0, 1) : 0;

    let lockAspectRatio = MsoTrue;
    try { lockAspectRatio = shape.LockAspectRatio; } catch (_) {}
    let name = "";
    let alternativeText = "";
    let zOrder = 1;
    try { name = String(shape.Name || ""); } catch (_) {}
    try { alternativeText = String(shape.AlternativeText || ""); } catch (_) {}
    try { zOrder = num(shape.ZOrderPosition) || 1; } catch (_) {}

    const rawHasCrop = (Math.abs(cropLeft) + Math.abs(cropRight) + Math.abs(cropTop) + Math.abs(cropBottom)) > 0.05;
    const captureFailed = rawHasCrop && !(naturalW > 0 && naturalH > 0);

    return {
      captureFailed: captureFailed,
      rawHasCrop: rawHasCrop,
      left: num(shape.Left),
      top: num(shape.Top),
      width: frameW,
      height: frameH,
      rotation: num(shape.Rotation),
      flipH: isTrue(shape.HorizontalFlip),
      flipV: isTrue(shape.VerticalFlip),
      lockAspectRatio: lockAspectRatio,
      name: name,
      alternativeText: alternativeText,
      zOrder: zOrder,
      cropLeft: cropLeft,
      cropRight: cropRight,
      cropTop: cropTop,
      cropBottom: cropBottom,
      fullW: fullW,
      fullH: fullH,
      fL: fL,
      fR: fR,
      fT: fT,
      fB: fB,
      hasCrop: (Math.abs(fL) + Math.abs(fR) + Math.abs(fT) + Math.abs(fB)) > 0.0005
    };
  }

  // Probe2-verified crop reconstruction: keep the same visible-window
  // fractions, scale the new picture so the visible part fills the old frame
  // exactly; extra width/height becomes balanced extra crop. Pure function,
  // exported for unit tests.
  function computeNewCrops(fL, fR, fT, fB, naturalW, naturalH, frameW, frameH) {
    if (!(naturalW > 0 && naturalH > 0)) throw new Error("新图片尺寸无效。");
    if (Math.abs(fL) + Math.abs(fR) + Math.abs(fT) + Math.abs(fB) <= 0.0005) {
      return { cropLeft: 0, cropRight: 0, cropTop: 0, cropBottom: 0, zoom: 1 };
    }
    const visW = naturalW * (1 - fL - fR);
    const visH = naturalH * (1 - fT - fB);
    if (!(visW > 0 && visH > 0)) throw new Error("新图片尺寸过小，无法保持原裁剪效果。");
    const zoom = Math.max(frameW / visW, frameH / visH);
    const extraW = Math.max(0, (zoom * visW - frameW) / zoom);
    const extraH = Math.max(0, (zoom * visH - frameH) / zoom);
    return {
      cropLeft: Math.max(0, fL * naturalW + extraW / 2),
      cropRight: Math.max(0, fR * naturalW + extraW / 2),
      cropTop: Math.max(0, fT * naturalH + extraH / 2),
      cropBottom: Math.max(0, fB * naturalH + extraH / 2),
      zoom: zoom
    };
  }

  function applyPreservedCrop(shape, state, naturalW, naturalH) {
    const crops = computeNewCrops(state.fL, state.fR, state.fT, state.fB, naturalW, naturalH, state.width, state.height);
    try { shape.LockAspectRatio = MsoFalse; } catch (_) {}
    const pf = shape.PictureFormat;
    try {
      pf.CropLeft = crops.cropLeft;
      pf.CropRight = crops.cropRight;
      pf.CropTop = crops.cropTop;
      pf.CropBottom = crops.cropBottom;
    } catch (_) {}
    shape.Width = state.width;
    shape.Height = state.height;
  }

  function copyCosmetics(source, target) {
    // These properties are optional in some WPS builds; each is isolated so a
    // missing cosmetic API cannot invalidate the geometric replacement.
    try { target.Line.Visible = source.Line.Visible; target.Line.ForeColor.RGB = source.Line.ForeColor.RGB; target.Line.Weight = source.Line.Weight; } catch (_) {}
    try { target.Shadow.Visible = source.Shadow.Visible; target.Shadow.Transparency = source.Shadow.Transparency; target.Shadow.Blur = source.Shadow.Blur; target.Shadow.OffsetX = source.Shadow.OffsetX; target.Shadow.OffsetY = source.Shadow.OffsetY; } catch (_) {}
    try { target.SoftEdge.Radius = source.SoftEdge.Radius; } catch (_) {}
    try { target.Reflection.Visible = source.Reflection.Visible; target.Reflection.Transparency = source.Reflection.Transparency; target.Reflection.Size = source.Reflection.Size; target.Reflection.Blur = source.Reflection.Blur; target.Reflection.Distance = source.Reflection.Distance; } catch (_) {}
    try { target.Glow.Visible = source.Glow.Visible; target.Glow.Radius = source.Glow.Radius; try { target.Glow.Color.RGB = source.Glow.Color.RGB; } catch (_) {} } catch (_) {}
    try { target.PictureFormat.Brightness = source.PictureFormat.Brightness; target.PictureFormat.Contrast = source.PictureFormat.Contrast; target.PictureFormat.ColorType = source.PictureFormat.ColorType; } catch (_) {}
    try { target.PictureFormat.TransparencyColor = source.PictureFormat.TransparencyColor; target.PictureFormat.TransparentBackground = source.PictureFormat.TransparentBackground; } catch (_) {}
    try { target.AlternativeText = source.AlternativeText; } catch (_) {}
  }

  // WPS interprets CropLeft/Right/Top/Bottom against the image's 96-dpi
  // baseline (pixels * 0.75 pt), NOT against the insert size that
  // AddPicture reports (large images are capped at ~11 in). Parse the real
  // pixel size from the file header so crops survive replacement of big
  // images; fall back to the inserted shape size when parsing fails.
  function imagePixelSize(binary) {
    if (!binary) return null;
    let bytes = binary;
    let asArray = false;
    if (binary instanceof ArrayBuffer) { bytes = new Uint8Array(binary); asArray = true; }
    else if (binary instanceof Uint8Array) { asArray = true; }
    const len = bytes.length;
    function u8(i) { if (i < 0 || i >= len) return -1; return asArray ? bytes[i] : (bytes.charCodeAt(i) & 0xff); }
    function u16be(i) { return (u8(i) << 8) | u8(i + 1); }
    function u32be(i) { return (((u8(i) << 24) | (u8(i + 1) << 16) | (u8(i + 2) << 8) | u8(i + 3)) >>> 0); }
    function u16le(i) { return u8(i) | (u8(i + 1) << 8); }
    function u32le(i) { return (u8(i) | (u8(i + 1) << 8) | (u8(i + 2) << 16) | (u8(i + 3) << 24)) >>> 0; }
    if (len < 10) return null;
    if (u8(0) === 0x89 && u8(1) === 0x50 && u8(2) === 0x4e && u8(3) === 0x47 &&
        u8(4) === 0x0d && u8(5) === 0x0a && u8(6) === 0x1a && u8(7) === 0x0a) {
      const w = u32be(16); const h = u32be(20);
      return (w > 0 && h > 0) ? { w: w, h: h } : null;
    }
    if (u8(0) === 0xff && u8(1) === 0xd8 && u8(2) === 0xff) {
      let i = 2;
      while (i + 9 < len) {
        if (u8(i) !== 0xff) { i += 1; continue; }
        while (i < len && u8(i) === 0xff) i += 1;
        if (i >= len) break;
        const marker = u8(i);
        if (marker === 0xd8 || marker === 0xd9) { i += 1; continue; }
        if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) { i += 1; continue; }
        if (i + 2 >= len) break;
        const segLen = u16be(i + 1);
        if (segLen < 2) break;
        if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
          const h = u16be(i + 4); const w = u16be(i + 6);
          if (w > 0 && h > 0) return { w: w, h: h };
        }
        i += 1 + segLen;
      }
      return null;
    }
    if (u8(0) === 0x47 && u8(1) === 0x49 && u8(2) === 0x46 &&
        u8(3) === 0x38 && (u8(4) === 0x37 || u8(4) === 0x39) && u8(5) === 0x61) {
      const w = u16le(6); const h = u16le(8);
      return (w > 0 && h > 0) ? { w: w, h: h } : null;
    }
    if (len >= 26 && u8(0) === 0x42 && u8(1) === 0x4d) {
      const w = u32le(18); const h = Math.abs(u32le(22) | 0);
      return (w > 0 && h > 0) ? { w: w, h: h } : null;
    }
    if (len >= 30 && u32be(0) === 0x52494646 && u8(8) === 0x57 && u8(9) === 0x45 && u8(10) === 0x42 && u8(11) === 0x50) {
      const fourcc = String.fromCharCode(u8(12), u8(13), u8(14), u8(15));
      if (fourcc === "VP8X") {
        const w = 1 + (u8(24) | (u8(25) << 8) | (u8(26) << 16));
        const h = 1 + (u8(27) | (u8(28) << 8) | (u8(29) << 16));
        return (w > 0 && h > 0) ? { w: w, h: h } : null;
      }
      if (fourcc === "VP8L") {
        const bits = u32le(21);
        const w = 1 + (bits & 0x3fff);
        const h = 1 + ((bits >>> 14) & 0x3fff);
        return (w > 0 && h > 0) ? { w: w, h: h } : null;
      }
      if (fourcc === "VP8 ") {
        const w = u16le(26) & 0x3fff;
        const h = u16le(28) & 0x3fff;
        return (w > 0 && h > 0) ? { w: w, h: h } : null;
      }
      return null;
    }
    return null;
  }

  function naturalSizeForImage(imagePath, inserted, cache) {
    const key = "img:" + imagePath;
    if (cache && Object.prototype.hasOwnProperty.call(cache, key)) return cache[key];
    let px = null;
    try {
      const binary = fileSystem().readAsBinaryString(imagePath);
      if (binary) px = imagePixelSize(binary);
    } catch (_) {}
    const result = (px && px.w > 0 && px.h > 0 && isFinite(px.w) && isFinite(px.h))
      ? { w: px.w * 0.75, h: px.h * 0.75 }
      : { w: num(inserted.Width), h: num(inserted.Height) };
    if (cache) cache[key] = result;
    return result;
  }

  // Best-effort pixel size for clipboard images (some WPS webviews expose
  // navigator.clipboard). Never throws; falls back to the inserted size.
  function readClipboardPixelSize() {
    return new Promise(function (resolve) {
      let settled = false;
      let timer = null;
      const done = function (value) {
        if (settled) return;
        settled = true;
        if (timer) clearTimeout(timer);
        resolve(value || null);
      };
      try {
        const nav = global.navigator || {};
        const clip = nav.clipboard;
        if (!clip || typeof clip.read !== "function") { done(null); return; }
        timer = setTimeout(function () { done(null); }, 800);
        clip.read().then(function (items) {
          const list = items || [];
          for (let i = 0; i < list.length; i += 1) {
            const item = list[i];
            const types = (item && item.types) || [];
            let target = null;
            for (let t = 0; t < types.length; t += 1) {
              if (String(types[t]).indexOf("image/") === 0) { target = types[t]; break; }
            }
            if (!target || !item || typeof item.getType !== "function") continue;
            item.getType(target).then(function (blob) {
              if (!blob || typeof blob.arrayBuffer !== "function") { done(null); return; }
              blob.arrayBuffer().then(function (buffer) { done(imagePixelSize(buffer)); })
                .catch(function () { done(null); });
            }).catch(function () { done(null); });
            return;
          }
          done(null);
        }).catch(function () { done(null); });
      } catch (_) { done(null); }
    });
  }

  function replacePictureKeepCrop(oldShape, imagePath, scratch, cache) {
    const ownsScratch = !scratch;
    const sc = scratch || createScratchManager();
    try {
      const state = captureState(oldShape, sc, cache);
      if (state.captureFailed) {
        throw new Error("无法读取该图片的裁剪信息（复制/粘贴失败），已跳过以保留原图。");
      }
      const slide = slideOf(oldShape);
      let inserted = null;
      try {
        // Width/Height are intentionally omitted: WPS inserts at natural size.
        inserted = asShape(slide.Shapes.AddPicture(imagePath, MsoFalse, MsoTrue, state.left, state.top));
        if (!inserted) throw new Error("插入新图片失败。");
        const natural = naturalSizeForImage(imagePath, inserted, cache);
        applyPreservedCrop(inserted, state, natural.w, natural.h);
        if (isTrue(inserted.HorizontalFlip) !== state.flipH) inserted.Flip(MsoFlipHorizontal);
        if (isTrue(inserted.VerticalFlip) !== state.flipV) inserted.Flip(MsoFlipVertical);
        try { inserted.Rotation = state.rotation; } catch (_) {}
        inserted.Left = state.left;
        inserted.Top = state.top;
        try { inserted.LockAspectRatio = state.lockAspectRatio; } catch (_) {}
        copyCosmetics(oldShape, inserted);
        let moves = Math.max(0, num(inserted.ZOrderPosition) - (state.zOrder + 1));
        let guard = 0;
        while (moves > 0 && guard < 1000) {
          try { inserted.ZOrder(MsoSendBackward); } catch (_) {}
          moves -= 1;
          guard += 1;
        }
        oldShape.Delete();
        try { inserted.Name = state.name; } catch (_) {}
        try { inserted.AlternativeText = state.alternativeText; } catch (_) {}
        try { inserted.Select(); } catch (_) {}
        invalidatePanelInventoryCache();
        return inserted;
      } catch (error) {
        try { if (inserted) inserted.Delete(); } catch (_) {}
        throw error;
      }
    } finally {
      if (ownsScratch) { try { sc.dispose(); } catch (_) {} }
    }
  }
  // Scratch presentation used for every raster export (WPS JSAPI Shape has
  // no SaveAsPicture, so we render through a throw-away slide + Slide.Export).
  function createScratchManager() {
    let presentation = null;
    let slide = null;
    let ownsPresentation = false;
    function ensure() {
      if (slide) return slide;
      try {
        const presentations = application().Presentations;
        if (!presentations || !hasMethod(presentations, "Add")) throw new Error("no Presentations.Add");
        try {
          presentation = presentations.Add(0);
        } catch (_) {
          try { presentation = presentations.Add(); } catch (__) { presentation = null; }
        }
        ownsPresentation = !!presentation;
      } catch (_) {
        presentation = activePresentation();
        ownsPresentation = false;
      }
      try { presentation.Saved = true; } catch (_) {}
      const slides = presentation.Slides;
      slide = slides.Add(Number(slides.Count) + 1, 12);
      return slide;
    }
    function clear() {
      if (!slide) return;
      const shapes = slide.Shapes;
      const count = Number(shapes.Count) || 0;
      for (let i = count; i >= 1; i -= 1) {
        try { shapes.Item(i).Delete(); } catch (_) {}
      }
    }
    function dispose() {
      const p = presentation;
      const s = slide;
      presentation = null;
      slide = null;
      if (!p) return;
      try { if (s) s.Delete(); } catch (_) {}
      if (ownsPresentation) {
        try { p.Saved = true; } catch (_) {}
        try { p.Close(); } catch (_) {}
      }
    }
    return { ensure: ensure, clear: clear, dispose: dispose };
  }

  function tempPath(label) {
    const fs = fileSystem();
    let base = "";
    try { base = fs.tmpdir(); } catch (_) {}
    if (!base) {
      try { base = application().Env.GetTempPath(); } catch (_) {}
    }
    if (!base) throw new Error("无法取得 WPS 临时目录。");
    if (!/[\\/]$/.test(base)) base += "\\";
    return base + "PictureReplaceTools_" + label + "_" + Date.now() + "_" + Math.random().toString(16).slice(2) + ".png";
  }

  function fileExists(path) {
    try { if (fileSystem().Exists && fileSystem().Exists(path)) return true; } catch (_) {}
    try { if (fileSystem().existsSync && fileSystem().existsSync(path)) return true; } catch (_) {}
    return false;
  }

  function removeFile(path) {
    if (!path) return;
    try { fileSystem().unlinkSync(path); return; } catch (_) {}
    try { fileSystem().Remove(path); } catch (_) {}
  }

  // -------------------------------------------------------------------
  // Long-task progress channel. Ribbon/context-menu operations run in the
  // add-in's main context while the taskpane is a separate context, so
  // progress is exchanged through two tiny files in the WPS temp dir:
  //   picture_replace_task.json   - current {running,title,done,total,label,cancelled}
  //   picture_replace_cancel.flag - cancel request (written by any context)
  // The "#progress" taskpane polls the JSON; long loops poll the cancel flag.
  // -------------------------------------------------------------------
  const TASK_FILE_NAME = "picture_replace_task.json";
  const CANCEL_FILE_NAME = "picture_replace_cancel.flag";
  let taskPane = null;
  let lastTaskWrite = 0;

  function taskPath(name) {
    const fs = fileSystem();
    let base = "";
    try { base = fs.tmpdir(); } catch (_) {}
    if (!base) { try { base = application().Env.GetTempPath(); } catch (_) {} }
    if (!base) return "";
    if (!/[\\/]$/.test(base)) base += "\\";
    return base + name;
  }

  function writeTaskState(state) {
    const path = taskPath(TASK_FILE_NAME);
    if (!path) return;
    try {
      const payload = JSON.stringify({
        running: !!state.running,
        title: String(state.title || ""),
        done: num(state.done),
        total: num(state.total),
        label: String(state.label || ""),
        cancelled: !!state.cancelled
      });
      try { fileSystem().writeAsBinaryString(path, payload); } catch (_) { fileSystem().WriteFile(path, payload); }
    } catch (_) {}
  }

  function readTaskState() {
    const path = taskPath(TASK_FILE_NAME);
    if (!path) return null;
    try {
      if (fileSystem().Exists && !fileSystem().Exists(path)) return null;
      const raw = String(fileSystem().readAsBinaryString(path) || "");
      if (!raw) return null;
      const parsed = JSON.parse(raw);
      return {
        running: !!parsed.running,
        title: String(parsed.title || ""),
        done: num(parsed.done),
        total: num(parsed.total),
        label: String(parsed.label || ""),
        cancelled: !!parsed.cancelled
      };
    } catch (_) { return null; }
  }

  function requestCancelTask() {
    const path = taskPath(CANCEL_FILE_NAME);
    if (!path) return;
    try { fileSystem().writeAsBinaryString(path, "1"); } catch (_) { try { fileSystem().WriteFile(path, "1"); } catch (__) {} }
  }

  function clearCancelTask() {
    const path = taskPath(CANCEL_FILE_NAME);
    if (!path) return;
    try { fileSystem().unlinkSync(path); return; } catch (_) {}
    try { fileSystem().Remove(path); } catch (_) {}
  }

  function taskCancelled() {
    const path = taskPath(CANCEL_FILE_NAME);
    if (!path) return false;
    try { if (fileSystem().Exists && fileSystem().Exists(path)) return true; } catch (_) {}
    try { if (fileSystem().existsSync && fileSystem().existsSync(path)) return true; } catch (_) {}
    return false;
  }

  function beginTask(title, total) {
    clearCancelTask();
    lastTaskWrite = 0;
    writeTaskState({ running: true, title: title, done: 0, total: total || 0, label: "准备中" });
    try {
      if (!taskPane) {
        taskPane = application().CreateTaskPane(addinUrl("#progress"), title || "图片替换进度");
        // Prefer a floating (popup-style) progress pane; falls back to the
        // default docked position when the build does not support float.
        try { taskPane.DockPosition = 4; } catch (_) {}
        try { taskPane.Width = 340; taskPane.Height = 210; } catch (_) {}
      }
      if (taskPane) taskPane.Visible = true;
    } catch (_) {}
  }

  function reportTask(done, total, label) {
    const now = Date.now();
    if (now - lastTaskWrite < 120) return;
    lastTaskWrite = now;
    writeTaskState({ running: true, title: "", done: done, total: total, label: label || "" });
  }

  function openProgressPanel(title) {
    try {
      clearCancelTask();
      lastTaskWrite = 0;
      writeTaskState({ running: true, title: String(title || "图片替换进度"), done: 0, total: 0, label: "准备中" });
      const pane = application().CreateTaskPane(addinUrl("#progress"), String(title || "图片替换进度"));
      if (!pane) return null;
      try { pane.DockPosition = 4; } catch (_) {}
      try { pane.Width = 340; pane.Height = 210; } catch (_) {}
      pane.Visible = true;
      return pane;
    } catch (_) { return null; }
  }

  function writeTaskProgress(done, total, label) {
    reportTask(done, total, label);
  }

  function closeProgressPanel(pane) {
    try {
      writeTaskState({ running: false, done: 0, total: 0, label: "完成", cancelled: false });
      clearCancelTask();
      setTimeout(function () {
        try { if (pane) pane.Visible = false; } catch (_) {}
        try { removeFile(taskPath(TASK_FILE_NAME)); } catch (_) {}
      }, 1600);
    } catch (_) {}
  }

  function endTask(label, cancelled) {
    writeTaskState({ running: false, done: 0, total: 0, label: label || (cancelled ? "任务已取消" : "完成"), cancelled: !!cancelled });
    clearCancelTask();
    const hideDelay = cancelled ? 4000 : 2000;
    setTimeout(function () {
      try { if (taskPane) taskPane.Visible = false; } catch (_) {}
      try { removeFile(taskPath(TASK_FILE_NAME)); } catch (_) {}
    }, hideDelay);
  }

  // Yield to the event loop so progress UIs (taskpane, dialog) can repaint
  // between heavy synchronous JSAPI steps. Without this, a batch of linked
  // shapes runs back-to-back bridge calls and the UI looks frozen.
  function yieldUI(delayMs) {
    return new Promise(function (resolve) { setTimeout(resolve, Math.max(0, Number(delayMs) || 0)); });
  }

  async function runBatchWithProgress(title, work) {
    beginTask(title, 0);
    try {
      const result = await work(
        function (done, total, label) { reportTask(done, total, label); },
        function () { return taskCancelled(); }
      );
      const cancelled = !!(result && result.cancelled);
      endTask(cancelled ? title + "已取消" : title + "完成", cancelled);
      return result;
    } catch (error) {
      endTask("任务失败", false);
      throw error;
    }
  }


  // Render the FULL (uncropped, unrotated, flip-normalized) image of a shape
  // to a PNG through a scratch slide. Used to fingerprint the original image
  // so different crop instances of the same picture match each other.
  function exportUncroppedPreview(shape, scratch, path, extraFlipH, extraFlipV, exportPx) {
    try {
      const slide = scratch.ensure();
      scratch.clear();
      const shapes = slide.Shapes;
      const before = Number(shapes.Count) || 0;
      let result = null;
      let after = Number(shapes.Count) || 0;
      for (let attempt = 0; attempt < 3 && after <= before; attempt += 1) {
        try { shape.Copy(); } catch (_) {}
        try { result = shapes.Paste(); } catch (_) {}
        after = Number(shapes.Count) || 0;
      }
      if (after <= before) throw new Error("粘贴失败：临时文稿未生成图片对象。");
      const pasted = asShape(result) || shapes.Item(after);
      const pf = pasted.PictureFormat;
      if (!pf) throw new Error("粘贴对象不是图片。");
      try { pasted.Rotation = 0; } catch (_) {}
      if (isTrue(pasted.HorizontalFlip)) pasted.Flip(MsoFlipHorizontal);
      if (isTrue(pasted.VerticalFlip)) pasted.Flip(MsoFlipVertical);
      if (extraFlipH) pasted.Flip(MsoFlipHorizontal);
      if (extraFlipV) pasted.Flip(MsoFlipVertical);
      try { pf.CropLeft = 0; } catch (_) {}
      try { pf.CropRight = 0; } catch (_) {}
      try { pf.CropTop = 0; } catch (_) {}
      try { pf.CropBottom = 0; } catch (_) {}
      const fullW = num(pasted.Width);
      const fullH = num(pasted.Height);
      if (!(fullW > 0 && fullH > 0)) throw new Error("无法取得图片完整尺寸。");
      try { pasted.Line.Visible = MsoFalse; } catch (_) {}
      try { pasted.Shadow.Visible = MsoFalse; } catch (_) {}
      try { pasted.LockAspectRatio = MsoFalse; } catch (_) {}
      // Normalize to a fixed scale: fullW/fullH is proportional to the
      // natural size at the instance's own zoom, so scaling both to a 1024px
      // long side makes every instance of the same image render identically.
      let canvasW = fullW;
      let canvasH = fullH;
      if (fullW >= fullH) {
        canvasW = NORMALIZE_PX;
        canvasH = Math.max(1, Math.round(NORMALIZE_PX * fullH / fullW));
      } else {
        canvasH = NORMALIZE_PX;
        canvasW = Math.max(1, Math.round(NORMALIZE_PX * fullW / fullH));
      }
      try {
        const pageSetup = presentationOf(slide).PageSetup;
        if (pageSetup) {
          pageSetup.SlideWidth = canvasW;
          pageSetup.SlideHeight = canvasH;
        }
      } catch (_) {}
      pasted.Left = 0;
      pasted.Top = 0;
      pasted.Width = canvasW;
      pasted.Height = canvasH;
      removeFile(path);
      slide.Export(path, "PNG", exportPx || PREVIEW_PX, exportPx || PREVIEW_PX);
      return { ok: fileExists(path), fullW: fullW, fullH: fullH };
    } catch (error) {
      removeFile(path);
      return { ok: false, fullW: 0, fullH: 0 };
    } finally {
      try { scratch.clear(); } catch (_) {}
    }
  }

  function binaryToBytes(binary) {
    if (binary instanceof Uint8Array) return binary;
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i) & 255;
    return bytes;
  }

  function fnv1a(binary) {
    let hash = 2166136261;
    for (let i = 0; i < binary.length; i += 1) {
      hash ^= binary.charCodeAt(i) & 255;
      hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0).toString(16);
  }
  function fnvBytes(bytes) {
    let hash = 2166136261;
    for (let i = 0; i < bytes.length; i += 1) {
      hash ^= bytes[i];
      hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0).toString(16);
  }

  function canvasSignature(binary) {
    return new Promise(function (resolve, reject) {
      if (!global.Image || !global.document || !global.URL || !global.Blob) { reject(new Error("canvas unavailable")); return; }
      const blob = new Blob([binaryToBytes(binary)], { type: "image/png" });
      const url = global.URL.createObjectURL(blob);
      const image = new global.Image();
      image.onload = function () {
        try {
          const canvas = global.document.createElement("canvas");
          canvas.width = PREVIEW_PX; canvas.height = PREVIEW_PX;
          const context = canvas.getContext("2d");
          context.drawImage(image, 0, 0, PREVIEW_PX, PREVIEW_PX);
          const pixels = context.getImageData(0, 0, PREVIEW_PX, PREVIEW_PX).data;
          let hash = 2166136261;
          for (let i = 0; i < pixels.length; i += 4) {
            hash ^= (pixels[i] & 0xf0) ^ ((pixels[i + 1] & 0xf0) >>> 1) ^ ((pixels[i + 2] & 0xf0) >>> 2) ^ (pixels[i + 3] & 0xf0);
            hash = Math.imul(hash, 16777619);
          }
          global.URL.revokeObjectURL(url);
          resolve((hash >>> 0).toString(16));
        } catch (error) { global.URL.revokeObjectURL(url); reject(error); }
      };
      image.onerror = function () { global.URL.revokeObjectURL(url); reject(new Error("图片解码失败")); };
      image.src = url;
    });
  }

  async function signatureFromPath(path) {
    const binary = fileSystem().readAsBinaryString(path);
    try { return await canvasSignature(binary); } catch (_) { return fnv1a(binary); }
  }

  async function signaturesForShape(shape, scratch, pathPrefix) {
    const signatures = [];
    const variants = [[false, false], [true, false], [false, true], [true, true]];
    for (let i = 0; i < variants.length; i += 1) {
      const path = tempPath(pathPrefix + i);
      if (!exportUncroppedPreview(shape, scratch, path, variants[i][0], variants[i][1]).ok) { removeFile(path); continue; }
      signatures.push(await signatureFromPath(path));
      removeFile(path);
    }
    return signatures;
  }

  async function replaceAllMatching(referenceShape, imagePath, onProgress, cancelCheck) {
    const presentation = activePresentation();
    const scratch = createScratchManager();
    const cache = {};
    let cancelled = false;
    try {
      const referenceSignatures = new Set(await signaturesForShape(referenceShape, scratch, "reference"));
      if (!referenceSignatures.size) {
        throw new Error("无法读取参考图片：未能导出参考图。请确认选中的是图片；若仍失败，请点击“兼容性诊断”查看 Slide.Export 是否可用。");
      }
      const candidates = [];
      for (let slideIndex = 1; slideIndex <= presentation.Slides.Count; slideIndex += 1) {
        const slide = presentation.Slides.Item(slideIndex);
        const count = slide.Shapes.Count;
        for (let shapeIndex = 1; shapeIndex <= count; shapeIndex += 1) {
          const shape = slide.Shapes.Item(shapeIndex);
          if (isPicture(shape)) candidates.push(shape);
        }
      }

      const matches = [];
      for (let i = 0; i < candidates.length; i += 1) {
        if (cancelCheck && cancelCheck()) { cancelled = true; break; }
        if (onProgress) { try { onProgress(i + 1, candidates.length, "扫描匹配同源图片 " + (i + 1) + "/" + candidates.length); } catch (_) {} }
        await yieldUI();
        const path = tempPath("candidate");
        try {
          if (!exportUncroppedPreview(candidates[i], scratch, path, false, false).ok) continue;
          const signature = await signatureFromPath(path);
          if (referenceSignatures.has(signature)) matches.push(candidates[i]);
        } finally { removeFile(path); }
      }

      let success = 0;
      const matchTotal = matches.length;
      let i = 0;
      for (i = 0; i < matches.length; i += 1) {
        if (cancelCheck && cancelCheck()) { cancelled = true; break; }
        if (onProgress) { try { onProgress(i + 1, matchTotal, "替换图片 " + (i + 1) + "/" + matchTotal); } catch (_) {} }
        await yieldUI();
        try { if (replacePictureKeepCrop(matches[i], imagePath, scratch, cache)) success += 1; } catch (_) {}
      }
      const failed = matches.length - success - (cancelled ? Math.max(0, matchTotal - i) : 0);
      if (success > 0) invalidatePanelInventoryCache();
      return { matched: matches.length, success: success, failed: Math.max(0, failed), cancelled: cancelled };
    } finally {
      scratch.dispose();
    }
  }

  function pasteClipboardAsPng(scratch, path) {
    const slide = scratch.ensure();
    scratch.clear();
    const shapes = slide.Shapes;
    const before = Number(shapes.Count) || 0;
    let pasted = false;
    const formats = [PP_PASTE_PNG, PP_PASTE_BITMAP, PP_PASTE_JPG, PP_PASTE_GIF];
    for (let i = 0; i < formats.length && !pasted; i += 1) {
      try { shapes.PasteSpecial(formats[i]); pasted = true; } catch (_) {}
    }
    if (!pasted) {
      try { shapes.Paste(); pasted = (Number(shapes.Count) || 0) > before; } catch (_) {}
    }
    if (!pasted) throw new Error("剪贴板没有可粘贴的图片格式。");
    const after = Number(shapes.Count) || 0;
    if (after <= before) throw new Error("剪贴板粘贴没有生成图片对象。");
    try {
      const pastedShape = shapes.Item(after);
      const pf = pastedShape.PictureFormat;
      try { pf.CropLeft = 0; } catch (_) {}
      try { pf.CropRight = 0; } catch (_) {}
      try { pf.CropTop = 0; } catch (_) {}
      try { pf.CropBottom = 0; } catch (_) {}
      try { pastedShape.Rotation = 0; } catch (_) {}
      let w = num(pastedShape.Width) || 960;
      let h = num(pastedShape.Height) || 540;
      // Cap the scratch slide size so PageSetup always succeeds (WPS caps
      // slide dimensions); the exported PNG keeps the image aspect ratio.
      if (w > MAX_SCRATCH_SLIDE || h > MAX_SCRATCH_SLIDE) {
        const scale = MAX_SCRATCH_SLIDE / Math.max(w, h);
        w = Math.max(1, Math.round(w * scale));
        h = Math.max(1, Math.round(h * scale));
      }
      try {
        const pageSetup = presentationOf(slide).PageSetup;
        if (pageSetup) { pageSetup.SlideWidth = w; pageSetup.SlideHeight = h; }
      } catch (_) {}
      try { pastedShape.LockAspectRatio = MsoFalse; } catch (_) {}
      pastedShape.Left = 0;
      pastedShape.Top = 0;
      pastedShape.Width = w;
      pastedShape.Height = h;
      let exportW = CLIPBOARD_MAX_PX;
      let exportH = Math.max(1, Math.round(CLIPBOARD_MAX_PX * h / w));
      if (h > w) {
        exportH = CLIPBOARD_MAX_PX;
        exportW = Math.max(1, Math.round(CLIPBOARD_MAX_PX * w / h));
      }
      removeFile(path);
      slide.Export(path, "PNG", exportW, exportH);
      if (!fileExists(path)) throw new Error("无法导出剪贴板图片。");
      return path;
    } finally {
      scratch.clear();
    }
  }

  function readBrowserFile(file) {
    if (!file) return Promise.reject(new Error("请选择图片文件。"));
    if (file.size > MAX_BROWSER_FILE_BYTES) return Promise.reject(new Error("图片文件不能超过 50 MB。"));
    return new Promise(function (resolve, reject) {
      const reader = new global.FileReader();
      reader.onload = function () { resolve(String(reader.result)); };
      reader.onerror = function () { reject(new Error("读取图片文件失败。")); };
      reader.readAsDataURL(file);
    });
  }

  async function writeBrowserFile(file) {
    const dataUrl = await readBrowserFile(file);
    const comma = dataUrl.indexOf(",");
    if (comma < 0) throw new Error("图片数据格式无效。");
    const binary = global.atob(dataUrl.slice(comma + 1));
    const path = tempPath("selected");
    try { fileSystem().writeAsBinaryString(path, binary); } catch (_) { fileSystem().WriteFile(path, binary); }
    if (!fileExists(path)) throw new Error("无法把图片写入 WPS 临时目录。");
    return path;
  }

  async function replaceSelectedFromFile(path) {
    try { return replacePictureKeepCrop(selectedPicture(), path); } finally { removeFile(path); }
  }

  async function replaceAllFromFile(path, onProgress, cancelCheck) {
    try {
      const result = await replaceAllMatching(selectedPicture(), path, onProgress, cancelCheck);
      if (!result.matched) throw new Error("没有找到匹配的原图实例。");
      return result;
    } finally { removeFile(path); }
  }

  // Direct clipboard replacement: paste the clipboard image onto the target
  // slide itself (no intermediate PNG), then apply the preserved crop using
  // the exact same math as file replacement. The paste happens BEFORE
  // captureState because captureState copies the old shape and would
  // otherwise overwrite the user's clipboard.
  async function pasteReplacePictureKeepCrop(target) {
    const owns = createScratchManager();
    try {
      const slide = slideOf(target);
      const shapes = slide.Shapes;
      const before = Number(shapes.Count) || 0;
      let pasted = false;
      const formats = [PP_PASTE_PNG, PP_PASTE_BITMAP, PP_PASTE_JPG, PP_PASTE_GIF];
      for (let i = 0; i < formats.length && !pasted; i += 1) {
        try { shapes.PasteSpecial(formats[i]); pasted = true; } catch (_) {}
      }
      if (!pasted) {
        try { shapes.Paste(); pasted = (Number(shapes.Count) || 0) > before; } catch (_) {}
      }
      if (!pasted) throw new Error("剪贴板没有可粘贴的图片格式。");
      const after = Number(shapes.Count) || 0;
      if (after <= before) throw new Error("剪贴板粘贴没有生成图片对象。");
      const inserted = asShape(shapes.Item(after));
      if (!inserted || !inserted.PictureFormat) {
        try { if (inserted) inserted.Delete(); } catch (_) {}
        throw new Error("剪贴板内容无法作为图片粘贴。");
      }
      // Now capture the old shape state (its Copy clobbers the clipboard,
      // but the paste above already consumed it).
      const clipboardPx = await readClipboardPixelSize();
      const state = captureState(target, owns);
      const natural = clipboardPx && clipboardPx.w > 0
        ? { w: clipboardPx.w * 0.75, h: clipboardPx.h * 0.75 }
        : { w: num(inserted.Width), h: num(inserted.Height) };
      applyPreservedCrop(inserted, state, natural.w, natural.h);
      if (isTrue(inserted.HorizontalFlip) !== state.flipH) inserted.Flip(MsoFlipHorizontal);
      if (isTrue(inserted.VerticalFlip) !== state.flipV) inserted.Flip(MsoFlipVertical);
      try { inserted.Rotation = state.rotation; } catch (_) {}
      inserted.Left = state.left;
      inserted.Top = state.top;
      try { inserted.LockAspectRatio = state.lockAspectRatio; } catch (_) {}
      copyCosmetics(target, inserted);
      let moves = Math.max(0, num(inserted.ZOrderPosition) - (state.zOrder + 1));
      let guard = 0;
      while (moves > 0 && guard < 1000) {
        try { inserted.ZOrder(MsoSendBackward); } catch (_) {}
        moves -= 1;
        guard += 1;
      }
      target.Delete();
      try { inserted.Name = state.name; } catch (_) {}
      // Clipboard replacement intentionally carries NO link metadata: the
      // pasted content has no source file. Preserve only the user's own
      // accessibility text, never the stale PICRENEW link payload.
      try {
        const oldMeta = parseLink(state.alternativeText);
        inserted.AlternativeText = oldMeta ? (oldMeta.userAlt || "") : state.alternativeText;
      } catch (_) {}
      try { inserted.Select(); } catch (_) {}
      invalidatePanelInventoryCache();
      return inserted;
    } finally {
      owns.dispose();
    }
  }

  async function replaceSelectedFromClipboard() {
    return pasteReplacePictureKeepCrop(selectedPicture());
  }
  async function replaceAllFromClipboard(onProgress, cancelCheck) {
    const reference = selectedPicture();
    const path = tempPath("clipboard_batch");
    const scratch = createScratchManager();
    try {
      pasteClipboardAsPng(scratch, path);
      const result = await replaceAllMatching(reference, path, onProgress, cancelCheck);
      if (!result.matched) throw new Error("没有找到匹配的原图实例。");
      return result;
    } finally { scratch.dispose(); removeFile(path); }
  }


  // =====================================================================
  // v1.2.0 Picture Manager panel: link metadata + deck image inventory
  // =====================================================================

  function normalizePath(path) {
    return String(path || "").replace(/\\/g, "/").replace(/\/+/g, "/").toLowerCase();
  }

  function baseName(path) {
    const parts = String(path || "").split(/[\\/]/);
    return parts[parts.length - 1] || String(path || "");
  }

  function parseLink(alt) {
    const text = String(alt || "");
    if (text.indexOf(LINK_PREFIX) !== 0) return null;
    try {
      const meta = JSON.parse(text.slice(LINK_PREFIX.length));
      if (!meta || typeof meta !== "object") return null;
      return {
        v: num(meta.v) || LINK_VERSION,
        name: String(meta.name || ""),
        src: String(meta.src || ""),
        fileFp: String(meta.fileFp || ""),
        contentFp: String(meta.contentFp || ""),
        aspect: num(meta.aspect) || 0,
        userAlt: meta.userAlt === undefined ? "" : String(meta.userAlt)
      };
    } catch (_) { return null; }
  }

  function formatLink(meta) {
    const payload = {
      v: LINK_VERSION,
      name: String(meta.name || "").slice(0, 80),
      src: String(meta.src || ""),
      fileFp: String(meta.fileFp || ""),
      contentFp: String(meta.contentFp || ""),
      aspect: num(meta.aspect) || 0
    };
    if (meta.userAlt) payload.userAlt = String(meta.userAlt);
    return LINK_PREFIX + JSON.stringify(payload);
  }

  // Write link metadata onto a shape.AlternativeText, preserving the user's
  // own accessibility text (userAlt) when it is already known.
  function attachLink(shape, meta) {
    const existing = String(shape.AlternativeText || "");
    const parsed = parseLink(existing);
    const userAlt = meta.userAlt !== undefined
      ? String(meta.userAlt)
      : (parsed ? parsed.userAlt : (existing.indexOf(LINK_PREFIX) === 0 ? existing.slice(LINK_PREFIX.length) : ""));
    try { shape.AlternativeText = formatLink(Object.assign({}, meta, { userAlt: userAlt })); } catch (_) {}
  }

  function fingerprintFile(path) {
    const binary = fileSystem().readAsBinaryString(path);
    return fnv1a(binary);
  }

  function hasCropOf(shape) {
    try {
      const pf = shape.PictureFormat;
      const s = Math.abs(num(pf.CropLeft)) + Math.abs(num(pf.CropRight)) +
                Math.abs(num(pf.CropTop)) + Math.abs(num(pf.CropBottom));
      return s > 0.05;
    } catch (_) { return false; }
  }

  function zoneOf(shape, slide) {
    let sw = 960;
    let sh = 540;
    try {
      const ps = slide.PageSetup;
      sw = num(ps.SlideWidth) || sw;
      sh = num(ps.SlideHeight) || sh;
    } catch (_) {}
    const width = Math.max(0.001, num(shape.Width));
    const height = Math.max(0.001, num(shape.Height));
    const cx = (num(shape.Left) + width / 2) / sw;
    const cy = (num(shape.Top) + height / 2) / sh;
    const area = (width * height) / (sw * sh);
    if (area > 0.8) return "铺满页面";
    const hz = cx < 0.34 ? "左" : cx > 0.66 ? "右" : "中";
    const vt = cy < 0.34 ? "上" : cy > 0.66 ? "下" : "中";
    const map = {
      "上左": "左上角", "上中": "顶部居中", "上右": "右上角",
      "中左": "左侧中部", "中中": "页面中央", "中右": "右侧中部",
      "下左": "左下角", "下中": "底部居中", "下右": "右下角"
    };
    return map[vt + hz] || (vt + hz);
  }

  function resolveShape(presentation, slideIndex, shapeIndex) {
    const count = Number(presentation.Slides.Count) || 0;
    if (slideIndex < 1 || slideIndex > count) return null;
    const slide = presentation.Slides.Item(slideIndex);
    const shapeCount = Number(slide.Shapes.Count) || 0;
    if (shapeIndex < 1 || shapeIndex > shapeCount) return null;
    const shape = slide.Shapes.Item(shapeIndex);
    return isPicture(shape) ? shape : null;
  }

  // Convert exported PNG bytes to a data: URL with layered fallbacks:
  // btoa (fast path) -> FileReader over a Blob -> canvas decode.
  // Any WPS host missing all of these gets an empty string and the panel
  // shows a color placeholder instead of a blank tile.
  // Perceptual hash (64-bit dHash) over an 8x8 grayscale gradient grid.
  // Two images that differ only by re-encoding/resampling produce the same
  // or very close hashes, which is exactly what same-source grouping needs
  // (strict byte fingerprints split visually identical but re-encoded files).
  function dHashFromPixels(pixels, width) {
    let hash = "";
    for (let row = 0; row < 8; row += 1) {
      for (let col = 0; col < 8; col += 1) {
        const i1 = ((row * width) + col) * 4;
        const i2 = ((row * width) + col + 1) * 4;
        const l1 = 0.299 * pixels[i1] + 0.587 * pixels[i1 + 1] + 0.114 * pixels[i1 + 2];
        const l2 = 0.299 * pixels[i2] + 0.587 * pixels[i2 + 1] + 0.114 * pixels[i2 + 2];
        hash += l1 >= l2 ? "1" : "0";
      }
    }
    return hash;
  }

  function dHashDistance(a, b) {
    if (!a || !b || a.length !== b.length) return 99;
    let d = 0;
    for (let i = 0; i < a.length; i += 1) if (a[i] !== b[i]) d += 1;
    return d;
  }

  // A second, independent 64-bit perceptual signature. dHash alone can be
  // identical for flat gradients or similarly composed landscape images;
  // requiring a compatible aHash prevents those false-positive merges while
  // still allowing harmless re-encoding/resampling differences.
  function aHashFromPixels(pixels, width) {
    const luminance = [];
    let sum = 0;
    for (let row = 0; row < 8; row += 1) {
      for (let col = 0; col < 8; col += 1) {
        const index = ((row * width) + col) * 4;
        const value = 0.299 * pixels[index] + 0.587 * pixels[index + 1] + 0.114 * pixels[index + 2];
        luminance.push(value);
        sum += value;
      }
    }
    const average = sum / Math.max(1, luminance.length);
    return luminance.map(function (value) { return value >= average ? "1" : "0"; }).join("");
  }

  // Decode PNG bytes to a 64-bit dHash via Image + canvas (CEF hosts).
  function perceptualHashesFromBinary(binary) {
    return new Promise(function (resolve) {
      if (!(global.Image && global.document && global.URL && global.Blob)) { resolve({ dHash: "", aHash: "" }); return; }
      try {
        const url = global.URL.createObjectURL(new global.Blob([binaryToBytes(binary)], { type: "image/png" }));
        const img = new global.Image();
        img.onload = function () {
          try {
            const canvas = global.document.createElement("canvas");
            canvas.width = 9;
            canvas.height = 8;
            const ctx = canvas.getContext("2d");
            ctx.drawImage(img, 0, 0, 9, 8);
            const data = ctx.getImageData(0, 0, 9, 8).data;
            global.URL.revokeObjectURL(url);
            resolve({ dHash: dHashFromPixels(data, 9), aHash: aHashFromPixels(data, 9) });
          } catch (_) { global.URL.revokeObjectURL(url); resolve({ dHash: "", aHash: "" }); }
        };
        img.onerror = function () { global.URL.revokeObjectURL(url); resolve({ dHash: "", aHash: "" }); };
        img.src = url;
      } catch (_) { resolve({ dHash: "", aHash: "" }); }
    });
  }

  function fileToDataUrl(binary) {
    return new Promise(function (resolve) {
      if (global.btoa) {
        try { resolve("data:image/png;base64," + global.btoa(binary)); return; } catch (_) {}
      }
      if (global.FileReader && global.Blob) {
        try {
          const reader = new global.FileReader();
          reader.onload = function () { resolve(String(reader.result || "")); };
          reader.onerror = function () { resolve(""); };
          reader.readAsDataURL(new global.Blob([binaryToBytes(binary)], { type: "image/png" }));
          return;
        } catch (_) {}
      }
      if (global.Image && global.document && global.URL && global.Blob) {
        try {
          const url = global.URL.createObjectURL(new global.Blob([binaryToBytes(binary)], { type: "image/png" }));
          const img = new global.Image();
          img.onload = function () {
            try {
              if (!(Number(img.width) > 0 && Number(img.height) > 0)) { global.URL.revokeObjectURL(url); resolve(""); return; }
              const canvas = global.document.createElement("canvas");
              canvas.width = img.width;
              canvas.height = img.height;
              canvas.getContext("2d").drawImage(img, 0, 0);
              global.URL.revokeObjectURL(url);
              resolve(canvas.toDataURL("image/png"));
            } catch (_) { global.URL.revokeObjectURL(url); resolve(""); }
          };
          img.onerror = function () { global.URL.revokeObjectURL(url); resolve(""); };
          img.src = url;
          return;
        } catch (_) {}
      }
      resolve("");
    });
  }

  function visibleThumbnailFromBinary(binary, cropWindow) {
    return new Promise(function (resolve) {
      if (!(global.Image && global.document && global.URL && global.Blob)) {
        fileToDataUrl(binary).then(resolve);
        return;
      }
      let url = "";
      try {
        url = global.URL.createObjectURL(new global.Blob([binaryToBytes(binary)], { type: "image/png" }));
        const img = new global.Image();
        img.onload = function () {
          try {
            const sourceW = Number(img.width) || THUMB_PX;
            const sourceH = Number(img.height) || THUMB_PX;
            const windowValue = normalizeCropWindow(cropWindow);
            const sourceX = clamp(Math.floor(windowValue.l * sourceW), 0, Math.max(0, sourceW - 1));
            const sourceY = clamp(Math.floor(windowValue.t * sourceH), 0, Math.max(0, sourceH - 1));
            const sourceRight = clamp(Math.ceil(windowValue.r * sourceW), sourceX + 1, sourceW);
            const sourceBottom = clamp(Math.ceil(windowValue.b * sourceH), sourceY + 1, sourceH);
            const sourceCropW = sourceRight - sourceX;
            const sourceCropH = sourceBottom - sourceY;
            const canvas = global.document.createElement("canvas");
            canvas.width = THUMB_PX;
            canvas.height = THUMB_PX;
            canvas.getContext("2d").drawImage(img, sourceX, sourceY, sourceCropW, sourceCropH, 0, 0, THUMB_PX, THUMB_PX);
            const dataUrl = canvas.toDataURL("image/png");
            global.URL.revokeObjectURL(url);
            resolve(dataUrl);
          } catch (_) {
            try { global.URL.revokeObjectURL(url); } catch (__) {}
            fileToDataUrl(binary).then(resolve);
          }
        };
        img.onerror = function () {
          try { global.URL.revokeObjectURL(url); } catch (_) {}
          fileToDataUrl(binary).then(resolve);
        };
        img.src = url;
      } catch (_) {
        try { if (url) global.URL.revokeObjectURL(url); } catch (__) {}
        fileToDataUrl(binary).then(resolve);
      }
    });
  }

  // Export the full uncropped image once: return both a content fingerprint
  // (for same-source grouping) and a base64 thumbnail for the panel.
  async function contentFingerprintAndThumb(shape, scratch) {
    const path = tempPath("panel");
    try {
      const exported = exportUncroppedPreview(shape, scratch, path, false, false, THUMB_PX);
      if (!exported.ok) {
        return { fp: "", dataUrl: "", dHash: "", aHash: "", cropWindow: cropWindowForShape(shape, 0, 0), w: 0, h: 0, nw: 0, nh: 0, aspect: 0 };
      }
      const binary = fileSystem().readAsBinaryString(path);
      const fp = fnv1a(binary);
      // nw/nh = full uncropped image size before normalization.
      const nw = num(exported.fullW);
      const nh = num(exported.fullH);
      const cropWindow = cropWindowForShape(shape, nw, nh);
      const dataUrl = await visibleThumbnailFromBinary(binary, cropWindow);
      const hashes = await perceptualHashesFromBinary(binary);
      const aspect = nw > 0 && nh > 0 ? Math.round(1000 * nw / nh) : 0;
      return { fp: fp, dataUrl: dataUrl, dHash: hashes.dHash, aHash: hashes.aHash, cropWindow: cropWindow, w: THUMB_PX, h: THUMB_PX, nw: nw, nh: nh, aspect: aspect };
    } finally { removeFile(path); }
  }

  // =====================================================================
  // Batch thumbnail/fingerprint rendering. WPS JSAPI has no Shape.Export,
  // so the old path exported one scratch slide PER picture (each export and
  // each clipboard round-trip crosses the JSAPI bridge). Instead we paste up
  // to four pictures into a 2x2 grid on ONE scratch slide and export the grid
  // once; cells are then decoded in JS (dHash + fingerprint + thumbnail).
  // The smaller atomic batch trades some total throughput for a responsive
  // WPS window, which is the correct trade-off for an interactive taskpane.
  // =====================================================================
  // Responsiveness is more important than minimizing the number of exports in
  // this interactive panel. Four pictures per atomic WPS batch keeps the
  // clipboard/scratch-slide work serial while returning control to WPS often.
  const BATCH_COLS = 2;
  const BATCH_CELL = THUMB_PX;
  const BATCH_GRID = BATCH_COLS * BATCH_CELL;

  function ensureBatchSlide(scratch) {
    const slide = scratch.ensure();
    try {
      if (slide.__batchSized) return slide;
    } catch (_) {}
    try {
      const pageSetup = presentationOf(slide).PageSetup;
      if (pageSetup) {
        pageSetup.SlideWidth = BATCH_GRID;
        pageSetup.SlideHeight = BATCH_GRID;
      }
    } catch (_) { return null; }
    try { slide.__batchSized = true; } catch (_) {}
    return slide;
  }

  function decodeBatchCells(binary, count, cropWindows) {
    return new Promise(function (resolve) {
      if (!(global.Image && global.document && global.URL && global.Blob)) { resolve(null); return; }
      try {
        const url = global.URL.createObjectURL(new global.Blob([binaryToBytes(binary)], { type: "image/png" }));
        const img = new global.Image();
        img.onload = function () {
          try {
            const cols = BATCH_COLS;
            const cellPx = Math.round(img.width / cols);
            if (!(cellPx > 0)) { global.URL.revokeObjectURL(url); resolve(null); return; }
            const full = global.document.createElement("canvas");
            full.width = img.width;
            full.height = img.height;
            full.getContext("2d").drawImage(img, 0, 0);
            global.URL.revokeObjectURL(url);
            const out = new Array(count).fill(null);
            for (let i = 0; i < count; i += 1) {
              const col = i % cols;
              const row = Math.floor(i / cols);
              try {
                const sourceCell = global.document.createElement("canvas");
                sourceCell.width = BATCH_CELL;
                sourceCell.height = BATCH_CELL;
                const sourceContext = sourceCell.getContext("2d");
                sourceContext.drawImage(full, col * cellPx, row * cellPx, cellPx, cellPx, 0, 0, BATCH_CELL, BATCH_CELL);
                const tiny = global.document.createElement("canvas");
                tiny.width = 9;
                tiny.height = 8;
                const tctx = tiny.getContext("2d");
                tctx.drawImage(sourceCell, 0, 0, 9, 8);
                const perceptualPixels = tctx.getImageData(0, 0, 9, 8).data;
                const dHash = dHashFromPixels(perceptualPixels, 9);
                const aHash = aHashFromPixels(perceptualPixels, 9);
                const pixels = sourceContext.getImageData(0, 0, BATCH_CELL, BATCH_CELL).data;
                // Fingerprint over raw RGBA: encoder-independent and stable
                // across WPS versions/re-encodes (old bytes-based fingerprint
                // is kept only as a fallback for linked-cache reuse).
                const fp = fnvBytes(pixels);
                const visible = global.document.createElement("canvas");
                visible.width = BATCH_CELL;
                visible.height = BATCH_CELL;
                const visibleContext = visible.getContext("2d");
                const windowValue = normalizeCropWindow(cropWindows && cropWindows[i]);
                const localX = clamp(Math.floor(windowValue.l * cellPx), 0, Math.max(0, cellPx - 1));
                const localY = clamp(Math.floor(windowValue.t * cellPx), 0, Math.max(0, cellPx - 1));
                const localRight = clamp(Math.ceil(windowValue.r * cellPx), localX + 1, cellPx);
                const localBottom = clamp(Math.ceil(windowValue.b * cellPx), localY + 1, cellPx);
                const sourceX = col * cellPx + localX;
                const sourceY = row * cellPx + localY;
                const sourceW = localRight - localX;
                const sourceH = localBottom - localY;
                visibleContext.drawImage(full, sourceX, sourceY, sourceW, sourceH, 0, 0, BATCH_CELL, BATCH_CELL);
                const dataUrl = visible.toDataURL("image/png");
                out[i] = { fp: fp, dHash: dHash, aHash: aHash, dataUrl: dataUrl };
              } catch (_) { out[i] = null; }
            }
            resolve(out);
          } catch (_) { global.URL.revokeObjectURL(url); resolve(null); }
        };
        img.onerror = function () { global.URL.revokeObjectURL(url); resolve(null); };
        img.src = url;
      } catch (_) { resolve(null); }
    });
  }

  async function renderThumbnailsBatch(items, scratch, onChunk) {
    const results = new Array(items.length).fill(null);
    if (!items.length) return results;
    const slide = ensureBatchSlide(scratch);
    if (!slide) return results;
    const shapes = slide.Shapes;
    const naturals = new Array(items.length).fill(null);
    const CHUNK = BATCH_COLS * BATCH_COLS;
    for (let start = 0; start < items.length; start += CHUNK) {
      const end = Math.min(start + CHUNK, items.length);
      scratch.clear();
      let pasted = 0;
      for (let i = start; i < end; i += 1) {
        const shape = items[i].shape;
        const idx = i - start;
        try {
          const before = Number(shapes.Count) || 0;
          let after = before;
          let result = null;
          for (let attempt = 0; attempt < 3 && after <= before; attempt += 1) {
            try { shape.Copy(); } catch (_) {}
            try { result = shapes.Paste(); } catch (_) {}
            after = Number(shapes.Count) || 0;
          }
          if (after <= before) continue;
          const pastedShape = asShape(result) || shapes.Item(after);
          const pf = pastedShape.PictureFormat;
          if (!pf) continue;
          try { pastedShape.Rotation = 0; } catch (_) {}
          if (isTrue(pastedShape.HorizontalFlip)) pastedShape.Flip(MsoFlipHorizontal);
          if (isTrue(pastedShape.VerticalFlip)) pastedShape.Flip(MsoFlipVertical);
          try { pf.CropLeft = 0; } catch (_) {}
          try { pf.CropRight = 0; } catch (_) {}
          try { pf.CropTop = 0; } catch (_) {}
          try { pf.CropBottom = 0; } catch (_) {}
          try { pastedShape.Line.Visible = MsoFalse; } catch (_) {}
          try { pastedShape.Shadow.Visible = MsoFalse; } catch (_) {}
          try { pastedShape.LockAspectRatio = MsoFalse; } catch (_) {}
          const nw = num(pastedShape.Width);
          const nh = num(pastedShape.Height);
          if (!(nw > 0 && nh > 0)) continue;
          naturals[i] = { nw: nw, nh: nh, cropWindow: cropWindowForShape(shape, nw, nh) };
          pastedShape.Left = (idx % BATCH_COLS) * BATCH_CELL;
          pastedShape.Top = Math.floor(idx / BATCH_COLS) * BATCH_CELL;
          pastedShape.Width = BATCH_CELL;
          pastedShape.Height = BATCH_CELL;
          pasted += 1;
        } catch (_) {}
      }
      if (!pasted) {
        if (onChunk) { try { await onChunk(end, items.length, results); } catch (_) {} }
        await yieldUI();
        continue;
      }
      const path = tempPath("batch");
      removeFile(path);
      try {
        slide.Export(path, "PNG", BATCH_GRID, BATCH_GRID);
        if (!fileExists(path)) {
          if (onChunk) { try { await onChunk(end, items.length, results); } catch (_) {} }
          await yieldUI();
          continue;
        }
        const binary = fileSystem().readAsBinaryString(path);
        const decoded = await decodeBatchCells(binary, end - start, naturals.slice(start, end).map(function (natural) {
          return natural && natural.cropWindow ? natural.cropWindow : fullCropWindow(false);
        }));
        if (!decoded) {
          if (onChunk) { try { await onChunk(end, items.length, results); } catch (_) {} }
          await yieldUI();
          continue;
        }
        for (let i = start; i < end; i += 1) {
          const info = decoded[i - start];
          const nat = naturals[i];
          if (!info || !nat) continue;
          const nw = nat ? nat.nw : 0;
          const nh = nat ? nat.nh : 0;
          results[i] = {
            fp: info.fp,
            dataUrl: info.dataUrl,
            dHash: info.dHash,
            aHash: info.aHash,
            cropWindow: nat.cropWindow,
            w: BATCH_CELL,
            h: BATCH_CELL,
            nw: nw,
            nh: nh,
            aspect: nw > 0 && nh > 0 ? Math.round(1000 * nw / nh) : 0
          };
        }
      } finally { removeFile(path); }
      if (onChunk) { try { await onChunk(end, items.length, results); } catch (_) {} }
      await yieldUI();
    }
    return results;
  }
  async function linkStateOf(group) {
    const linked = group.instances.filter(function (i) { return i.linked && i.src; });
    if (!linked.length) return "none";
    const srcSet = {};
    let src = "";
    for (let i = 0; i < linked.length; i += 1) {
      srcSet[normalizePath(linked[i].src)] = true;
      src = linked[i].src;
    }
    if (Object.keys(srcSet).length > 1) return "mixed";
    if (!fileExists(src)) return "missing";
    let current = "";
    const cached = linked.every(function (inst) { return !!inst.currentFileFp; });
    if (cached) current = linked[0].currentFileFp;
    else {
      try { current = fingerprintFile(src); } catch (_) { return "unreadable"; }
    }
    return current === linked[0].fileFp ? "linked" : "modified";
  }

  // Collect every picture in the active presentation, group by content
  // fingerprint (same source image), and report link status per group.
  // onProgress(done, total) is invoked between instances.
  function documentKey(presentation) {
    try {
      const full = String(presentation.FullName || "");
      if (full) {
        // Keep the stable file identity, but include the cheap structural
        // shape counts so a cache from a document that gained/lost objects is
        // never presented as current.  This avoids exporting thumbnails just
        // to validate the cache while still catching the common edit case.
        let key = "file:" + normalizePath(full);
        try {
          const slides = presentation.Slides;
          const count = Number(slides.Count) || 0;
          key += "|slides:" + count;
          for (let i = 1; i <= count; i += 1) {
            const slide = slides.Item(i);
            key += ":" + (Number(slide.Shapes.Count) || 0);
            try {
              const layout = slide.CustomLayout;
              key += "@" + String(layout && layout.Name || "");
            } catch (_) {}
          }
          try {
            const master = presentation.SlideMaster;
            key += "|master:" + (master && master.Shapes ? (Number(master.Shapes.Count) || 0) : 0);
            const layouts = master && master.CustomLayouts;
            const layoutCount = layouts ? (Number(layouts.Count) || 0) : 0;
            key += "|layouts:" + layoutCount;
            for (let i = 1; i <= layoutCount; i += 1) {
              const layout = layouts.Item(i);
              key += ":" + (layout && layout.Shapes ? (Number(layout.Shapes.Count) || 0) : 0);
            }
          } catch (_) {}
        } catch (_) {}
        return key;
      }
    } catch (_) {}
    let key = "unsaved:";
    try { key += String(presentation.Name || "") + ":"; } catch (_) {}
    try { key += Number(presentation.Slides.Count) || 0; } catch (_) { key += "?"; }
    try {
      const sc = Number(presentation.Slides.Count) || 0;
      for (let i = 1; i <= sc; i += 1) {
        const slide = presentation.Slides.Item(i);
        key += ":" + (Number(slide.Shapes.Count) || 0);
        try {
          const layout = slide.CustomLayout;
          key += "@" + String(layout && layout.Name || "");
        } catch (_) {}
      }
    } catch (_) {}
    try {
      const master = presentation.SlideMaster;
      key += "|master:" + (master && master.Shapes ? (Number(master.Shapes.Count) || 0) : 0);
      const layouts = master && master.CustomLayouts;
      const layoutCount = layouts ? (Number(layouts.Count) || 0) : 0;
      key += "|layouts:" + layoutCount;
      for (let i = 1; i <= layoutCount; i += 1) {
        const layout = layouts.Item(i);
        key += ":" + (layout && layout.Shapes ? (Number(layout.Shapes.Count) || 0) : 0);
      }
    } catch (_) {}
    return key;
  }

  function assertDocument(docKey) {
    if (!docKey) return;
    const current = documentKey(activePresentation());
    if (current !== docKey) {
      throw new Error("当前演示文稿已变化，请先刷新图片清单再操作。");
    }
  }

  // =====================================================================
  // Picture-panel inventory cache
  // =====================================================================
  // A task pane can be destroyed and recreated when the user closes and
  // reopens it.  Keeping only the DOM state therefore does not solve the
  // repeated full-deck scan.  The cache below has two tiers:
  //   1) live in-memory data for hash/view switches in one pane context;
  //   2) a small JSON snapshot in the WPS temp directory for a new pane
  //      context.  The snapshot contains no COM objects; those are rebound
  //      to the current presentation by slide/layout/shape index on read.
  // Only a completed scan is persisted.  Any mutating operation calls
  // invalidatePanelInventoryCache(), and the toolbar refresh uses force=true
  // to discard the snapshot before scanning again.
  const PANEL_CACHE_VERSION = 2;
  const PANEL_CACHE_FILE_NAME = "PictureReplaceTools_panel_inventory_v2.json";
  const PANEL_CACHE_BUSY_FILE_NAME = "PictureReplaceTools_panel_inventory_v2.busy.json";
  const PANEL_CACHE_MAX_CHARS = 12 * 1024 * 1024;
  const PANEL_CACHE_BUSY_HEARTBEAT_MS = 2000;
  const PANEL_CACHE_BUSY_STALE_MS = 15000;
  let panelInventoryMemoryCache = null;
  let panelInventoryDiskLoaded = false;
  let panelInventoryDiskEnvelope = null;
  let panelInventoryScan = null;
  let panelInventoryEpoch = 0;
  let panelInventoryPreloadTimer = null;
  let panelInventoryPreloadPromise = null;
  let panelInventoryPreloadReadyKey = "";
  let panelInventoryPreloadEventsBound = false;
  // Background preloads are only worth their cost (full deck walk + thumbnail
  // rendering through clipboard Copy/Paste) in sessions where the user has
  // actually opened the picture panel.  Until then every preload trigger is
  // ignored and the pane simply scans in the foreground on first open.
  let panelInventoryPanelUsed = false;

  // Lightweight perf trace: events slower than the threshold are appended to
  // %TEMP%\dongsidaye_perf.log so stutter reports can be correlated with real
  // measurements instead of guesswork.
  const PERF_TRACE_LIMIT_MS = 120;
  function perfTraceLog(label, detail) {
    try {
      const fs = fileSystem();
      let base = "";
      try { base = fs.tmpdir(); } catch (_) {}
      if (!base) return;
      if (!/[\\/]$/.test(base)) base += "\\";
      const path = base + "dongsidaye_perf.log";
      let prev = "";
      try { prev = String(fs.readAsBinaryString(path) || ""); } catch (_) {}
      if (prev.length > 180000) prev = prev.slice(-90000);
      const line = new Date().toISOString() + " " + label + (detail ? " " + detail : "") + "\n";
      try { fs.writeAsBinaryString(path, prev + line); } catch (_) { try { fs.WriteFile(path, prev + line); } catch (_) {} }
    } catch (_) {}
  }
  function perfTraceTime(label, startedAt) {
    const ms = Date.now() - startedAt;
    if (ms >= PERF_TRACE_LIMIT_MS) perfTraceLog(label, ms + "ms");
    return ms;
  }

  function panelInventoryCachePath() {
    let base = "";
    try { base = String(fileSystem().tmpdir() || ""); } catch (_) {}
    if (!base) {
      try { base = String(application().Env.GetTempPath() || ""); } catch (_) {}
    }
    if (!base) return "";
    if (!/[\\/]$/.test(base)) base += "\\";
    return base + PANEL_CACHE_FILE_NAME;
  }

  function panelInventoryBusyPath() {
    const cachePath = panelInventoryCachePath();
    if (!cachePath) return "";
    return cachePath.slice(0, Math.max(0, cachePath.length - PANEL_CACHE_FILE_NAME.length)) + PANEL_CACHE_BUSY_FILE_NAME;
  }

  function panelInventoryReadText(path) {
    if (!path) return "";
    try {
      const fs = fileSystem();
      if (hasMethod(fs, "ReadFile")) return String(fs.ReadFile(path) || "");
      if (hasMethod(fs, "readAsBinaryString")) return String(fs.readAsBinaryString(path) || "");
    } catch (_) {}
    return "";
  }

  function panelInventoryWriteText(path, text) {
    if (!path) return false;
    try {
      const fs = fileSystem();
      // WriteFile is the text-oriented WPS API and preserves Chinese shape
      // names.  Binary fallback keeps compatibility with older builds.
      if (hasMethod(fs, "WriteFile")) { fs.WriteFile(path, text); return true; }
      if (hasMethod(fs, "writeAsBinaryString")) { fs.writeAsBinaryString(path, text); return true; }
    } catch (_) {}
    return false;
  }

  function panelInventoryReadBusy() {
    const path = panelInventoryBusyPath();
    const raw = panelInventoryReadText(path);
    if (!raw) return null;
    try {
      const marker = JSON.parse(raw);
      if (!marker || Number(marker.version) !== PANEL_CACHE_VERSION || !marker.docKey || !marker.token) return null;
      return marker;
    } catch (_) { return null; }
  }

  function panelInventoryWriteBusy(docKey, token) {
    const now = Date.now();
    const payload = JSON.stringify({ version: PANEL_CACHE_VERSION, docKey: String(docKey || ""), token: String(token || ""), startedAt: now, heartbeatAt: now });
    return panelInventoryWriteText(panelInventoryBusyPath(), payload);
  }

  function panelInventoryTouchBusy(docKey, token) {
    const current = panelInventoryReadBusy();
    if (!current || String(current.docKey || "") !== String(docKey || "") || String(current.token || "") !== String(token || "")) return false;
    current.heartbeatAt = Date.now();
    return panelInventoryWriteText(panelInventoryBusyPath(), JSON.stringify(current));
  }

  function panelInventoryBusyAge(marker) {
    if (!marker) return Infinity;
    const heartbeatAt = Number(marker.heartbeatAt || marker.startedAt || 0);
    return heartbeatAt > 0 ? Math.max(0, Date.now() - heartbeatAt) : Infinity;
  }

  function panelInventoryClearBusy(docKey, token) {
    const current = panelInventoryReadBusy();
    if (!current || String(current.docKey) !== String(docKey || "") || String(current.token) !== String(token || "")) return;
    removeFile(panelInventoryBusyPath());
  }

  function panelInventoryRemovePersisted() {
    const path = panelInventoryCachePath();
    if (path) removeFile(path);
    panelInventoryDiskLoaded = true;
    panelInventoryDiskEnvelope = null;
  }

  async function panelInventoryWaitForExternalScan(docKey, presentation, timeoutMs) {
    const deadline = Date.now() + Math.max(1000, Number(timeoutMs) || 90000);
    while (Date.now() < deadline) {
      // A different JS context may have written the completed snapshot since
      // the last read.  Drop the local “loaded” bit before each poll.
      panelInventoryDiskLoaded = false;
      panelInventoryDiskEnvelope = null;
      const envelope = panelInventoryLoadPersisted();
      if (envelope && String(envelope.docKey || "") === String(docKey || "")) {
        const hydrated = panelInventoryHydrate(envelope, presentation);
        if (hydrated) {
          panelInventoryMemoryCache = { docKey: docKey, presentation: presentation, result: hydrated, snapshot: envelope };
          return hydrated;
        }
      }
      const busy = panelInventoryReadBusy();
      if (!busy || String(busy.docKey || "") !== String(docKey || "")) return null;
      // The owner refreshes heartbeatAt while a long scan is alive.  A marker
      // without a recent heartbeat is treated as a crashed/stale context so
      // reopening the panel cannot wait indefinitely after WPS was killed.
      if (panelInventoryBusyAge(busy) > PANEL_CACHE_BUSY_STALE_MS) return null;
      await yieldUI(120);
    }
    return null;
  }

  async function panelInventoryWaitForBusyClear(marker, timeoutMs) {
    if (!marker || !marker.token) return true;
    const deadline = Date.now() + Math.max(1000, Number(timeoutMs) || 90000);
    while (Date.now() < deadline) {
      const current = panelInventoryReadBusy();
      if (!current || String(current.token || "") !== String(marker.token || "")) return true;
      if (panelInventoryBusyAge(current) > PANEL_CACHE_BUSY_STALE_MS) return false;
      await yieldUI(120);
    }
    return false;
  }

  function panelInventoryLoadPersisted() {
    if (panelInventoryDiskLoaded) return panelInventoryDiskEnvelope;
    panelInventoryDiskLoaded = true;
    const path = panelInventoryCachePath();
    const raw = panelInventoryReadText(path);
    if (!raw) return null;
    try {
      const envelope = JSON.parse(raw);
      if (!envelope || Number(envelope.version) !== PANEL_CACHE_VERSION || !Array.isArray(envelope.groups)) return null;
      panelInventoryDiskEnvelope = envelope;
      return envelope;
    } catch (_) { return null; }
  }

  function panelInventoryNumber(value) {
    const n = Number(value);
    return isFinite(n) ? n : 0;
  }

  function panelInventoryInstanceSnapshot(instance) {
    const keys = [
      "uid", "slideIndex", "shapeIndex", "kind", "layoutIndex", "layoutName",
      "shapeName", "visible", "left", "top", "width", "height", "overlap",
      "name", "zone", "hasCrop", "linked", "src", "fileFp", "currentFileFp",
      "aspect", "userAlt", "thumb", "thumbW", "thumbH", "cropWindow"
    ];
    const out = {};
    for (let i = 0; i < keys.length; i += 1) {
      const key = keys[i];
      out[key] = instance && instance[key] !== undefined ? instance[key] : "";
    }
    out.slideIndex = Number(out.slideIndex) || 0;
    out.shapeIndex = Number(out.shapeIndex) || 0;
    out.layoutIndex = Number(out.layoutIndex) || 0;
    out.appliedPages = Array.isArray(instance && instance.appliedPages)
      ? instance.appliedPages.map(function (page) { return Number(page) || 0; }).filter(function (page) { return page > 0; })
      : [];
    try { out.altFingerprint = fnv1a(String(instance && instance.shape && instance.shape.AlternativeText || "")); } catch (_) { out.altFingerprint = ""; }
    try { out.shapeType = Number(instance && instance.shape && instance.shape.Type); } catch (_) { out.shapeType = 0; }
    try {
      const pf = instance && instance.shape && instance.shape.PictureFormat;
      out.cropLeft = panelInventoryNumber(pf && pf.CropLeft);
      out.cropRight = panelInventoryNumber(pf && pf.CropRight);
      out.cropTop = panelInventoryNumber(pf && pf.CropTop);
      out.cropBottom = panelInventoryNumber(pf && pf.CropBottom);
    } catch (_) {
      out.cropLeft = 0; out.cropRight = 0; out.cropTop = 0; out.cropBottom = 0;
    }
    return out;
  }

  function panelInventorySnapshot(result) {
    return {
      version: PANEL_CACHE_VERSION,
      createdAt: Date.now(),
      docKey: String(result && result.docKey || ""),
      total: Number(result && result.total) || 0,
      slideCount: Number(result && result.slideCount) || 0,
      groups: (result && Array.isArray(result.groups) ? result.groups : []).map(function (group) {
        return {
          key: String(group && group.key || ""),
          name: String(group && group.name || ""),
          src: String(group && group.src || ""),
          fileFp: String(group && group.fileFp || ""),
          contentFp: String(group && group.contentFp || ""),
          dHash: String(group && group.dHash || ""),
          aHash: String(group && group.aHash || ""),
          aspect: Number(group && group.aspect) || 0,
          linkState: String(group && group.linkState || ""),
          instances: (group && Array.isArray(group.instances) ? group.instances : []).map(panelInventoryInstanceSnapshot)
        };
      })
    };
  }

  function panelInventoryPersist(result, epoch) {
    if (!result || epoch !== panelInventoryEpoch) return false;
    const snapshot = panelInventorySnapshot(result);
    if (!snapshot.docKey) return false;
    let payload = "";
    try { payload = JSON.stringify(snapshot); } catch (_) { return false; }
    // A very large deck should still work from the in-memory tier; refusing a
    // huge temp-file write avoids turning a background preload into another
    // host freeze or exhausting the WPS FileSystem quota.
    if (!payload || payload.length > PANEL_CACHE_MAX_CHARS) return false;
    if (!panelInventoryWriteText(panelInventoryCachePath(), payload)) return false;
    panelInventoryDiskLoaded = true;
    panelInventoryDiskEnvelope = snapshot;
    return true;
  }

  function panelInventoryResolvedShape(presentation, instance) {
    if (!presentation || !instance) return null;
    const kind = String(instance.kind || "slide");
    try {
      if (kind === "master") {
        const master = presentation.SlideMaster;
        return master && master.Shapes ? master.Shapes.Item(Number(instance.shapeIndex)) : null;
      }
      if (kind === "layout") {
        const layouts = presentation.SlideMaster && presentation.SlideMaster.CustomLayouts;
        const layout = layouts && layouts.Item(Number(instance.layoutIndex));
        return layout && layout.Shapes ? layout.Shapes.Item(Number(instance.shapeIndex)) : null;
      }
      const slide = presentation.Slides.Item(Number(instance.slideIndex));
      return slide && slide.Shapes ? slide.Shapes.Item(Number(instance.shapeIndex)) : null;
    } catch (_) { return null; }
  }

  function panelInventoryNear(a, b, tolerance) {
    return Math.abs(panelInventoryNumber(a) - panelInventoryNumber(b)) <= (tolerance || 0.75);
  }

  function panelInventoryVerifyShape(shape, cached) {
    if (!shape || !cached || !isPicture(shape)) return false;
    if (cached.shapeName && String(shape.Name || "") !== String(cached.shapeName)) return false;
    if (cached.altFingerprint) {
      let currentAlt = "";
      try { currentAlt = String(shape.AlternativeText || ""); } catch (_) {}
      if (fnv1a(currentAlt) !== String(cached.altFingerprint)) return false;
    }
    if (!panelInventoryNear(shape.Left, cached.left) || !panelInventoryNear(shape.Top, cached.top) ||
        !panelInventoryNear(shape.Width, cached.width) || !panelInventoryNear(shape.Height, cached.height)) return false;
    if (cached.visible !== "" && isTrue(shape.Visible) !== !!cached.visible) return false;
    if (cached.shapeType && Number(shape.Type) !== Number(cached.shapeType)) return false;
    try {
      const pf = shape.PictureFormat;
      if (!panelInventoryNear(pf && pf.CropLeft, cached.cropLeft) ||
          !panelInventoryNear(pf && pf.CropRight, cached.cropRight) ||
          !panelInventoryNear(pf && pf.CropTop, cached.cropTop) ||
          !panelInventoryNear(pf && pf.CropBottom, cached.cropBottom)) return false;
    } catch (_) {}
    return true;
  }

  function panelInventoryHydrate(snapshot, presentation) {
    if (!snapshot || Number(snapshot.version) !== PANEL_CACHE_VERSION || !presentation) return null;
    if (String(snapshot.docKey || "") !== documentKey(presentation)) return null;
    const groups = [];
    let total = 0;
    for (let g = 0; g < snapshot.groups.length; g += 1) {
      const cachedGroup = snapshot.groups[g] || {};
      const instances = [];
      const cachedInstances = Array.isArray(cachedGroup.instances) ? cachedGroup.instances : [];
      for (let i = 0; i < cachedInstances.length; i += 1) {
        const cached = cachedInstances[i];
        const shape = panelInventoryResolvedShape(presentation, cached);
        if (!panelInventoryVerifyShape(shape, cached)) return null;
        const instance = Object.assign({}, cached);
        instance.appliedPages = Array.isArray(cached.appliedPages) ? cached.appliedPages.slice() : [];
        instance.shape = shape;
        delete instance.altFingerprint;
        delete instance.shapeType;
        instances.push(instance);
        total += 1;
      }
      const group = Object.assign({}, cachedGroup);
      group.instances = instances;
      const hasLinkedSource = instances.some(function (instance) { return instance && instance.linked && instance.src; });
      if (!hasLinkedSource) group.linkState = "none";
      else if (!group.linkState || group.linkState === "checking") group.linkState = "unchecked";
      groups.push(group);
    }
    return {
      groups: groups,
      total: Number(snapshot.total) || total,
      slideCount: Number(snapshot.slideCount) || 0,
      docKey: String(snapshot.docKey || ""),
      fallbackCount: 0
    };
  }

  function panelInventoryResultView(result, cacheHit, source) {
    const view = {
      groups: result && result.groups ? result.groups : [],
      total: Number(result && result.total) || 0,
      slideCount: Number(result && result.slideCount) || 0,
      docKey: String(result && result.docKey || ""),
      fallbackCount: Number(result && result.fallbackCount) || 0,
      cacheHit: cacheHit === true,
      cacheSource: String(source || "")
    };
    return view;
  }

  async function panelInventoryDeliverCached(result, onProgress, onPartial, source) {
    await yieldUI();
    if (onPartial) {
      try {
        onPartial({
          phase: "cache",
          groups: result.groups || [],
          total: result.total || 0,
          slideCount: result.slideCount || 0,
          docKey: result.docKey || "",
          complete: false,
          cached: true
        });
      } catch (_) {}
    }
    if (onProgress) { try { onProgress(result.total || 0, result.total || 0, "读取图片清单缓存"); } catch (_) {} }
    await yieldUI();
    return panelInventoryResultView(result, true, source);
  }

  function invalidatePanelInventoryCache() {
    panelInventoryEpoch += 1;
    panelInventoryMemoryCache = null;
    panelInventoryPreloadReadyKey = "";
    panelInventoryRemovePersisted();
    // Do not remove another JS context's busy marker here.  A mutation can
    // invalidate the eventual result while that context is still scanning;
    // the scan owner clears its own token in panelInventoryRunScan.finally.
    return true;
  }

  async function panelInventoryRunScan(presentation, docKey, onProgress, onPartial) {
    const epoch = panelInventoryEpoch;
    const scanStartedAt = Date.now();
    perfTraceLog("scan.start", "slides=" + (function () { try { return Number(presentation.Slides.Count) || 0; } catch (_) { return "?"; } }()));
    const busyToken = String(Date.now()) + "-" + Math.random().toString(16).slice(2);
    panelInventoryWriteBusy(docKey, busyToken);
    let heartbeatTimer = null;
    try {
      heartbeatTimer = setInterval(function () {
        try { panelInventoryTouchBusy(docKey, busyToken); } catch (_) {}
      }, PANEL_CACHE_BUSY_HEARTBEAT_MS);
    } catch (_) {}
    const promise = (async function () {
      try {
        const result = await collectDeckImages(onProgress, onPartial);
        let current = null;
        try { current = activePresentation(); } catch (_) {}
        if (current && documentKey(current) === docKey && epoch === panelInventoryEpoch) {
          const snapshot = panelInventorySnapshot(result);
          panelInventoryMemoryCache = { docKey: docKey, presentation: current, result: result, snapshot: snapshot };
          panelInventoryPersist(result, epoch);
        }
        return result;
      } finally {
        if (heartbeatTimer) { try { clearInterval(heartbeatTimer); } catch (_) {} }
        panelInventoryClearBusy(docKey, busyToken);
        perfTraceLog("scan.end", (Date.now() - scanStartedAt) + "ms");
      }
    }());
    panelInventoryScan = { docKey: docKey, promise: promise };
    try { return await promise; }
    finally {
      if (panelInventoryScan && panelInventoryScan.promise === promise) panelInventoryScan = null;
    }
  }

  // Cached counterpart used by the task pane.  `options.force` is reserved
  // for the explicit toolbar refresh; ordinary route changes reuse memory or
  // the persisted snapshot and never start another thumbnail export.
  async function collectDeckImagesCached(onProgress, onPartial, options) {
    const force = options === true || !!(options && options.force === true);
    const presentation = activePresentation();
    const docKey = documentKey(presentation);

    // Never run two WPS thumbnail scans concurrently. If the active document
    // changed while a background preload was running, let that scan finish
    // and then retry for the new document.
    if (panelInventoryScan) {
      const pending = panelInventoryScan;
      if (pending.docKey === docKey) {
        try {
          const result = await pending.promise;
          return panelInventoryDeliverCached(result, onProgress, onPartial, "preload");
        } catch (_) {
          // A best-effort background preload may fail on an older WPS build;
          // opening the panel must get one clean retry instead of surfacing
          // the stale rejected promise forever.
          return collectDeckImagesCached(onProgress, onPartial, options);
        }
      }
      try { await pending.promise; } catch (_) {}
      return collectDeckImagesCached(onProgress, onPartial, options);
    }

    if (!force) {
      if (panelInventoryMemoryCache && panelInventoryMemoryCache.docKey === docKey) {
        let liveResult = panelInventoryMemoryCache.result;
        if (panelInventoryMemoryCache.presentation !== presentation) {
          liveResult = panelInventoryHydrate(panelInventoryMemoryCache.snapshot, presentation);
          if (liveResult) {
            panelInventoryMemoryCache = {
              docKey: docKey,
              presentation: presentation,
              result: liveResult,
              snapshot: panelInventoryMemoryCache.snapshot
            };
          }
        }
        if (liveResult) return panelInventoryDeliverCached(liveResult, onProgress, onPartial, "memory");
        panelInventoryMemoryCache = null;
      }

      const persisted = panelInventoryLoadPersisted();
      if (persisted && String(persisted.docKey || "") === docKey) {
        const hydrated = panelInventoryHydrate(persisted, presentation);
        if (hydrated) {
          panelInventoryMemoryCache = { docKey: docKey, presentation: presentation, result: hydrated, snapshot: persisted };
          return panelInventoryDeliverCached(hydrated, onProgress, onPartial, "disk");
        }
        // A stale/partially invalid snapshot is never allowed to poison the
        // next open; the next call will perform a clean scan.
        panelInventoryRemovePersisted();
      }
    } else {
      // If the index/add-in context is already warming this document, wait for
      // that single scan instead of starting a second WPS scratch/export loop.
      // This is the cross-context counterpart to panelInventoryScan above.
      const externalBusy = panelInventoryReadBusy();
      if (externalBusy && String(externalBusy.docKey || "") === docKey) {
        const warmed = await panelInventoryWaitForExternalScan(docKey, presentation, 90000);
        if (warmed) return panelInventoryDeliverCached(warmed, onProgress, onPartial, "preload");
      } else if (externalBusy) {
        await panelInventoryWaitForBusyClear(externalBusy, 90000);
        return collectDeckImagesCached(onProgress, onPartial, options);
      }
      invalidatePanelInventoryCache();
    }

    if (!force) {
      const externalBusy = panelInventoryReadBusy();
      if (externalBusy && String(externalBusy.docKey || "") === docKey) {
        const warmed = await panelInventoryWaitForExternalScan(docKey, presentation, 90000);
        if (warmed) return panelInventoryDeliverCached(warmed, onProgress, onPartial, "preload");
      } else if (externalBusy) {
        await panelInventoryWaitForBusyClear(externalBusy, 90000);
        return collectDeckImagesCached(onProgress, onPartial, options);
      }
    }

    const result = await panelInventoryRunScan(presentation, docKey, onProgress, onPartial);
    return panelInventoryResultView(result, false, "scan");
  }

  function schedulePanelInventoryPreload(delay) {
    if (panelInventoryPreloadPromise) return;
    if (!panelInventoryPanelUsed) return;
    // True debounce: a burst of load/activate/open events collapses into one
    // timer, and the structural documentKey COM walk happens once when the
    // timer fires instead of on every event. A 3s floor keeps the scan away
    // from the moments the user is actively switching windows or documents.
    if (panelInventoryPreloadTimer) { try { clearTimeout(panelInventoryPreloadTimer); } catch (_) {} panelInventoryPreloadTimer = null; }
    const wait = Math.max(3000, Number(delay) || 0);
    panelInventoryPreloadTimer = setTimeout(function () {
      panelInventoryPreloadTimer = null;
      const t0 = Date.now();
      let currentKey = "";
      try { currentKey = documentKey(activePresentation()); } catch (_) {}
      perfTraceTime("preload.docKey", t0);
      if (currentKey && (currentKey === panelInventoryPreloadReadyKey ||
          (panelInventoryMemoryCache && panelInventoryMemoryCache.docKey === currentKey) ||
          (panelInventoryDiskLoaded && panelInventoryDiskEnvelope && String(panelInventoryDiskEnvelope.docKey || "") === currentKey))) return;
      preloadDeckImages();
    }, wait);
  }

  function preloadDeckImages() {
    if (panelInventoryPreloadPromise) return panelInventoryPreloadPromise;
    panelInventoryPreloadPromise = (async function () {
      try {
        const presentation = activePresentation();
        const docKey = documentKey(presentation);
        const result = await collectDeckImagesCached(null, null);
        if (result && result.ok === false) return result;
        panelInventoryPreloadReadyKey = String(result && result.docKey || docKey);
        return { ok: true, docKey: docKey, cacheHit: !!(result && result.cacheHit), result: result };
      } catch (error) {
        // Opening WPS without a presentation is normal; preload is best
        // effort and must never surface a modal error or block the ribbon.
        return { ok: false, error: String(error && error.message || error) };
      } finally { panelInventoryPreloadPromise = null; }
    }());
    return panelInventoryPreloadPromise;
  }

  function bindPanelInventoryPreloadEvents() {
    if (panelInventoryPreloadEventsBound) return true;
    let apiEvent = null;
    try { apiEvent = application().ApiEvent; } catch (_) {}
    if (!apiEvent || !hasMethod(apiEvent, "AddApiEventListener")) return false;
    const names = ["WindowActivate", "PresentationOpen", "NewPresentation"];
    let bound = false;
    for (let i = 0; i < names.length; i += 1) {
      try {
        apiEvent.AddApiEventListener(names[i], function () { schedulePanelInventoryPreload(650); });
        bound = true;
      } catch (_) {}
    }
    panelInventoryPreloadEventsBound = bound;
    return bound;
  }

  // Resolve each selected instance to a live shape reference ONCE before any
  // replacement. Re-resolving by shapeIndex afterwards is unsafe because
  // replacePictureKeepCrop deletes and re-inserts shapes, shifting indexes.
  function resolveTargets(presentation, instances) {
    const targets = [];
    for (let i = 0; i < instances.length; i += 1) {
      const inst = instances[i];
      let shape = null;
      if (inst.shape) {
        try { void inst.shape.PictureFormat; shape = inst.shape; } catch (_) {}
      }
      if (!shape) shape = resolveShape(presentation, inst.slideIndex, inst.shapeIndex);
      if (!shape) continue;
      targets.push({ shape: shape, inst: inst });
    }
    return targets;
  }

  // Recompute link status (source file hash) for all groups. Kept out of the
  // initial collect so the panel paints instantly; badges update afterwards.
  async function refreshLinkStates(groups, onProgress) {
    const states = [];
    const stateCache = {};
    for (let g = 0; g < groups.length; g += 1) {
      const linked = groups[g].instances.filter(function (i) { return i.linked && i.src; });
      let cacheKey = "";
      if (linked.length) cacheKey = normalizePath(linked[0].src) + "|" + String(linked[0].fileFp || "");
      let state = cacheKey && Object.prototype.hasOwnProperty.call(stateCache, cacheKey) ? stateCache[cacheKey] : "";
      if (!state) {
        state = await linkStateOf(groups[g]);
        if (cacheKey) stateCache[cacheKey] = state;
      }
      try { groups[g].linkState = state; } catch (_) {}
      states.push(state);
      if (onProgress) { try { onProgress(g + 1, groups.length); } catch (_) {} }
      await yieldUI();
    }
    return states;
  }
  // Collect every picture in the active presentation, group by content
  // fingerprint (same source image), and report link status per group.
  // onProgress(done, total) is invoked between instances.
  // One scratch presentation is kept alive for the whole WPS session and
  // reused by every panel refresh (creating/disposing a presentation crosses
  // the JSAPI bridge and is expensive).
  let sharedScratch = null;
  function getSharedScratch() {
    if (!sharedScratch) sharedScratch = createScratchManager();
    return sharedScratch;
  }

  function emptyPanelInfo() {
    return { fp: "", dataUrl: "", dHash: "", aHash: "", cropWindow: null, w: 0, h: 0, nw: 0, nh: 0, aspect: 0 };
  }

  function panelInstanceUid(item) {
    return item.kind === "master"
      ? "M:" + item.shapeIndex
      : (item.kind === "layout"
        ? "L:" + item.layoutName + ":" + item.shapeIndex
        : item.slideIndex + ":" + item.shapeIndex);
  }

  function makePanelInstance(item, info) {
    const safeInfo = info || emptyPanelInfo();
    const isTemplate = item.kind !== "slide";
    return {
      uid: panelInstanceUid(item),
      slideIndex: item.slideIndex,
      shapeIndex: item.shapeIndex,
      shape: item.shape,
      kind: item.kind,
      layoutIndex: item.layoutIndex || 0,
      layoutName: item.layoutName || "",
      appliedPages: item.appliedPages || [],
      shapeName: String(item.shape.Name || ""),
      visible: isTrue(item.shape.Visible),
      left: num(item.shape.Left),
      top: num(item.shape.Top),
      width: num(item.shape.Width),
      height: num(item.shape.Height),
      overlap: false,
      name: item.name,
      zone: isTemplate ? "模板" : zoneOf(item.shape, item.slide),
      hasCrop: hasCropOf(item.shape),
      linked: !!item.meta,
      src: item.meta ? item.meta.src : "",
      fileFp: item.meta ? item.meta.fileFp : "",
      currentFileFp: item._sourceFileFp || "",
      aspect: item.meta ? item.meta.aspect : safeInfo.aspect,
      userAlt: item.meta ? item.meta.userAlt : "",
      thumb: safeInfo.dataUrl,
      thumbW: safeInfo.w,
      thumbH: safeInfo.h,
      cropWindow: normalizeCropWindow(safeInfo.cropWindow || cropWindowForShape(item.shape, 0, 0))
    };
  }

  function applyPanelInfo(item, info) {
    const safeInfo = info || emptyPanelInfo();
    if (!item._instance) item._instance = makePanelInstance(item, safeInfo);
    const instance = item._instance;
    instance.currentFileFp = item._sourceFileFp || "";
    instance.aspect = item.meta ? item.meta.aspect : safeInfo.aspect;
    instance.thumb = safeInfo.dataUrl || "";
    instance.thumbW = safeInfo.w || 0;
    instance.thumbH = safeInfo.h || 0;
    instance.cropWindow = normalizeCropWindow(safeInfo.cropWindow || cropWindowForShape(item.shape, 0, 0));
    return instance;
  }

  function provisionalGroupsFromPending(pending) {
    const groups = [];
    const groupsByKey = new Map();
    for (let i = 0; i < pending.length; i += 1) {
      const item = pending[i];
      const inst = applyPanelInfo(item, item._info || emptyPanelInfo());
      const sourceKey = item.meta && item.meta.contentFp ? "m:" + item.meta.contentFp : "p:" + inst.uid;
      let group = groupsByKey.get(sourceKey);
      if (!group) {
        group = {
          key: sourceKey,
          name: item.name,
          src: item.meta ? item.meta.src : "",
          fileFp: item.meta ? item.meta.fileFp : "",
          contentFp: item.meta ? item.meta.contentFp : "",
          dHash: "",
          aHash: "",
          aspect: item.meta ? item.meta.aspect : 0,
          // Source-file status is intentionally deferred until the user
          // clicks “更新已修改链接”; do not present that idle state as an
          // active background check.
          linkState: "none",
          instances: []
        };
        groupsByKey.set(sourceKey, group);
        groups.push(group);
      }
      group.instances.push(inst);
      if (inst.linked && inst.src) group.linkState = "unchecked";
    }
    return groups;
  }

  async function collectDeckImages(onProgress, onPartial) {
    const presentation = activePresentation();
    // Let the taskpane paint its loading state before the first WPS bridge
    // traversal. Promise/microtask boundaries are insufficient here; WPS and
    // CEF need a real timer turn to process window messages.
    await yieldUI();
    const docKey = documentKey(presentation);
    const pending = [];
    let lastYieldAt = Date.now();
    async function checkpoint(force) {
      if (force || Date.now() - lastYieldAt >= 12) {
        await yieldUI();
        lastYieldAt = Date.now();
      }
    }
    try {
      const slideCount = Number(presentation.Slides.Count) || 0;
      // per-slide layout name (for "applied to pages" of master/layout pics)
      const slideLayouts = {};
      for (let slideIndex = 1; slideIndex <= slideCount; slideIndex += 1) {
        const slide = presentation.Slides.Item(slideIndex);
        try {
          const lo = slide.CustomLayout;
          slideLayouts[slideIndex] = lo ? String(lo.Name || "") : "";
        } catch (_) { slideLayouts[slideIndex] = ""; }
        const shapeCount = Number(slide.Shapes.Count) || 0;
        for (let shapeIndex = 1; shapeIndex <= shapeCount; shapeIndex += 1) {
          const shape = slide.Shapes.Item(shapeIndex);
          if (!isPicture(shape)) { await checkpoint(false); continue; }
          const meta = parseLink(String(shape.AlternativeText || ""));
          const item = {
            slideIndex: slideIndex,
            shapeIndex: shapeIndex,
            shape: shape,
            slide: slide,
            meta: meta,
            kind: "slide",
            layoutName: "",
            appliedPages: [slideIndex],
            name: meta && meta.name ? meta.name : String(shape.Name || "图片")
          };
          item._instance = makePanelInstance(item, emptyPanelInfo());
          pending.push(item);
          await checkpoint(false);
        }
        if (onProgress) { try { onProgress(slideIndex, slideCount, "扫描幻灯片"); } catch (_) {} }
        await checkpoint(true);
      }
      // SlideMaster pictures (applied to every slide)
      let master = null;
      try { master = presentation.SlideMaster; } catch (_) {}
      const layoutsMeta = [];
      if (master && master.Shapes) {
        const mc = Number(master.Shapes.Count) || 0;
        for (let mi = 1; mi <= mc; mi += 1) {
          const shape = master.Shapes.Item(mi);
          if (!isPicture(shape)) { await checkpoint(false); continue; }
          const meta = parseLink(String(shape.AlternativeText || ""));
          const item = {
            slideIndex: 0,
            shapeIndex: mi,
            shape: shape,
            slide: null,
            meta: meta,
            kind: "master",
            layoutName: "",
            appliedPages: slideCount > 0 ? Array.from({ length: slideCount }, function (_, k) { return k + 1; }) : [],
            name: meta && meta.name ? meta.name : String(shape.Name || "母版图片")
          };
          item._instance = makePanelInstance(item, emptyPanelInfo());
          pending.push(item);
          await checkpoint(false);
        }
        // layout pictures (applied to slides using that layout)
        try {
          const cs = master.CustomLayouts;
          if (cs) {
            const lc = Number(cs.Count) || 0;
            for (let li = 1; li <= lc; li += 1) {
              const lo = cs.Item(li);
              let loName = "";
              try { loName = String(lo.Name || ""); } catch (_) {}
              layoutsMeta.push({ index: li, name: loName });
              if (lo.Shapes) {
                const lsc = Number(lo.Shapes.Count) || 0;
                for (let lsi = 1; lsi <= lsc; lsi += 1) {
                  const shape = lo.Shapes.Item(lsi);
                  if (!isPicture(shape)) { await checkpoint(false); continue; }
                  const meta = parseLink(String(shape.AlternativeText || ""));
                  const applied = [];
                  for (let si2 = 1; si2 <= slideCount; si2 += 1) {
                    if (slideLayouts[si2] === loName) applied.push(si2);
                    await checkpoint(false);
                  }
                  const item = {
                    slideIndex: 0,
                    shapeIndex: lsi,
                    shape: shape,
                    slide: null,
                    meta: meta,
                    kind: "layout",
                    layoutName: loName,
                    layoutIndex: li,
                    appliedPages: applied,
                    name: meta && meta.name ? meta.name : String(shape.Name || ("版式图片-" + loName))
                  };
                  item._instance = makePanelInstance(item, emptyPanelInfo());
                  pending.push(item);
                  await checkpoint(false);
                }
              }
              await checkpoint(true);
            }
          }
        } catch (_) {}
      }
      const total = pending.length;
      // Deliver the structural inventory before any Copy/Paste or Slide.Export
      // work. The taskpane can paint a useful list immediately while actions
      // remain disabled until the accurate thumbnail/group scan completes.
      for (let i = 0; i < pending.length; i += 1) {
        if (!pending[i]._instance) pending[i]._instance = makePanelInstance(pending[i], emptyPanelInfo());
      }
      if (onPartial) {
        try {
          onPartial({
            groups: provisionalGroupsFromPending(pending),
            total: total,
            slideCount: slideCount,
            docKey: docKey,
            complete: false
          });
        } catch (_) {}
      }
      // The structural list is now usable. Give CEF a full turn to paint it
      // before creating the scratch presentation or touching the clipboard.
      await yieldUI();
      lastYieldAt = Date.now();
      if (onProgress) { try { onProgress(0, total); } catch (_) {} }
      const groups = [];
      const thumbCache = {};
      const representativeByContent = {};
      // Batch-render every picture that cannot be satisfied by a known
      // fingerprint: grid the shapes onto one scratch slide and export the
      // whole small grid once per four pictures instead of doing one long
      // uninterrupted clipboard/export pass.
      const renderQueue = [];
      for (let i = 0; i < pending.length; i += 1) {
        const item = pending[i];
        const cacheKey = item.meta && item.meta.contentFp ? item.meta.contentFp : "";
        const previewKey = cacheKey ? cacheKey + "|crop:" + rawCropPreviewKey(item.shape) : "";
        // Metadata already identifies repeated linked instances. Render one
        // embedded representative and reuse its small preview; do not read the
        // original source file during panel opening because a single huge file
        // read can freeze WPS for seconds.
        if (previewKey && Object.prototype.hasOwnProperty.call(representativeByContent, previewKey)) {
          item._previewFromIndex = representativeByContent[previewKey];
        } else {
          if (previewKey) representativeByContent[previewKey] = i;
          renderQueue.push(i);
        }
        await checkpoint(true);
      }
      let batchResults = [];
      let fallbackCount = 0;
      if (renderQueue.length) {
        // Creating the hidden scratch presentation is itself a WPS bridge
        // operation, so defer it until the structural list has painted.
        const scratch = getSharedScratch();
        await yieldUI();
        let lastBatchDone = 0;
        batchResults = await renderThumbnailsBatch(renderQueue.map(function (i) { return pending[i]; }), scratch, async function (done, count, liveResults) {
          if (onProgress) { try { onProgress(done, count, "生成缩略图"); } catch (_) {} }
          const updates = [];
          for (let q = lastBatchDone; q < done; q += 1) {
            const info = liveResults[q];
            if (!info) continue;
            const item = pending[renderQueue[q]];
            item._info = info;
            updates.push(applyPanelInfo(item, info));
          }
          lastBatchDone = done;
          if (onPartial && updates.length) {
            try { onPartial({ phase: "thumbnails", updates: updates, done: done, total: count, docKey: docKey, complete: false }); } catch (_) {}
          }
        });
        for (let q = 0; q < renderQueue.length; q += 1) {
          const item = pending[renderQueue[q]];
          let info = batchResults[q];
          if (!info || !info.fp) {
            // Graceful fallback: process one shape per UI turn on older WPS
            // builds without batch canvas decoding.
            await yieldUI();
            info = await contentFingerprintAndThumb(item.shape, scratch);
            fallbackCount += 1;
          }
          const thumbKey = info && info.fp ? info.fp + "|crop:" + rawCropPreviewKey(item.shape) : "";
          if (thumbKey && !thumbCache[thumbKey]) thumbCache[thumbKey] = info;
          item._info = info || emptyPanelInfo();
          const updated = applyPanelInfo(item, item._info);
          if (onPartial && (!batchResults[q] || !batchResults[q].fp)) {
            try { onPartial({ phase: "thumbnails", updates: [updated], done: q + 1, total: renderQueue.length, docKey: docKey, complete: false }); } catch (_) {}
            await yieldUI();
          }
        }
      }
      const reusedUpdates = [];
      for (let i = 0; i < pending.length; i += 1) {
        const item = pending[i];
        if (item._previewFromIndex === undefined) continue;
        const representative = pending[item._previewFromIndex];
        item._info = representative && representative._info ? representative._info : emptyPanelInfo();
        reusedUpdates.push(applyPanelInfo(item, item._info));
        await checkpoint(false);
      }
      if (onPartial && reusedUpdates.length) {
        try { onPartial({ phase: "thumbnails", updates: reusedUpdates, done: total, total: total, docKey: docKey, complete: false }); } catch (_) {}
        await yieldUI();
      }
      for (let i = 0; i < pending.length; i += 1) {
        const item = pending[i];
        if (!item._info) {
          const cacheKey = item.meta && item.meta.contentFp ? item.meta.contentFp : "";
          const thumbKey = cacheKey ? cacheKey + "|crop:" + rawCropPreviewKey(item.shape) : "";
          item._info = (thumbKey && thumbCache[thumbKey]) || emptyPanelInfo();
        }
        const info = item._info;
        const cropWindow = normalizeCropWindow(info.cropWindow || cropWindowForShape(item.shape, 0, 0));
        // Source equality is necessary but no longer sufficient. Every member
        // of a group must also expose a mutually compatible visible window.
        // Pairwise checking prevents transitive chains (full -> top -> bottom)
        // from reconnecting two disjoint crops through a broad middle crop.
        let group = null;
        if (info.fp || info.dHash) {
          for (let gi = 0; gi < groups.length; gi += 1) {
            const candidate = groups[gi];
            const perceptualComparable = !!(info.dHash && info.aHash && candidate.dHash && candidate.aHash);
            const perceptualSource = !!(perceptualComparable &&
              dHashDistance(candidate.dHash, info.dHash) <= 2 &&
              dHashDistance(candidate.aHash, info.aHash) <= 4);
            // The strict fingerprint is only 32-bit for WPS compatibility.
            // When perceptual data exists, cross-check it so an accidental or
            // deliberately constructed FNV collision cannot force a merge.
            const exactSource = !!(info.fp && candidate.contentFp && info.fp === candidate.contentFp &&
              (!perceptualComparable || perceptualSource));
            if (!exactSource && !perceptualSource) continue;
            const windows = candidate._cropWindows || [];
            if (windows.every(function (existing) { return cropWindowCompatible(existing, cropWindow); })) {
              group = candidate;
              break;
            }
          }
        }
        if (!group && (info.fp || info.dHash)) {
          const sourceKey = info.dHash ? "d:" + info.dHash : "f:" + info.fp;
          let uniqueKey = sourceKey + "|crop:" + cropWindowGroupKey(cropWindow);
          let duplicateKey = 2;
          while (groups.some(function (candidate) { return candidate.key === uniqueKey; })) {
            uniqueKey = sourceKey + "|crop:" + cropWindowGroupKey(cropWindow) + "-" + duplicateKey;
            duplicateKey += 1;
          }
          group = {
            key: uniqueKey,
            name: item.name,
            src: item.meta ? item.meta.src : "",
            fileFp: item.meta ? item.meta.fileFp : "",
            contentFp: info.fp,
            dHash: info.dHash,
            aHash: info.aHash,
            aspect: info.aspect,
            instances: [],
            _cropWindows: []
          };
          groups.push(group);
        }
        if (!group) {
          // Failure to fingerprint one shape must never merge every unknown
          // picture into a single destructive batch-replacement target.
          group = { key: "u:" + panelInstanceUid(item), name: item.name || "无法识别", src: "", fileFp: "", contentFp: "", dHash: "", aHash: "", aspect: 0, instances: [], _cropWindows: [] };
          groups.push(group);
        }
        const instance = applyPanelInfo(item, info);
        group.instances.push(instance);
        group._cropWindows.push(cropWindow);
        if (onProgress) { try { onProgress(i + 1, total); } catch (_) {} }
        await checkpoint(false);
      }
      // Flag instances that fully overlap another picture on the same page
      // (same position and size within rounding tolerance) so users can spot
      // leftover/duplicate objects they cannot see in the slide.
      const bySlide = {};
      for (let g = 0; g < groups.length; g += 1) {
        for (let k = 0; k < groups[g].instances.length; k += 1) {
          const inst = groups[g].instances[k];
          (bySlide[inst.slideIndex] = bySlide[inst.slideIndex] || []).push(inst);
        }
      }
      const overlapSlideKeys = Object.keys(bySlide);
      for (let ski = 0; ski < overlapSlideKeys.length; ski += 1) {
        const slideKey = overlapSlideKeys[ski];
        const insts = bySlide[slideKey];
        for (let a = 0; a < insts.length; a += 1) {
          for (let b = a + 1; b < insts.length; b += 1) {
            const x = insts[a];
            const y = insts[b];
            const cx = Math.abs((x.left + x.width / 2) - (y.left + y.width / 2)) < 8;
            const cy = Math.abs((x.top + x.height / 2) - (y.top + y.height / 2)) < 8;
            const sw = Math.abs(x.width - y.width) < 3;
            const sh = Math.abs(x.height - y.height) < 3;
            if (cx && cy && sw && sh) {
              x.overlap = true;
              y.overlap = true;
            }
            await checkpoint(false);
          }
        }
        await checkpoint(true);
      }
      for (let g = 0; g < groups.length; g += 1) {
        const hasLinkedSource = groups[g].instances.some(function (instance) { return instance && instance.linked && instance.src; });
        groups[g].linkState = hasLinkedSource ? "unchecked" : "none";
        delete groups[g]._cropWindows;
      }
      return { groups: groups, total: total, slideCount: slideCount, docKey: docKey, fallbackCount: fallbackCount };
    } finally {
      // shared scratch presentation stays alive across refreshes
    }
  }

  // Replace the given instances with one image file and attach fresh link
  // metadata. Instances are resolved to live shapes before any replacement,
  // so index shifts caused by earlier replacements cannot hit the wrong
  // picture. A custom name from a previous link is preserved.
  async function replaceInstances(instances, imagePath, docKey, onProgress, cancelCheck) {
    assertDocument(docKey);
    const presentation = activePresentation();
    const scratch = createScratchManager();
    const cache = {};
    const srcName = baseName(imagePath);
    let fileFp = "";
    try { fileFp = fingerprintFile(imagePath); } catch (_) { throw new Error("无法读取所选图片文件，请确认文件仍然存在。"); }
    let replaced = 0;
    let failed = 0;
    const failures = [];
    const contentInfoByPath = {};
    let cancelled = false;
    let totalTargets = 0;
    let skippedRemainder = 0;
    try {
      const targets = resolveTargets(presentation, instances);
      totalTargets = targets.length;
      for (let i = 0; i < targets.length; i += 1) {
        if (cancelCheck && cancelCheck()) { cancelled = true; skippedRemainder = Math.max(0, totalTargets - i); break; }
        const target = targets[i];
        let oldName = "";
        try {
          const oldMeta = parseLink(String(target.shape.AlternativeText || ""));
          if (oldMeta && oldMeta.name) oldName = oldMeta.name;
        } catch (_) {}
        try {
          const newShape = replacePictureKeepCrop(target.shape, imagePath, scratch, cache);
          if (!newShape) {
            failed += 1;
            failures.push({ uid: target.inst && target.inst.uid ? target.inst.uid : "", slideIndex: target.inst ? target.inst.slideIndex : 0, name: oldName, reason: "插入新图片失败" });
            continue;
          }
          if (!contentInfoByPath[imagePath]) {
            const info = await contentFingerprintAndThumb(newShape, scratch);
            contentInfoByPath[imagePath] = { fp: info.fp, aspect: info.aspect };
          }
          attachLink(newShape, { name: oldName || srcName, src: imagePath, fileFp: fileFp, contentFp: contentInfoByPath[imagePath].fp, aspect: contentInfoByPath[imagePath].aspect });
          replaced += 1;
        } catch (error) {
          failed += 1;
          failures.push({ uid: target.inst && target.inst.uid ? target.inst.uid : "", slideIndex: target.inst ? target.inst.slideIndex : 0, name: oldName, reason: String(error && error.message || error) });
        }
        if (onProgress) { try { onProgress(i + 1, totalTargets, "替换图片 " + (i + 1) + "/" + totalTargets); } catch (_) {} }
        await yieldUI();
      }
    } finally { scratch.dispose(); }
    const skipped = (instances.length - totalTargets) + (cancelled ? skippedRemainder : 0);
    if (replaced > 0) invalidatePanelInventoryCache();
    return { replaced: replaced, failed: failed, skipped: skipped, failures: failures, cancelled: cancelled };
  }

  // For linked instances whose source file changed on disk, re-apply the
  // source image in place (keeping each instance's crop) and refresh the link.
  async function updateLinkedInstances(instances, docKey, onProgress, cancelCheck) {
    assertDocument(docKey);
    const presentation = activePresentation();
    const scratch = createScratchManager();
    const cache = {};
    let updated = 0;
    let failed = 0;
    let skipped = 0;
    let cancelled = false;
    const failures = [];
    const contentInfoBySrc = {};
    let totalTargets = 0;
    let skippedRemainder = 0;
    try {
      const targets = resolveTargets(presentation, instances);
      totalTargets = targets.length;
      for (let i = 0; i < targets.length; i += 1) {
        if (cancelCheck && cancelCheck()) { cancelled = true; skippedRemainder = Math.max(0, totalTargets - i); break; }
        const inst = targets[i].inst;
        const shape = targets[i].shape;
        if (!inst.linked || !inst.src) { skipped += 1; continue; }
        if (!fileExists(inst.src)) {
          failed += 1;
          failures.push({ uid: inst.uid || "", slideIndex: inst.slideIndex || 0, name: inst.name || "", reason: "源文件不存在" });
          continue;
        }
        try {
          const fileFp = fingerprintFile(inst.src);
          const newShape = replacePictureKeepCrop(shape, inst.src, scratch, cache);
          if (!newShape) {
            failed += 1;
            failures.push({ uid: inst.uid || "", slideIndex: inst.slideIndex || 0, name: inst.name || "", reason: "插入新图片失败" });
            continue;
          }
          if (!contentInfoBySrc[inst.src]) {
            const info = await contentFingerprintAndThumb(newShape, scratch);
            contentInfoBySrc[inst.src] = { fp: info.fp, aspect: info.aspect };
          }
          attachLink(newShape, { name: inst.name || baseName(inst.src), src: inst.src, fileFp: fileFp, contentFp: contentInfoBySrc[inst.src].fp, aspect: contentInfoBySrc[inst.src].aspect, userAlt: inst.userAlt });
          updated += 1;
        } catch (error) {
          failed += 1;
          failures.push({ uid: inst.uid || "", slideIndex: inst.slideIndex || 0, name: inst.name || "", reason: String(error && error.message || error) });
        }
        if (onProgress) { try { onProgress(i + 1, totalTargets, "更新链接 " + (i + 1) + "/" + totalTargets); } catch (_) {} }
        await yieldUI();
      }
    } finally { scratch.dispose(); }
    skipped = skipped + (cancelled ? skippedRemainder : 0);
    if (updated > 0) invalidatePanelInventoryCache();
    return { updated: updated, failed: failed, skipped: skipped, failures: failures, cancelled: cancelled };
  }

  // Remove link metadata but keep the user's own accessibility text.
  async function unlinkInstances(instances, docKey) {
    assertDocument(docKey);
    const presentation = activePresentation();
    let unlinked = 0;
    const targets = resolveTargets(presentation, instances);
    for (let i = 0; i < targets.length; i += 1) {
      const shape = targets[i].shape;
      try {
        const alt = String(shape.AlternativeText || "");
        const meta = parseLink(alt);
        if (!meta) continue;
        shape.AlternativeText = meta.userAlt || "";
        unlinked += 1;
      } catch (_) {}
    }
    if (unlinked > 0) invalidatePanelInventoryCache();
    return unlinked;
  }

  // Rename the shape (also syncs the link JSON name when present).
  function renameShape(shape, name) {
    const clean = String(name || "").trim();
    if (!clean) return false;
    try { shape.Name = clean; } catch (_) { return false; }
    try {
      const meta = parseLink(String(shape.AlternativeText || ""));
      if (meta) attachLink(shape, Object.assign({}, meta, { name: clean }));
    } catch (_) {}
    invalidatePanelInventoryCache();
    return true;
  }

  // WPS uses 2 for the slide-master view and 9 for the normal editing view
  // (the same value as PowerPoint's ppViewNormal).  GotoSlide is not enough
  // when the window is still in master view: WPS keeps the user on the
  // template canvas and silently ignores the requested slide.  Always leave
  // master view before navigating to a normal slide.
  const NORMAL_VIEW_TYPE = 9;

  function setWindowViewType(windowObject, viewType) {
    if (!windowObject) return false;
    try {
      windowObject.ViewType = viewType;
      // Do not busy-wait here. WPS applies view changes through its window
      // message loop; repeatedly reading ViewType in the same turn prevents
      // the requested transition from completing.
      return true;
    } catch (_) { return false; }
  }

  function readWindowViewType(windowObject) {
    try {
      const value = Number(windowObject.ViewType);
      return isFinite(value) ? value : null;
    } catch (_) { return null; }
  }

  function readWindowSlideIndex(windowObject) {
    try {
      const value = Number(windowObject.View.Slide.SlideIndex);
      if (value > 0) return value;
    } catch (_) {}
    // Test/fallback hosts may expose only a simple current index.
    try {
      const value = Number(windowObject.View.current);
      if (value > 0) return value;
    } catch (_) {}
    return null;
  }

  async function waitForWindowState(read, expected, timeoutMs) {
    const deadline = Date.now() + Math.max(200, Number(timeoutMs) || 1200);
    let value = null;
    while (Date.now() < deadline) {
      try { value = read(); } catch (_) { value = null; }
      if (value === expected) return true;
      await yieldUI(35);
    }
    return false;
  }

  async function requestWindowViewType(windowObject, viewType) {
    const before = readWindowViewType(windowObject);
    if (before === viewType) return true;
    if (!setWindowViewType(windowObject, viewType)) return false;
    // If the host does not expose a readable ViewType, wait two paint turns
    // and continue with selection verification as the final success signal.
    if (readWindowViewType(windowObject) === null) {
      await yieldUI(70);
      return true;
    }
    return waitForWindowState(function () { return readWindowViewType(windowObject); }, viewType, 1200);
  }

  function exitMasterView() {
    const windowObject = application().ActiveWindow;
    if (!windowObject) return false;
    let current = null;
    try { current = Number(windowObject.ViewType); } catch (_) {}
    // If the host does not expose ViewType, retain the old navigation path;
    // only a confirmed master-view state requires an explicit switch.
    if (current !== 2) return true;
    return setWindowViewType(windowObject, NORMAL_VIEW_TYPE);
  }

  function gotoSlide(slideIndex) {
    const windowObject = application().ActiveWindow;
    if (!windowObject || !windowObject.View || !exitMasterView()) return false;
    try { windowObject.View.GotoSlide(Number(slideIndex)); return true; } catch (_) { return false; }
  }

  function activateWindow(windowObject) {
    try { if (windowObject && hasMethod(windowObject, "Activate")) windowObject.Activate(); } catch (_) {}
  }

  // Shape.Select() is the documented way to highlight a picture on the
  // slide canvas.  Verify the host selection when WPS exposes ShapeRange so a
  // taskpane-only success cannot be mistaken for a canvas highlight.
  function selectionMatchesShape(windowObject, shape) {
    try {
      const range = windowObject && windowObject.Selection && windowObject.Selection.ShapeRange;
      if (!range) return null;
      const count = Number(range.Count) || 0;
      if (!count) return false;
      let targetId = "";
      let targetName = "";
      try { targetId = String(shape.Id || ""); } catch (_) {}
      try { targetName = String(shape.Name || ""); } catch (_) {}
      for (let i = 1; i <= count; i += 1) {
        let selected = null;
        try { selected = range.Item(i); } catch (_) { continue; }
        if (targetId) {
          try { if (String(selected.Id || "") === targetId) return true; } catch (_) {}
        }
        if (targetName) {
          try { if (String(selected.Name || "") === targetName) return true; } catch (_) {}
        }
      }
      return false;
    } catch (_) { return null; }
  }

  function selectCanvasShape(windowObject, shape) {
    if (!shape || !hasMethod(shape, "Select")) return false;
    if (layerGuardShapeLocked(shape)) return false;
    try {
      activateWindow(windowObject);
      shape.Select(MsoTrue);
      const verified = selectionMatchesShape(windowObject, shape);
      // Older WPS builds may not expose Selection.ShapeRange to JS. In that
      // case the successful Select() call remains the best available signal;
      // a host that exposes the range must match the requested shape.
      return verified === null ? true : verified;
    } catch (_) { return false; }
  }

  async function selectCanvasShapeAsync(windowObject, shape, shapeIndex) {
    if (!shape || !hasMethod(shape, "Select")) return false;
    if (layerGuardShapeLocked(shape)) return false;
    let requested = false;
    try {
      activateWindow(windowObject);
      shape.Select(MsoTrue);
      requested = true;
    } catch (_) {
      try { shape.Select(); requested = true; } catch (__) {}
    }
    if (!requested) return false;
    // Selection and its native resize handles are painted asynchronously.
    await yieldUI(55);
    let verified = selectionMatchesShape(windowObject, shape);
    if (verified === true) return true;

    // Some WPS builds are more reliable when selecting through Shapes.Range.
    // Retry only after the target slide/layout is active, and still require
    // Selection.ShapeRange to match before reporting a canvas highlight.
    try {
      const parent = shape.Parent;
      const shapes = parent && parent.Shapes ? parent.Shapes : parent;
      if (shapes && hasMethod(shapes, "Range")) {
        let name = "";
        try { name = String(shape.Name || ""); } catch (_) {}
        const candidates = [];
        if (name) { candidates.push(name); candidates.push([name]); }
        if (Number(shapeIndex) > 0) candidates.push(Number(shapeIndex));
        for (let i = 0; i < candidates.length; i += 1) {
          try {
            const range = shapes.Range(candidates[i]);
            if (!range || !hasMethod(range, "Select")) continue;
            range.Select(MsoTrue);
            await yieldUI(55);
            verified = selectionMatchesShape(windowObject, shape);
            if (verified === true) return true;
          } catch (_) {}
        }
      }
    } catch (_) {}
    // Returning false here is deliberate: a JSAPI call that did not throw is
    // not proof that the PPT正文 canvas actually displays the selection box.
    return false;
  }

  // Select a normal-slide picture after navigation so WPS shows its native
  // selection handles.  This is the document-side highlight users expect
  // from the panel's “定位” action; it also makes the target unambiguous when
  // several pictures share the same source image.
  function selectSlideShape(slideIndex, shapeIndex) {
    const presentation = activePresentation();
    try {
      const slide = presentation.Slides.Item(Number(slideIndex));
      if (!slide || !slide.Shapes) return false;
      const shape = slide.Shapes.Item(Number(shapeIndex));
      if (!isPicture(shape)) return false;
      if (layerGuardShapeLocked(shape)) return false;
      return selectCanvasShape(application().ActiveWindow, shape);
    } catch (_) { return false; }
  }

  async function locateSlideShape(slideIndex, shapeIndex, shapeReference) {
    const windowObject = application().ActiveWindow;
    if (!windowObject || !windowObject.View) return false;
    const targetIndex = Number(slideIndex);
    if (!(targetIndex > 0)) return false;
    let shape = shapeReference || null;
    let slide = null;
    try {
      if (shape) slide = slideOf(shape);
    } catch (_) { slide = null; }
    if (!slide || !shape) {
      try {
        slide = activePresentation().Slides.Item(targetIndex);
        shape = slide.Shapes.Item(Number(shapeIndex));
      } catch (_) { return false; }
    }
    if (!isPicture(shape)) return false;
    if (layerGuardShapeLocked(shape)) return false;

    activateWindow(windowObject);
    if (!await requestWindowViewType(windowObject, NORMAL_VIEW_TYPE)) return false;
    try { if (hasMethod(slide, "Select")) slide.Select(); } catch (_) {}
    try { windowObject.View.GotoSlide(targetIndex); } catch (_) { return false; }

    const readableIndex = readWindowSlideIndex(windowObject);
    if (readableIndex === null) {
      await yieldUI(90);
    } else if (!await waitForWindowState(function () { return readWindowSlideIndex(windowObject); }, targetIndex, 1500)) {
      return false;
    }
    return selectCanvasShapeAsync(windowObject, shape, shapeIndex);
  }

  function gotoMasterView() {
    const windowObject = application().ActiveWindow;
    if (!windowObject) return false;
    // Verified against real WPS (2026-08): ViewType = 2 (ppViewSlideMaster)
    // switches to the slide-master view and reads back the new value;
    // Office's 11 (ppViewThumbnails) is a silent no-op in WPS. WPS does not
    // define global PpViewType constants in the add-in JS context, so use the
    // numeric value directly.
    return setWindowViewType(windowObject, 2);
  }

  function selectMasterShape(shapeIndex) {
    const presentation = activePresentation();
    try {
      const master = presentation.SlideMaster;
      if (!master || !master.Shapes) return false;
      const shape = master.Shapes.Item(Number(shapeIndex));
      if (layerGuardShapeLocked(shape)) return false;
      return selectCanvasShape(application().ActiveWindow, shape);
    } catch (_) { return false; }
  }

  async function locateMasterShape(shapeIndex, shapeReference) {
    const windowObject = application().ActiveWindow;
    if (!windowObject || !await requestWindowViewType(windowObject, 2)) return false;
    let shape = shapeReference || null;
    if (!shape) {
      try { shape = activePresentation().SlideMaster.Shapes.Item(Number(shapeIndex)); } catch (_) { return false; }
    }
    if (layerGuardShapeLocked(shape)) return false;
    await yieldUI(55);
    return selectCanvasShapeAsync(windowObject, shape, shapeIndex);
  }

  function selectLayoutShape(layoutIndex, shapeIndex) {
    const presentation = activePresentation();
    try {
      const layouts = presentation.SlideMaster && presentation.SlideMaster.CustomLayouts;
      if (!layouts) return false;
      const windowObject = application().ActiveWindow;
      const layout = layouts.Item(Number(layoutIndex));
      // Selecting the layout is only navigation. The success signal must come
      // from selecting and (when exposed) verifying the requested picture;
      // otherwise the pane could claim a canvas highlight while only the
      // layout container is selected.
      try { if (hasMethod(layout, "Select")) { activateWindow(windowObject); layout.Select(MsoTrue); } } catch (_) {}
      let shapeSelected = false;
      try {
        const shape = layout.Shapes.Item(Number(shapeIndex));
        if (layerGuardShapeLocked(shape)) return false;
        shapeSelected = selectCanvasShape(windowObject, shape);
      } catch (_) {}
      return shapeSelected;
    } catch (_) { return false; }
  }

  async function locateLayoutShape(layoutIndex, shapeIndex, shapeReference) {
    const windowObject = application().ActiveWindow;
    if (!windowObject || !await requestWindowViewType(windowObject, 2)) return false;
    let shape = shapeReference || null;
    let layout = null;
    try { if (shape) layout = shape.Parent; } catch (_) {}
    if (!layout || !shape) {
      try {
        layout = activePresentation().SlideMaster.CustomLayouts.Item(Number(layoutIndex));
        shape = layout.Shapes.Item(Number(shapeIndex));
      } catch (_) { return false; }
    }
    if (layerGuardShapeLocked(shape)) return false;
    try { if (hasMethod(layout, "Select")) layout.Select(MsoTrue); } catch (_) {}
    await yieldUI(70);
    return selectCanvasShapeAsync(windowObject, shape, shapeIndex);
  }

  function addinPageUrl(page, fragment) {
    // WPS hosts offline add-in pages under http://taskpane.html/ and the
    // official SDK derives the add-in root from document.location.
    // CurrentWPSAddIn.Path/Name differ across WPS builds, so file://
    // candidates are verified on disk before being used; when nothing can
    // be verified we fall back to a relative URL that WPS resolves against
    // the add-in root.
    const fileName = String(page || "taskpane.html");
    let locationCandidate = "";
    const fileCandidates = [];
    try {
      if (typeof document !== "undefined" && document.location) {
        const loc = decodeURI(String(document.location));
        const idx = loc.indexOf("/");
        if (idx >= 0) {
          const base = loc.substring(0, loc.lastIndexOf("/"));
          if (base && base.indexOf("://") > 0) {
            locationCandidate = base + "/" + fileName + fragment;
          }
        }
      }
    } catch (_) {}
    try {
      const current = application().CurrentWPSAddIn;
      if (current) {
        const path = String(current.Path || "").replace(/[\\/]+$/, "");
        const name = String(current.Name || "").replace(/^[\\/]+|[\\/]+$/g, "");
        if (path) {
          fileCandidates.push("file:///" + path.replace(/\\/g, "/") + "/" + fileName + fragment);
          if (name) {
            fileCandidates.push("file:///" + path.replace(/\\/g, "/") + "/" + name + "/" + fileName + fragment);
          }
        }
      }
    } catch (_) {}
    for (let i = 0; i < fileCandidates.length; i += 1) {
      const url = fileCandidates[i];
      let filePath = url.slice(8).replace(/^\/+/, "").replace(/\//g, "\\");
      const hash = filePath.indexOf("#");
      if (hash >= 0) filePath = filePath.substring(0, hash);
      const query = filePath.indexOf("?");
      if (query >= 0) filePath = filePath.substring(0, query);
      try { if (fileExists(filePath)) return url; } catch (_) {}
    }
    if (locationCandidate) return locationCandidate;
    return fileName + fragment;
  }

  function addinUrl(fragment) {
    return addinPageUrl("taskpane.html", fragment);
  }

    function chooseImageFile(title) {
    const app = application();
    if (!hasMethod(app, "FileDialog")) throw new Error("当前 WPS 版本没有提供系统文件选择器。");
    // msoFileDialogFilePicker is 3 in the WPS/Office JSAPI enum.
    const dialog = app.FileDialog(3);
    dialog.Title = title || "选择图片文件";
    dialog.AllowMultiSelect = false;
    try { if (dialog.Filters && hasMethod(dialog.Filters, "Clear")) dialog.Filters.Clear(); } catch (_) {}
    try {
      if (dialog.Filters && hasMethod(dialog.Filters, "Add")) {
        dialog.Filters.Add("图片文件", "*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.tif;*.tiff;*.webp");
      }
    } catch (_) {}
    if (Number(dialog.Show()) !== MsoTrue) return "";
    if (!dialog.SelectedItems || Number(dialog.SelectedItems.Count) < 1) return "";
    return String(dialog.SelectedItems.Item(1));
  }

  let smartZoomPaneRef = null;
  let designToolsPaneRef = null;

  function openPane(fragment, title) {
    const pane = application().CreateTaskPane(addinUrl(fragment), title);
    if (!pane) throw new Error("无法创建 WPS 任务窗格，请确认已启用 JS 加载项。");
    pane.Visible = true;
    return pane;
  }

  function runAsync(work) {
    Promise.resolve().then(work).catch(function (error) { tell(error && error.message ? error.message : error); });
  }

  // =====================================================================
  // GitHub update check + one-click update/restart (v1.2.17)
  // =====================================================================
  const ADDIN_VERSION = "2.1.7";
  const UPDATE_MANIFEST_URL = "https://raw.githubusercontent.com/Dongsidaye/ppt-picture-replace-tools/agent/wps-adaptation-1-1-1/wps_addin/package.json";
  const UPDATE_RELEASE_BASE = "https://github.com/Dongsidaye/ppt-picture-replace-tools/releases/download/";
  const UPDATE_RELEASE_PAGE = "https://github.com/Dongsidaye/ppt-picture-replace-tools/releases/latest";

  function compareVersions(a, b) {
    const pa = String(a || "").split(".").map(function (x) { return parseInt(x, 10) || 0; });
    const pb = String(b || "").split(".").map(function (x) { return parseInt(x, 10) || 0; });
    const n = Math.max(pa.length, pb.length);
    for (let i = 0; i < n; i += 1) {
      const x = pa[i] || 0;
      const y = pb[i] || 0;
      if (x > y) return 1;
      if (x < y) return -1;
    }
    return 0;
  }

  function releaseDownloadUrl(version) {
    return UPDATE_RELEASE_BASE + "v" + String(version || "") + "/PictureReplaceTools-WPS-" + String(version || "") + ".exe";
  }

  function httpGetText(url, timeoutMs) {
    return new Promise(function (resolve) {
      let xhr = null;
      try {
        if (global.WpsInvoke && hasMethod(global.WpsInvoke, "CreateXHR")) xhr = global.WpsInvoke.CreateXHR();
        else if (typeof XMLHttpRequest !== "undefined") xhr = new XMLHttpRequest();
      } catch (_) { xhr = null; }
      if (!xhr) { resolve({ ok: false, error: "当前 WPS 环境不支持网络请求（无 XHR）" }); return; }
      let done = false;
      const finish = function (result) { if (!done) { done = true; resolve(result); } };
      try {
        xhr.open("GET", url, true);
        try { if (xhr.timeout !== undefined) xhr.timeout = timeoutMs || 20000; } catch (_) {}
        xhr.onreadystatechange = function () {
          if (xhr.readyState === 4) {
            const status = Number(xhr.status);
            const text = String(xhr.responseText || "");
            let finalUrl = "";
            try { finalUrl = String(xhr.responseURL || ""); } catch (_) {}
            if (status >= 200 && status < 300) finish({ ok: true, status: status, text: text, finalUrl: finalUrl, error: "" });
            else finish({ ok: false, status: status, text: text, finalUrl: finalUrl, error: "HTTP " + status });
          }
        };
        xhr.onerror = function () { finish({ ok: false, error: "网络请求失败" }); };
        try { if (xhr.ontimeout !== undefined) xhr.ontimeout = function () { finish({ ok: false, error: "网络请求超时" }); }; } catch (_) {}
        xhr.send();
      } catch (err) { finish({ ok: false, error: String(err && err.message || err) }); }
    });
  }

  async function checkForUpdates() {
    const result = { ok: false, current: ADDIN_VERSION, latest: "", hasUpdate: false, error: "", manifestUrl: UPDATE_MANIFEST_URL, downloadUrl: "", releasePage: UPDATE_RELEASE_PAGE };
    try {
      // Primary source: raw package.json on the release branch.
      const res = await httpGetText(UPDATE_MANIFEST_URL, 20000);
      let latest = "";
      if (res.ok) {
        let manifest = null;
        try { manifest = JSON.parse(res.text); } catch (_) { manifest = null; }
        if (manifest && manifest.version) latest = String(manifest.version);
      }
      // Fallback: GitHub releases/latest redirect -> tag path vX.Y.Z (works
      // even when raw.githubusercontent.com is unreachable from CN networks).
      if (!latest) {
        const tagRes = await httpGetText(UPDATE_RELEASE_PAGE, 20000);
        const tagUrl = String(tagRes && tagRes.finalUrl || (tagRes && tagRes.text) || "");
        const m = tagUrl.match(/\/releases\/tag\/v?(\d+(?:\.\d+)*)/i) || tagUrl.match(/v?(\d+\.\d+\.\d+)/);
        if (m && m[1]) latest = m[1];
        else result.error = (tagRes && tagRes.error) || "无法获取版本信息";
      }
      if (!latest) {
        if (!result.error) result.error = "版本信息格式错误";
        return result;
      }
      result.latest = latest;
      result.downloadUrl = releaseDownloadUrl(latest);
      result.ok = true;
      result.hasUpdate = compareVersions(latest, ADDIN_VERSION) > 0;
      return result;
    } catch (err) {
      result.error = String(err && err.message || err);
      return result;
    }
  }

  function downloadInstaller(url, destPath) {
    return new Promise(function (resolve) {
      let xhr = null;
      try {
        if (global.WpsInvoke && hasMethod(global.WpsInvoke, "CreateXHR")) xhr = global.WpsInvoke.CreateXHR();
        else if (typeof XMLHttpRequest !== "undefined") xhr = new XMLHttpRequest();
      } catch (_) { xhr = null; }
      if (!xhr) { resolve({ ok: false, error: "当前 WPS 环境不支持下载（无 XHR）" }); return; }
      let done = false;
      const finish = function (result) { if (!done) { done = true; resolve(result); } };
      try {
        xhr.open("GET", url, true);
        try { xhr.responseType = "arraybuffer"; } catch (_) {}
        try { if (xhr.overrideMimeType) xhr.overrideMimeType("text/plain; charset=x-user-defined"); } catch (_) {}
        xhr.onreadystatechange = function () {
          if (xhr.readyState !== 4) return;
          const status = Number(xhr.status);
          if (status < 200 || status >= 300) { finish({ ok: false, status: status, error: "下载失败 HTTP " + status }); return; }
          try {
            const fs = fileSystem();
            let bytes = null;
            try { bytes = xhr.response; } catch (_) { bytes = null; }
            let bin = "";
            if (bytes && bytes.byteLength) {
              const chunk = 32768;
              for (let i = 0; i < bytes.byteLength; i += chunk) {
                const slice = bytes.slice(i, Math.min(i + chunk, bytes.byteLength));
                const u8 = new Uint8Array(slice);
                let part = "";
                for (let k = 0; k < u8.length; k += 1) part += String.fromCharCode(u8[k]);
                bin += part;
              }
            } else {
              bin = String(xhr.responseText || "");
            }
            if (!bin.length) { finish({ ok: false, error: "下载内容为空" }); return; }
            // Lightweight sanity check: the installer is a Windows PE (MZ)
            // SFX and should be well over 100 KB; guards against truncated
            // downloads or a hostile replacement served in place of the exe.
            const firstTwo = bin.charCodeAt(0) === 0x4d && bin.charCodeAt(1) === 0x5a;
            if (!firstTwo) { finish({ ok: false, error: "下载内容不是有效的 Windows 安装程序（缺少 PE 头）" }); return; }
            if (bin.length < 100 * 1024) { finish({ ok: false, error: "下载内容过小，可能不完整" }); return; }
            try {
              if (fs.writeAsBinaryString) fs.writeAsBinaryString(destPath, bin);
              else if (fs.WriteFile) fs.WriteFile(destPath, bin);
            } catch (err) { finish({ ok: false, error: "写入安装包失败：" + String(err && err.message || err) }); return; }
            let exists = false;
            try { exists = !!(fs.Exists && fs.Exists(destPath)); } catch (_) {}
            finish({ ok: exists, size: bin.length, path: destPath, error: exists ? "" : "安装包写入后无法确认" });
          } catch (err) { finish({ ok: false, error: String(err && err.message || err) }); }
        };
        xhr.onerror = function () { finish({ ok: false, error: "下载网络失败" }); };
        try { if (xhr.ontimeout !== undefined) xhr.ontimeout = function () { finish({ ok: false, error: "下载超时" }); }; } catch (_) {}
        try { if (xhr.timeout !== undefined) xhr.timeout = 60000; } catch (_) {}
        xhr.send();
      } catch (err) { finish({ ok: false, error: String(err && err.message || err) }); }
    });
  }

  function installTargetPath() {
    try {
      const app = application();
      const addin = app.CurrentWPSAddIn;
      if (addin && addin.Path) return String(addin.Path);
    } catch (_) {}
    try {
      const fs = fileSystem();
      let base = "";
      try { base = fs.tmpdir(); } catch (_) {}
      if (base) return base;
    } catch (_) {}
    return "";
  }

  function writeAutoRestartMarker() {
    const fs = fileSystem();
    const markerName = "picture_replace_auto_restart.flag";
    try {
      const jsRoot = String(application().CurrentWPSAddIn && application().CurrentWPSAddIn.Path || "");
      if (jsRoot) {
        const marker = String(jsRoot).replace(/[\\/]+$/, "") + "\\" + markerName;
        try { fs.writeAsBinaryString(marker, "1"); } catch (_) {}
      }
    } catch (_) {}
    try {
      let base = "";
      try { base = fs.tmpdir(); } catch (_) {}
      if (base) {
        if (!/[\\/]$/.test(base)) base += "\\";
        try { fs.writeAsBinaryString(base + markerName, "1"); } catch (_) {}
      }
    } catch (_) {}
  }

  function shellExecutePath(pathOrUrl, params) {
    try {
      const app = application();
      if (app.OAAssist && hasMethod(app.OAAssist, "ShellExecute")) {
        try {
          if (params) app.OAAssist.ShellExecute(pathOrUrl, params);
          else app.OAAssist.ShellExecute(pathOrUrl);
        } catch (err1) {
          try { app.OAAssist.ShellExecute(pathOrUrl); } catch (err2) { return { ok: false, error: String(err2 && err2.message || err2) }; }
        }
        return { ok: true };
      }
    } catch (err) { return { ok: false, error: String(err && err.message || err) }; }
    return { ok: false, error: "当前 WPS 版本不支持自动启动安装程序（ShellExecute）" };
  }

  // WPP's documented OAAssist.ShellExecute is the most reliable way to leave
  // the task pane and open a browser. Some WPS builds also expose
  // Document.FollowHyperlink, so retain it as a compatibility fallback.
  function openExternalUrl(url) {
    const target = String(url || "").trim();
    if (!/^https?:\/\//i.test(target)) return { ok: false, error: "仅允许打开 http(s) 链接。" };
    let app = null;
    try { app = application(); } catch (error) { return { ok: false, error: String(error && error.message || error) }; }
    const shell = shellExecutePath(target, "");
    if (shell && shell.ok) return { ok: true, method: "ShellExecute", url: target };
    const owners = [];
    try { if (app.ActiveDocument) owners.push(app.ActiveDocument); } catch (_) {}
    try { if (app.ActivePresentation) owners.push(app.ActivePresentation); } catch (_) {}
    owners.push(app);
    for (let i = 0; i < owners.length; i += 1) {
      const owner = owners[i];
      if (!hasMethod(owner, "FollowHyperlink")) continue;
      try {
        owner.FollowHyperlink(target, true, true, true);
        return { ok: true, method: "FollowHyperlink", url: target };
      } catch (_) {}
    }
    return { ok: false, error: shell && shell.error ? shell.error : "当前 WPS 无法打开外部链接。", url: target };
  }

  function OpenProjectHome() {
    // Ribbon callbacks must perform the host call before the callback returns;
    // deferring it to a Promise can lose the WPS callback context and silently
    // do nothing on some builds.
    let result = null;
    try { result = openExternalUrl(PROJECT_HOME_URL); } catch (error) {
      result = { ok: false, error: String(error && error.message || error) };
    }
    if (!result || !result.ok) tell("无法打开项目主页：" + ((result && result.error) || "未知错误") + "\n请手动访问：\n" + PROJECT_HOME_URL, "项目主页");
    return result;
  }

  async function updateAndRestart() {
    const info = await checkForUpdates();
    if (!info.ok) { tell("检查更新失败：" + (info.error || "未知错误") + "\n可手动打开下载页：\n" + UPDATE_RELEASE_PAGE, "检查更新"); return false; }
    if (!info.hasUpdate) { tell("当前已是最新版本 v" + info.current + "。", "检查更新"); return false; }
    tell("正在下载 v" + info.latest + " 安装包，请稍候……", "一键更新");
    let dest = installTargetPath();
    if (dest) {
      if (!/[\\/]$/.test(dest)) dest += "\\";
      dest += "PictureReplaceTools-WPS-" + info.latest + ".exe";
    } else {
      dest = "PictureReplaceTools-WPS-" + info.latest + ".exe";
    }
    const dl = await downloadInstaller(info.downloadUrl, dest);
    if (!dl.ok) {
      tell("自动下载失败：" + (dl.error || "未知错误") + "\n可手动打开下载页安装：\n" + info.releasePage, "一键更新");
      return false;
    }
    try { writeAutoRestartMarker(); } catch (_) {}
    const se = shellExecutePath(dl.path, "-y");
    if (se.ok) {
      tell("已启动安装程序 v" + info.latest + "。\n安装完成后 WPS 会自动重启，请稍候。", "一键更新");
      return true;
    }
    tell("已下载安装包：\n" + dl.path + "\n\n请双击运行完成更新。", "一键更新");
    return false;
  }

  function openUpdatePanel() {
    runAsync(function () { openPane("#update", "检查更新"); });
  }

  function CheckForUpdates() { openUpdatePanel(); }

  function openUpdatePage() {
    const res = shellExecutePath(UPDATE_RELEASE_PAGE, "");
    if (!res.ok) tell("无法打开下载页：" + (res.error || "未知错误") + "\n请手动访问：\n" + UPDATE_RELEASE_PAGE, "检查更新");
  }


  // =====================================================================
  // Flag-gated self test (automated E2E + diagnostics). Only runs when
  // %TEMP%\picture_replace_selftest.flag exists; the flag JSON is:
  //   { "reportPath": "...", "newImage": "...", "deckWaitMs": 30000,
  //     "openPanel": true }
  // =====================================================================
  function selfTestFlagPaths() {
    const paths = [];
    try {
      const fs = fileSystem();
      let base = "";
      try { base = fs.tmpdir(); } catch (_) {}
      if (base) {
        if (!/[\\/]$/.test(base)) base += "\\";
        paths.push(base + "picture_replace_selftest.flag");
      }
    } catch (_) {}
    try {
      const addin = application().CurrentWPSAddIn;
      if (addin && addin.Path) {
        const base = String(addin.Path).replace(/[\\/]+$/, "") + "\\";
        paths.push(base + "picture_replace_selftest.flag");
      }
    } catch (_) {}
    return paths;
  }

  function sleep(ms) {
    return new Promise(function (resolve) {
      if (global.setTimeout) global.setTimeout(resolve, ms);
      else {
        const end = Date.now() + ms;
        while (Date.now() < end) { /* busy fallback */ }
        resolve();
      }
    });
  }

  async function waitForDeck(timeoutMs) {
    const end = Date.now() + timeoutMs;
    while (Date.now() < end) {
      try {
        const p = application().ActivePresentation;
        if (p && Number(p.Slides.Count) > 0) return true;
      } catch (_) {}
      await sleep(400);
    }
    return false;
  }

  async function runSelfTest(spec) {
    const report = { started: new Date().toISOString(), steps: [], ok: false };
    const push = function (name, data) { report.steps.push({ name: name, data: data }); };
    const scratch = createScratchManager();
    try {
      const ready = await waitForDeck(Number(spec.deckWaitMs) || 30000);
      push("deck_ready", ready);
      if (!ready) throw new Error("no active presentation with slides");

      const collect = await collectDeckImages();
      await refreshLinkStates(collect.groups);
      push("collect", {
        total: collect.total,
        groups: collect.groups.map(function (g) {
          return { name: g.name, count: g.instances.length, linkState: g.linkState, src: g.src };
        })
      });
      if (!collect.groups.length) throw new Error("collect returned no groups");
      const firstGroup = collect.groups[0];
      const targets = firstGroup.instances.slice(0, Math.min(2, firstGroup.instances.length));
      const rep = await replaceInstances(targets, spec.newImage, collect.docKey);
      push("replace_instances", rep);
      if (rep.replaced !== targets.length) throw new Error("replaceInstances partial failure");

      const collect2 = await collectDeckImages();
      await refreshLinkStates(collect2.groups);
      push("collect_after_replace", {
        total: collect2.total,
        groups: collect2.groups.map(function (g) {
          return { name: g.name, count: g.instances.length, linkState: g.linkState };
        })
      });
      const replacedGroup = collect2.groups[0];
      if (replacedGroup.linkState !== "linked" && replacedGroup.linkState !== "modified") {
        throw new Error("link state after replace unexpected: " + replacedGroup.linkState);
      }

      let batch = null;
      try {
        const anyShape = collect2.groups[0].instances[0].shape;
        anyShape.Select();
        batch = await replaceAllMatching(anyShape, spec.newImage);
        push("batch_replace", batch);
        if (!batch.matched || batch.success < 1) throw new Error("batch replace no matches");
      } catch (err) {
        push("batch_error", String(err && err.message || err));
        throw err;
      }

      const collect3 = await collectDeckImages();
      await refreshLinkStates(collect3.groups);
      const inst3 = collect3.groups[0].instances[0];
      const renamed = inst3.shape ? renameShape(inst3.shape, "自检重命名图") : false;
      push("rename", renamed);
      const gone = gotoSlide(inst3.slideIndex);
      push("goto_slide", gone);
      const un = await unlinkInstances([inst3], collect3.docKey);
      push("unlink", un);

      report.ok = true;
    } catch (err) {
      report.error = String(err && err.message || err);
    } finally {
      try { scratch.dispose(); } catch (_) {}
      try {
        const p = application().ActivePresentation;
        if (p) { try { p.Saved = true; } catch (_) {} }
      } catch (_) {}
    }
    report.finished = new Date().toISOString();
    return report;
  }

  // =====================================================================
  // Flag-gated performance probe (diagnostic only; no flag file = no-op):
  //   %TEMP%\picture_replace_profile.flag -> { "reportPath": "..." }
  // Times the (batched) picture scan and a legacy single-shape export loop,
  // then writes a JSON report and removes the flag.
  // =====================================================================
  function profileFlagPaths() {
    const paths = [];
    try {
      const fs = fileSystem();
      let base = "";
      try { base = fs.tmpdir(); } catch (_) {}
      if (base) {
        if (!/[\\/]$/.test(base)) base += "\\";
        paths.push(base + "picture_replace_profile.flag");
      }
    } catch (_) {}
    return paths;
  }

  async function maybeRunProfile() {
    try {

    } catch (_) {}
    const flagPaths = profileFlagPaths();
    let flagPath = "";
    let raw = "";
    for (let i = 0; i < flagPaths.length; i += 1) {
      try {
        const fs = fileSystem();
        if (fs.Exists && fs.Exists(flagPaths[i])) {
          flagPath = flagPaths[i];
          raw = String(fs.readAsBinaryString(flagPath) || "");
          break;
        }
      } catch (_) {}
    }
    if (!flagPath) return;
    try {
      const spec = {};
      try { Object.assign(spec, JSON.parse(raw || "{}")); } catch (_) {}
      const reportPath = String(spec.reportPath || "");
      const report = { started: new Date().toISOString(), elapsedMs: 0, total: 0, groups: 0, legacyPerExportMs: 0, fallbackCount: 0 };
      const ready = await waitForDeck(Number(spec.deckWaitMs) || 30000);
      if (!ready) throw new Error("no active presentation");
      const t0 = Date.now();
      const collect = await collectDeckImages();
      report.elapsedMs = Date.now() - t0;
      report.total = collect.total;
      report.groups = collect.groups.length;
      report.fallbackCount = collect.fallbackCount || 0;
      // Legacy per-shape export baseline on the first picture.
      try {
        const firstShape = collect.groups[0] && collect.groups[0].instances[0] && collect.groups[0].instances[0].shape;
        if (firstShape) {
          const scratch2 = createScratchManager();
          const n = 3;
          const ts = Date.now();
          for (let k = 0; k < n; k += 1) {
            const path = tempPath("legacy");
            try { exportUncroppedPreview(firstShape, scratch2, path, false, false, THUMB_PX); } catch (_) {}
            removeFile(path);
          }
          scratch2.dispose();
          report.legacyPerExportMs = Math.round((Date.now() - ts) / n);
        }
      } catch (_) {}
      if (reportPath) {
        try {
          const fs = fileSystem();
          const payload = JSON.stringify(report);
          if (fs.WriteFile) fs.WriteFile(reportPath, payload);
          else if (fs.writeAsBinaryString) fs.writeAsBinaryString(reportPath, payload);
        } catch (_) {}
      }
    } catch (err) {
      try {
        const fs = fileSystem();
        const payload = JSON.stringify({ error: String(err && err.message || err) });
        const reportPath = (() => { try { return String(JSON.parse(raw || "{}").reportPath || ""); } catch (_) { return ""; } })();
        if (reportPath) {
          if (fs.WriteFile) fs.WriteFile(reportPath, payload);
          else if (fs.writeAsBinaryString) fs.writeAsBinaryString(reportPath, payload);
        }
      } catch (_) {}
    } finally {
      for (let i = 0; i < flagPaths.length; i += 1) {
        try { fileSystem().Remove(flagPaths[i]); } catch (_) {}
        try { fileSystem().unlinkSync(flagPaths[i]); } catch (_) {}
      }
    }
  }
  // =====================================================================
  // Flag-gated view probe (diagnostic only; no flag file = no-op):
  //   %TEMP%\picture_replace_viewprobe.flag -> { "reportPath": "..." }
  // Surveys which view-switching APIs this WPS build actually implements:
  // ViewType read/write, global PpViewType constants, ExecuteMso ids,
  // and master-shape selection.
  // =====================================================================
  function viewProbeFlagPaths() {
    const paths = [];
    try {
      const fs = fileSystem();
      let base = "";
      try { base = fs.tmpdir(); } catch (_) {}
      if (base) {
        if (!/[\\/]$/.test(base)) base += "\\";
        paths.push(base + "picture_replace_viewprobe.flag");
      }
    } catch (_) {}
    try {
      const addin = application().CurrentWPSAddIn;
      if (addin && addin.Path) {
        const base = String(addin.Path).replace(/[\\/]+$/, "") + "\\";
        paths.push(base + "picture_replace_viewprobe.flag");
      }
    } catch (_) {}
    return paths;
  }

  function viewProbeNum(object, name) {
    try {
      const value = object[name];
      return value === undefined || value === null ? null : Number(value);
    } catch (_) { return null; }
  }

  async function maybeRunViewProbe() {
    try {

    } catch (_) {}
    const fs = fileSystem();
    const flagPaths = viewProbeFlagPaths();
    let flagPath = "";
    let raw = "";
    for (let i = 0; i < flagPaths.length; i += 1) {
      try {
        if (fs.Exists && fs.Exists(flagPaths[i])) {
          flagPath = flagPaths[i];
          raw = String(fs.readAsBinaryString(flagPath) || "");
          break;
        }
      } catch (_) {}
    }
    if (!flagPath || !raw) return;
    let spec = null;
    try { spec = JSON.parse(raw); } catch (_) { spec = null; }
    if (spec && spec.tracePath) setTracePath(spec.tracePath);
    else if (spec && spec.reportPath) setTracePath(spec.reportPath + ".trace");
    trace("spec parsed; report=" + (spec && spec.reportPath ? spec.reportPath : ""));
    if (!spec || !spec.reportPath) { trace("bad spec");
      for (let i = 0; i < flagPaths.length; i += 1) { try { fs.Remove(flagPaths[i]); } catch (_) {} }
      return;
    }
    const report = { started: new Date().toISOString(), entries: [], ok: false };
    const push = function (name, data) { report.entries.push({ name: name, data: data }); };
    try {
      const p = activePresentation();
      const app = application();
      const w = app.ActiveWindow;
      push("constants", {
        ppViewSlideMaster: viewProbeNum(global, "ppViewSlideMaster"),
        ppViewSlide: viewProbeNum(global, "ppViewSlide"),
        ppViewNormal: viewProbeNum(global, "ppViewNormal"),
        ppViewThumbnails: viewProbeNum(global, "ppViewThumbnails"),
        ppViewSlideSorter: viewProbeNum(global, "ppViewSlideSorter")
      });
      push("window_present", !!w);
      if (w) {
        const windowKeys = [];
        try { for (const key in w) windowKeys.push(String(key)); } catch (_) {}
        push("window_keys", windowKeys);
        push("viewtype_initial", viewProbeNum(w, "ViewType"));
        const view = w.View;
        if (view) {
          const viewKeys = [];
          try { for (const key in view) viewKeys.push(String(key)); } catch (_) {}
          push("view_keys", viewKeys);
          push("view_type", viewProbeNum(view, "Type"));
        }
        const results = [];
        const candidates = [2, 11, 12, 1, 7, 9];
        for (let i = 0; i < candidates.length; i += 1) {
          const value = candidates[i];
          const before = viewProbeNum(w, "ViewType");
          let setError = null;
          let after = null;
          try {
            w.ViewType = value;
            await sleep(500);
            after = viewProbeNum(w, "ViewType");
          } catch (err) {
            setError = String(err && err.message || err);
          }
          results.push({ set: value, before: before, after: after, changed: before !== null && after === value });
        }
        push("viewtype_writes", results);
        if (view) push("view_type_after_writes", viewProbeNum(view, "Type"));
        try { w.ViewType = 9; } catch (_) {}
      }
      let hasCommandBars = false;
      try { hasCommandBars = !!app.CommandBars; } catch (_) {}
      push("commandbars_present", hasCommandBars);
      const cmdResults = [];
      if (hasCommandBars) {
        const commandIds = ["ViewSlideMasterView", "SlideMasterView", "MasterView", "ViewNormal", "ViewSlideMaster"];
        for (let i = 0; i < commandIds.length; i += 1) {
          const id = commandIds[i];
          let ok = false;
          let error = null;
          let after = null;
          try {
            app.CommandBars.ExecuteMso(id);
            await sleep(500);
            ok = true;
            after = w ? viewProbeNum(w, "ViewType") : null;
          } catch (err) {
            error = String(err && err.message || err);
          }
          cmdResults.push({ id: id, ok: ok, error: error, afterViewType: after });
        }
      }
      push("executemso_results", cmdResults);
      try { if (w) w.ViewType = 9; } catch (_) {}
      // Enter master view, verify immediate read-back, then probe selection.
      let masterViewEntered = false;
      try { w.ViewType = 2; masterViewEntered = viewProbeNum(w, "ViewType") === 2; } catch (_) {}
      push("master_view_immediate", masterViewEntered);
      const selectResult = { ok: false, masterView: masterViewEntered };
      try {
        const master = p.SlideMaster;
        if (master && master.Shapes && Number(master.Shapes.Count) > 0) {
          const shape = master.Shapes.Item(1);
          shape.Select();
          selectResult.ok = true;
          try { selectResult.name = String(shape.Name || ""); } catch (_) {}
          try {
            selectResult.selectionCount = Number(w.Selection.ShapeRange.Count);
            try { selectResult.selectionName = String(w.Selection.ShapeRange.Item(1).Name || ""); } catch (_) {}
          } catch (_) {}
        } else {
          selectResult.skip = "no master shapes";
        }
      } catch (err) {
        selectResult.error = String(err && err.message || err);
      }
      push("master_shape_select", selectResult);
      const layoutProbe = {};
      try {
        const layouts = p.SlideMaster.CustomLayouts;
        const lo = layouts.Item(1);
        try { lo.Select(); layoutProbe.layoutSelect = true; } catch (err) { layoutProbe.layoutSelectError = String(err && err.message || err); }
        try {
          const lshape = lo.Shapes.Item(1);
          lshape.Select();
          layoutProbe.layoutShapeSelect = true;
        } catch (err) { layoutProbe.layoutShapeSelectError = String(err && err.message || err); }
      } catch (err) {
        layoutProbe.error = String(err && err.message || err);
      }
      push("layout_select_probe", layoutProbe);
      try { if (w) { w.ViewType = 9; push("viewtype_final", viewProbeNum(w, "ViewType")); } } catch (_) {}
      report.ok = true;
    } catch (err) {
      report.error = String(err && err.message || err);
    }
    report.finished = new Date().toISOString();
    try {
      const payload = JSON.stringify(report, null, 2);
      if (fs.WriteFile) fs.WriteFile(spec.reportPath, payload);
      else if (fs.writeAsBinaryString) fs.writeAsBinaryString(spec.reportPath, payload);
    } catch (_) {}
    for (let i = 0; i < flagPaths.length; i += 1) {
      try { fs.Remove(flagPaths[i]); } catch (_) {}
      try { fs.unlinkSync(flagPaths[i]); } catch (_) {}
    }
  }
  // =====================================================================
  // Flag-gated update-capability probe (diagnostic only; no flag = no-op):
  //   %TEMP%\picture_replace_updateprobe.flag -> { "reportPath": "..." }
  // Verifies GitHub reachability (XHR), OAAssist.ShellExecute, and whether
  // the add-in can write into the WPS jsaddins directory.
  // =====================================================================
  function updateProbeFlagPaths() {
    const paths = [];
    try {
      const fs = fileSystem();
      let base = "";
      try { base = fs.tmpdir(); } catch (_) {}
      if (base) {
        if (!/[\\/]$/.test(base)) base += "\\";
        paths.push(base + "picture_replace_updateprobe.flag");
      }
    } catch (_) {}
    try {
      const addin = application().CurrentWPSAddIn;
      if (addin && addin.Path) {
        const base = String(addin.Path).replace(/[\\/]+$/, "") + "\\";
        paths.push(base + "picture_replace_updateprobe.flag");
      }
    } catch (_) {}
    return paths;
  }

  async function maybeRunUpdateProbe() {
    const fs = fileSystem();
    const flagPaths = updateProbeFlagPaths();
    let flagPath = "";
    let raw = "";
    for (let i = 0; i < flagPaths.length; i += 1) {
      try {
        if (fs.Exists && fs.Exists(flagPaths[i])) {
          flagPath = flagPaths[i];
          raw = String(fs.readAsBinaryString(flagPath) || "");
          break;
        }
      } catch (_) {}
    }
    if (!flagPath || !raw) return;
    let spec = null;
    try { spec = JSON.parse(raw); } catch (_) { spec = null; }
    if (!spec || !spec.reportPath) {
      for (let i = 0; i < flagPaths.length; i += 1) { try { fs.Remove(flagPaths[i]); } catch (_) {} }
      return;
    }
    const report = { started: new Date().toISOString(), entries: [], ok: false };
    const push = function (name, data) { report.entries.push({ name: name, data: data }); };
    try {
      const app = application();
      const probeRoot = "C:/Users/Administrator/Documents/powerpoint/pic_replace_addin/";
      push("oaassist", { present: !!app.OAAssist, shellExecute: app.OAAssist ? hasMethod(app.OAAssist, "ShellExecute") : false });
      push("wpsinvoke", { present: typeof global.WpsInvoke !== "undefined", createXhr: typeof global.WpsInvoke !== "undefined" && hasMethod(global.WpsInvoke, "CreateXHR") });
      const xhr = global.WpsInvoke ? global.WpsInvoke.CreateXHR() : null;
      if (xhr) {
        const res = await new Promise(function (resolve) {
          try {
            xhr.open("GET", "https://raw.githubusercontent.com/Dongsidaye/ppt-picture-replace-tools/agent/wps-adaptation-1-1-1/wps_addin/package.json", true);
            xhr.onreadystatechange = function () {
              if (xhr.readyState === 4) {
                const t = String(xhr.responseText || "");
                resolve({ status: xhr.status, len: t.length, head: t.slice(0, 100) });
              }
            };
            xhr.onerror = function () { resolve({ error: "xhr error" }); };
            xhr.send();
          } catch (err) { resolve({ error: String(err && err.message || err) }); }
        });
        push("github_xhr", res);
      }
      const cmdPath = probeRoot + "shellexec_probe.cmd";
      const outPath = probeRoot + "shellexec_probe.txt";
      try {
        const cmdBody = "@echo off\r\necho ok> \"" + probeRoot.replace(/\//g, "\\") + "shellexec_probe.txt\"\r\n";
        fs.writeAsBinaryString(cmdPath, cmdBody);
      } catch (_) {}
      try { removeFile(outPath); } catch (_) {}
      const se = { attempted: false };
      try {
        if (app.OAAssist && hasMethod(app.OAAssist, "ShellExecute")) {
          se.attempted = true;
          try { app.OAAssist.ShellExecute(cmdPath); se.called = true; } catch (err) { se.callErr = String(err && err.message || err); }
        }
      } catch (err) { se.err = String(err && err.message || err); }
      await sleep(4000);
      try { se.result = !!(fs.Exists && fs.Exists(outPath)); } catch (_) {}
      push("shell_execute", se);
      const jsProbe = "C:/Users/Administrator/AppData/Roaming/kingsoft/wps/jsaddins/picture_replace_write_probe.txt";
      try {
        fs.writeAsBinaryString(jsProbe, "ok");
        push("jsaddins_write", !!(fs.Exists && fs.Exists(jsProbe)));
        try { fs.Remove(jsProbe); } catch (_) {}
      } catch (err) { push("jsaddins_write", String(err && err.message || err)); }
      report.ok = true;
    } catch (err) {
      report.error = String(err && err.message || err);
    }
    report.finished = new Date().toISOString();
    try {
      const payload = JSON.stringify(report, null, 2);
      if (fs.WriteFile) fs.WriteFile(spec.reportPath, payload);
      else if (fs.writeAsBinaryString) fs.writeAsBinaryString(spec.reportPath, payload);
    } catch (_) {}
    for (let i = 0; i < flagPaths.length; i += 1) {
      try { fs.Remove(flagPaths[i]); } catch (_) {}
      try { fs.unlinkSync(flagPaths[i]); } catch (_) {}
    }
  }

  async function maybeRunSelfTest() {
    const fs = fileSystem();
    const flagPaths = selfTestFlagPaths();
    let flagPath = "";
    let raw = "";
    for (let i = 0; i < flagPaths.length; i += 1) {
      try {
        if (fs.Exists && fs.Exists(flagPaths[i])) {
          flagPath = flagPaths[i];
          raw = String(fs.readAsBinaryString(flagPath) || "");
          break;
        }
      } catch (_) {}
    }
    if (!flagPath || !raw) return;
    let spec = null;
    try { spec = JSON.parse(raw); } catch (_) { spec = null; }
    if (!spec || !spec.reportPath) {
      try { fs.Remove(flagPath); } catch (_) {}
      return;
    }
    const report = await runSelfTest(spec);
    try {
      fs.writeAsBinaryString(spec.reportPath, JSON.stringify(report, null, 2));
    } catch (_) {
      try { fs.WriteFile(spec.reportPath, JSON.stringify(report, null, 2)); } catch (__) {}
    }
    for (let i = 0; i < flagPaths.length; i += 1) {
      try { if (fs.Exists && fs.Exists(flagPaths[i])) fs.Remove(flagPaths[i]); } catch (_) {}
    }
    if (spec.openPanel) {
      try { openPane("#panel", "图片清单"); } catch (_) {}
    }
  }

  var PICTURE_PANEL_ICON = "icon.png";
  function OnGetPicturePanelImage() {
    return PICTURE_PANEL_ICON;
  }
  function OnGetPanelImage() { return "icon.png"; }
  function OnGetFilterImage() { return "icon_filter.png"; }
  var TOOLS_ICON_BY_ID = {
    DesignStyleBrushButton: "icon_design_style.png",
    DesignTextToolsButton: "icon_design_text.png",
    DesignLayoutToolsButton: "icon_design_layout.png",
    DesignCleanupToolsButton: "icon_design_cleanup.png",
    DesignExportToolsButton: "icon_design_export.png",
    DesignColorToolsButton: "icon_design_color.png",
    DesignPhotoshopToolsButton: "icon_design_photoshop.png"
  };
  function OnGetToolsImage(control) {
    var cid = typeof control === "string" ? control : "";
    if (!cid && control) {
      if (control.Id !== undefined) cid = control.Id;
      else if (control.id !== undefined) cid = control.id;
      else if (control.Tag) cid = control.Tag;
    }
    if (cid && TOOLS_ICON_BY_ID[cid]) return TOOLS_ICON_BY_ID[cid];
    return "icon_filter.png";
  }
  function OnGetFileImage() { return "icon_file.png"; }
  function OnGetFileAllImage() { return "icon_file_all.png"; }
  function OnGetClipboardImage() { return "icon_clipboard.png"; }
  function OnGetClipboardAllImage() { return "icon_clipboard_all.png"; }
  function OnGetInfoImage() { return "icon_info.png"; }
  function OnGetGithubImage() { return "icon_github.png"; }
  function OnGetUpdateImage() { return "icon_update.png"; }
  var RIBBON_ICON_BY_ID = {
    OpenPicturePanelButton: "icon.png",
    CtxOpenPanel: "icon.png",
    ObjectFilterMenu: "icon_filter.png",
    SmartZoomButton: "icon_smart_zoom.png",
    CtxSmartZoom: "icon_smart_zoom.png",
    ReplacePictureFile: "icon_file.png",
    ReplaceAllFile: "icon_file_all.png",
    CtxReplaceFile: "icon_file.png",
    CtxReplaceAllFile: "icon_file_all.png",
    ReplacePictureClipboard: "icon_clipboard.png",
    ReplaceAllClipboard: "icon_clipboard_all.png",
    CtxReplaceClipboard: "icon_clipboard.png",
    CtxReplaceAllClipboard: "icon_clipboard_all.png",
    PictureReplaceCompatibility: "icon_info.png",
    OpenProjectHome: "icon_github.png"
  };
  function OnGetRibbonImage(control) {
    var cid = typeof control === "string" ? control : "";
    if (!cid && control) {
      if (control.Id !== undefined) cid = control.Id;
      else if (control.id !== undefined) cid = control.id;
      else if (control.Tag) cid = control.Tag;
    }
    if (cid && RIBBON_ICON_BY_ID[cid]) return RIBBON_ICON_BY_ID[cid];
    return "icon.png";
  }

  // Flag-gated ping (diagnostic only): create %TEMP%\picture_replace_ping.flag
  // to make the add-in write %TEMP%\picture_replace_ping.txt on load.
  function maybeRunPing() {
    try {
      const fs = fileSystem();
      let base = "";
      try { base = fs.tmpdir(); } catch (_) {}
      if (!base) return;
      if (!/[\\/]$/.test(base)) base += "\\";
      const flag = base + "picture_replace_ping.flag";
      if (!fs.Exists || !fs.Exists(flag)) return;
      const payload = "addin_ping " + new Date().toISOString();
      try { fs.writeAsBinaryString(base + "picture_replace_ping.txt", payload); } catch (_) { try { fs.WriteFile(base + "picture_replace_ping.txt", payload); } catch (__) {} }
      try { fs.Remove(flag); } catch (_) {}
      try { fs.unlinkSync(flag); } catch (_) {}
    } catch (_) {}
  }

  function OnAddInLoad() {
    try { maybeRunPing(); } catch (_) {}
    try { layerEnsureSelectionGuard(); } catch (_) {}
    // Start a best-effort warm-up after the host has had time to finish
    // opening the document.  WindowActivate/PresentationOpen/NewPresentation
    // callbacks repeat this for documents opened later in the same WPS run.
    try { bindPanelInventoryPreloadEvents(); } catch (_) {}
    try { schedulePanelInventoryPreload(650); } catch (_) {}
    runAsync(async function () { await maybeRunProfile(); await maybeRunViewProbe(); await maybeRunUpdateProbe(); await maybeRunSelfTest(); });
  }
  function OpenPicturePanel() {
    panelInventoryPanelUsed = true;
    runAsync(function () { openPane("#panel", "图片清单"); });
  }
  function OpenSmartZoomPane() {
    runAsync(function () {
      try {
        // A task pane can keep delayed slider callbacks alive after it is
        // hidden.  Close the previous pane/session before opening a fresh
        // one so old callbacks cannot act on a new selection.
        try { if (smartZoomPaneRef) smartZoomPaneRef.Visible = false; } catch (_) {}
        smartZoomPaneRef = null;
        try { smartZoomEnd(); } catch (_) {}
        smartZoomPaneRef = openPane("#zoom", "智能缩放");
      } catch (error) {
        tell(error && error.message ? error.message : error, "智能缩放");
      }
    });
  }
  const DESIGN_TOOL_PANEL_BY_ID = {
    DesignStyleBrushButton: "style",
    DesignTextToolsButton: "text",
    DesignTextExtractButton: "text",
    DesignLayoutToolsButton: "layout",
    DesignCleanupToolsButton: "cleanup",
    DesignCleanupHiddenButton: "cleanup",
    DesignCleanupOutsideButton: "cleanup",
    DesignCleanupNotesButton: "cleanup",
    DesignExportToolsButton: "export",
    DesignColorToolsButton: "color",
    DesignColorReplaceButton: "color",
    DesignColorAdjustButton: "color",
    DesignPhotoshopToolsButton: "photoshop"
  };

  function ribbonControlId(control) {
    if (typeof control === "string") return control;
    if (!control) return "";
    if (control.Id !== undefined && control.Id !== null) return String(control.Id);
    if (control.id !== undefined && control.id !== null) return String(control.id);
    return "";
  }

  function OpenDesignToolsPane(control) {
    const id = ribbonControlId(control);
    const panel = DESIGN_TOOL_PANEL_BY_ID[id] || (id.indexOf("Style") >= 0 ? "style"
      : id.indexOf("Text") >= 0 ? "text"
      : id.indexOf("Layout") >= 0 || id.indexOf("Align") >= 0 || id.indexOf("Distribute") >= 0 || id.indexOf("Uniform") >= 0 ? "layout"
      : id.indexOf("Cleanup") >= 0 ? "cleanup"
      : id.indexOf("Export") >= 0 || id.indexOf("LayerStamp") >= 0 || id.indexOf("ExtractSlides") >= 0 ? "export"
      : id.indexOf("Color") >= 0 ? "color"
      : id.indexOf("Photoshop") >= 0 || id.indexOf("PhotoshopReload") >= 0 ? "photoshop" : "style");
    runAsync(function () {
      try {
        try { if (designToolsPaneRef) designToolsPaneRef.Visible = false; } catch (_) {}
        designToolsPaneRef = null;
        designToolsPaneRef = openPane("#tools/" + panel, "设计工具");
      } catch (error) {
        tell(error && error.message ? error.message : error, "设计工具");
      }
    });
  }

  function RunDesignToolsCommand(control) {
    const id = ribbonControlId(control);
    runAsync(async function () {
      try {
        if (id === "DesignTextSwapButton") {
          const result = designTextSwap();
          tell(result.message, "批量文字");
          return;
        }
        if (id === "DesignAlignPageCenterButton") {
          const result = designAlignRun("align-page-center", {});
          tell(result.message, "排版工具");
          return;
        }
        if (id === "DesignDistributeHButton") {
          const result = designAlignRun("distribute-h", {});
          tell(result.message, "排版工具");
          return;
        }
        if (id === "DesignDistributeVButton") {
          const result = designAlignRun("distribute-v", {});
          tell(result.message, "排版工具");
          return;
        }
        if (id === "DesignUniformSizeButton") {
          const result = designAlignRun("uniform-size", {});
          tell(result.message, "排版工具");
          return;
        }
        if (id === "DesignLayerStampButton") {
          const result = await designLayerStamp();
          tell(result.message, "导出页面");
          return;
        }
        if (id === "DesignExtractSlidesButton") {
          const result = await designExtractSlides("selected");
          tell(result.message, "导出页面");
          return;
        }
        if (id === "DesignPhotoshopReloadButton") {
          const result = designPhotoshopReload();
          tell(result.message, "PS助手");
          return;
        }
        OpenDesignToolsPane(id || control);
      } catch (error) {
        tell(error && error.message ? error.message : error, "设计工具");
      }
    });
  }
  function ShowCompatibilityStatus() { tell(capabilityText(), "东四大爷的工具箱兼容性"); }
  function OpenSingleFilePane() {
    runAsync(function () {
      const target = selectedPicture();
      const path = chooseImageFile("文件原位替换 - 选择新图片");
      if (!path) return;
      replacePictureKeepCrop(target, path);
      tell("文件原位替换完成。", "东四大爷的工具箱");
    });
  }
  function OpenBatchFilePane() {
    runAsync(async function () {
      selectedPicture();
      const path = chooseImageFile("批量用文件替换 - 选择新图片");
      if (!path) return;
      const result = await runBatchWithProgress("批量文件替换", function (onProgress, cancelled) {
        return replaceAllFromFile(path, onProgress, cancelled);
      });
      tell(formatBatchResult(result) + (result && result.cancelled ? "（已取消）" : ""), "批量文件替换");
    });
  }
  function ReplaceSelectedFromClipboard() { runAsync(async function () { await replaceSelectedFromClipboard(); tell("剪贴板原位替换完成。"); }); }
  function formatBatchResult(result) {
    return "批量替换完成：匹配 " + result.matched + " 张，成功 " + result.success + " 张，失败 " + result.failed + " 张。";
  }
  function ReplaceAllFromClipboard() {
    runAsync(async function () {
      selectedPicture();
      const result = await runBatchWithProgress("批量剪贴板替换", function (onProgress, cancelled) {
        return replaceAllFromClipboard(onProgress, cancelled);
      });
      tell(formatBatchResult(result) + (result && result.cancelled ? "（已取消）" : ""), "批量剪贴板替换");
    });
  }

  global.OnAddInLoad = OnAddInLoad;
  global.OpenPicturePanel = OpenPicturePanel;
  global.OpenSmartZoomPane = OpenSmartZoomPane;
  global.OpenDesignToolsPane = OpenDesignToolsPane;
  global.RunDesignToolsCommand = RunDesignToolsCommand;
  global.OnGetPicturePanelImage = OnGetPicturePanelImage;
  global.OnGetRibbonImage = OnGetRibbonImage;
  global.OnGetFilterImage = OnGetFilterImage;
  global.OnGetToolsImage = OnGetToolsImage;
  global.OnGetPanelImage = OnGetPanelImage;
  global.OnGetFileImage = OnGetFileImage;
  global.OnGetFileAllImage = OnGetFileAllImage;
  global.OnGetClipboardImage = OnGetClipboardImage;
  global.OnGetClipboardAllImage = OnGetClipboardAllImage;
  global.OnGetInfoImage = OnGetInfoImage;
  global.OnGetGithubImage = OnGetGithubImage;
  global.OpenAnimationPane = OpenAnimationPane;
  global.OpenSelectionPane = OpenSelectionPane;
  global.OpenProjectHome = OpenProjectHome;
  global.SelectAllObjects = SelectAllObjects;
  global.InvertSelection = InvertSelection;
  global.SelectSameType = SelectSameType;
  global.SelectSameFontSize = SelectSameFontSize;
  global.SelectSameWidth = SelectSameWidth;
  global.SelectSameHeight = SelectSameHeight;
  global.SelectSameColor = SelectSameColor;
  global.SelectAllLines = SelectAllLines;
  global.SelectAllText = SelectAllText;
  global.SelectAllGroups = SelectAllGroups;
  global.OpenSingleFilePane = OpenSingleFilePane;
  global.OpenBatchFilePane = OpenBatchFilePane;
  global.ReplaceSelectedFromClipboard = ReplaceSelectedFromClipboard;
  global.ReplaceAllFromClipboard = ReplaceAllFromClipboard;
  global.ShowCompatibilityStatus = ShowCompatibilityStatus;
  global.CheckForUpdates = CheckForUpdates;
  global.OnGetUpdateImage = OnGetUpdateImage;
  global.WpsPictureReplace = {
    writeBrowserFile: writeBrowserFile,
    replaceSelectedFromFile: replaceSelectedFromFile,
    replaceAllMatching: replaceAllMatching,
    replaceAllFromFile: replaceAllFromFile,
    replaceSelectedFromClipboard: replaceSelectedFromClipboard,
    replaceAllFromClipboard: replaceAllFromClipboard,
    replacePictureKeepCrop: replacePictureKeepCrop,
    pasteReplacePictureKeepCrop: pasteReplacePictureKeepCrop,
    capabilityProbe: capabilityProbe,
    capabilityText: capabilityText,
    formatBatchResult: formatBatchResult,
    chooseImageFile: chooseImageFile,
    compareVersions: compareVersions,
    addinVersion: ADDIN_VERSION,
    checkForUpdates: checkForUpdates,
    updateAndRestart: updateAndRestart,
    downloadInstaller: downloadInstaller,
    shellExecutePath: shellExecutePath,
    openExternalUrl: openExternalUrl,
    projectHomeUrl: PROJECT_HOME_URL,
    getUpdateInfo: checkForUpdates,
    openUpdatePanel: openUpdatePanel,
    openUpdatePage: openUpdatePage,
    addinUrl: addinUrl,
    collectDeckImages: collectDeckImages,
    collectDeckImagesCached: collectDeckImagesCached,
    clearDeckImageCache: invalidatePanelInventoryCache,
    preloadDeckImages: preloadDeckImages,
    refreshLinkStates: refreshLinkStates,
    replaceInstances: replaceInstances,
    updateLinkedInstances: updateLinkedInstances,
    requestCancelTask: requestCancelTask,
    taskCancelled: taskCancelled,
    readTaskState: readTaskState,
    openProgressPanel: openProgressPanel,
    writeTaskProgress: writeTaskProgress,
    closeProgressPanel: closeProgressPanel,
    objectFilterRun: objectFilterRun,
    objectFilterSelectedShapes: objectFilterSelectedShapes,
    layerList: layerList,
    layerSelect: layerSelect,
    layerSelectMany: layerSelectMany,
    layerSelectAll: layerSelectAll,
    layerSetVisible: layerSetVisible,
    layerSetVisibleMany: layerSetVisibleMany,
    layerSetLocked: layerSetLocked,
    layerSetLockedMany: layerSetLockedMany,
    designStyleCapture: designStyleCapture,
    designStyleInfo: designStyleInfo,
    designStyleApply: designStyleApply,
    designTextFindReplace: designTextFindReplace,
    designTextSwap: designTextSwap,
    designTextExtract: designTextExtract,
    designAlignRun: designAlignRun,
    designCleanup: designCleanup,
    designExportSlides: designExportSlides,
    designLayerStamp: designLayerStamp,
    designExtractSlides: designExtractSlides,
    designColorAdjust: designColorAdjust,
    designColorReplace: designColorReplace,
    designPhotoshopOpen: designPhotoshopOpen,
    designPhotoshopReload: designPhotoshopReload,
    unlinkInstances: unlinkInstances,
    renameShape: renameShape,
    gotoSlide: gotoSlide,
    exitMasterView: exitMasterView,
    selectSlideShape: selectSlideShape,
    locateSlideShape: locateSlideShape,
    gotoMasterView: gotoMasterView,
    selectMasterShape: selectMasterShape,
    selectLayoutShape: selectLayoutShape,
    locateMasterShape: locateMasterShape,
    locateLayoutShape: locateLayoutShape,
    smartZoomBegin: smartZoomBegin,
    smartZoomApply: smartZoomApply,
    smartZoomReset: smartZoomReset,
    smartZoomInfo: smartZoomInfo,
    smartZoomPercentForWidth: smartZoomPercentForWidth,
    smartZoomEnd: smartZoomEnd,
    parseLink: parseLink,
    formatLink: formatLink,
    baseName: baseName,
    _math: {
      recoverNaturalSize: recoverNaturalSize,
      computeNewCrops: computeNewCrops,
      imagePixelSize: imagePixelSize,
      linkedNaturalSize: linkedNaturalSize,
      PREVIEW_PX: PREVIEW_PX
    }
  };
}(typeof window !== "undefined" ? window : globalThis));
