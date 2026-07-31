/* Picture Replace Tools - WPS WPP JavaScript add-in.
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
      if (global.console) global.console.error(title || "图片原位替换", message);
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
    try { void shape.PictureFormat; } catch (_) { throw new Error("当前选中对象不是图片。"); }
    return shape;
  }

  function slideOf(shape) {
    let slide = shape && shape.Parent;
    if (!slide || !slide.Shapes) {
      try { slide = shape.Parent.Parent; } catch (_) {}
    }
    if (!slide || !slide.Shapes) throw new Error("无法取得图片所在幻灯片。");
    return slide;
  }

  function isPicture(shape) {
    try { void shape.PictureFormat; return true; } catch (_) { return false; }
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

  function captureState(shape, scratch) {
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
        for (let attempt = 0; attempt < 3 && after <= before; attempt += 1) {
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
      } catch (_) {}
      try { if (dup) dup.Delete(); } catch (_) {}
    }

    const recovered = recoverNaturalSize(frameW, frameH, cropLeft, cropRight, cropTop, cropBottom, fullW, fullH);
    const naturalW = recovered.naturalW;
    const naturalH = recovered.naturalH;
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

    return {
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
    try { target.PictureFormat.Brightness = source.PictureFormat.Brightness; target.PictureFormat.Contrast = source.PictureFormat.Contrast; target.PictureFormat.ColorType = source.PictureFormat.ColorType; } catch (_) {}
    try { target.AlternativeText = source.AlternativeText; } catch (_) {}
  }

  function replacePictureKeepCrop(oldShape, imagePath, scratch) {
    const ownsScratch = !scratch;
    const sc = scratch || createScratchManager();
    try {
      const state = captureState(oldShape, sc);
      const slide = slideOf(oldShape);
      let inserted = null;
      try {
        // Width/Height are intentionally omitted: WPS inserts at natural size.
        inserted = asShape(slide.Shapes.AddPicture(imagePath, MsoFalse, MsoTrue, state.left, state.top));
        if (!inserted) throw new Error("插入新图片失败。");
        const naturalW = num(inserted.Width);
        const naturalH = num(inserted.Height);
        applyPreservedCrop(inserted, state, naturalW, naturalH);
        if (isTrue(inserted.HorizontalFlip) !== state.flipH) inserted.Flip(MsoFlipHorizontal);
        if (isTrue(inserted.VerticalFlip) !== state.flipV) inserted.Flip(MsoFlipVertical);
        try { inserted.Rotation = state.rotation; } catch (_) {}
        inserted.Left = state.left;
        inserted.Top = state.top;
        try { inserted.LockAspectRatio = state.lockAspectRatio; } catch (_) {}
        copyCosmetics(oldShape, inserted);
        let guard = 0;
        while (num(inserted.ZOrderPosition) > state.zOrder + 1 && guard < 1000) {
          try { inserted.ZOrder(MsoSendBackward); } catch (_) {}
          guard += 1;
        }
        oldShape.Delete();
        try { inserted.Name = state.name; } catch (_) {}
        try { inserted.AlternativeText = state.alternativeText; } catch (_) {}
        try { inserted.Select(); } catch (_) {}
        return true;
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

  // Render the FULL (uncropped, unrotated, flip-normalized) image of a shape
  // to a PNG through a scratch slide. Used to fingerprint the original image
  // so different crop instances of the same picture match each other.
  function exportUncroppedPreview(shape, scratch, path, extraFlipH, extraFlipV) {
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
      slide.Export(path, "PNG", PREVIEW_PX, PREVIEW_PX);
      return fileExists(path);
    } catch (error) {
      removeFile(path);
      return false;
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
      if (!exportUncroppedPreview(shape, scratch, path, variants[i][0], variants[i][1])) { removeFile(path); continue; }
      signatures.push(await signatureFromPath(path));
      removeFile(path);
    }
    return signatures;
  }

  async function replaceAllMatching(referenceShape, imagePath) {
    const presentation = activePresentation();
    const scratch = createScratchManager();
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
        const path = tempPath("candidate");
        try {
          if (!exportUncroppedPreview(candidates[i], scratch, path, false, false)) continue;
          const signature = await signatureFromPath(path);
          if (referenceSignatures.has(signature)) matches.push(candidates[i]);
        } finally { removeFile(path); }
      }

      let success = 0;
      for (let i = 0; i < matches.length; i += 1) {
        try { if (replacePictureKeepCrop(matches[i], imagePath, scratch)) success += 1; } catch (_) {}
      }
      return { matched: matches.length, success: success, failed: matches.length - success };
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

  async function replaceAllFromFile(path) {
    try {
      const result = await replaceAllMatching(selectedPicture(), path);
      if (!result.matched) throw new Error("没有找到匹配的原图实例。");
      return result;
    } finally { removeFile(path); }
  }

  // Direct clipboard replacement: paste the clipboard image onto the target
  // slide itself (no intermediate PNG), then apply the preserved crop using
  // the exact same math as file replacement. The paste happens BEFORE
  // captureState because captureState copies the old shape and would
  // otherwise overwrite the user's clipboard.
  function pasteReplacePictureKeepCrop(target) {
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
      const state = captureState(target, owns);
      const naturalW = num(inserted.Width);
      const naturalH = num(inserted.Height);
      applyPreservedCrop(inserted, state, naturalW, naturalH);
      if (isTrue(inserted.HorizontalFlip) !== state.flipH) inserted.Flip(MsoFlipHorizontal);
      if (isTrue(inserted.VerticalFlip) !== state.flipV) inserted.Flip(MsoFlipVertical);
      try { inserted.Rotation = state.rotation; } catch (_) {}
      inserted.Left = state.left;
      inserted.Top = state.top;
      try { inserted.LockAspectRatio = state.lockAspectRatio; } catch (_) {}
      copyCosmetics(target, inserted);
      let guard = 0;
      while (num(inserted.ZOrderPosition) > state.zOrder + 1 && guard < 1000) {
        try { inserted.ZOrder(MsoSendBackward); } catch (_) {}
        guard += 1;
      }
      target.Delete();
      try { inserted.Name = state.name; } catch (_) {}
      try { inserted.AlternativeText = state.alternativeText; } catch (_) {}
      try { inserted.Select(); } catch (_) {}
      return true;
    } finally {
      owns.dispose();
    }
  }

  async function replaceSelectedFromClipboard() {
    return pasteReplacePictureKeepCrop(selectedPicture());
  }
  async function replaceAllFromClipboard() {
    const reference = selectedPicture();
    const path = tempPath("clipboard_batch");
    const scratch = createScratchManager();
    try {
      pasteClipboardAsPng(scratch, path);
      const result = await replaceAllMatching(reference, path);
      if (!result.matched) throw new Error("没有找到匹配的原图实例。");
      return result;
    } finally { scratch.dispose(); removeFile(path); }
  }

  function addinUrl(fragment) {
    try {
      const current = application().CurrentWPSAddIn;
      if (current && current.Path) {
        const base = String(current.Path).replace(/[\\/]$/, "");
        const name = String(current.Name || "").replace(/^[\\/]+|[\\/]+$/g, "");
        // WPS versions differ: some return the plugin folder from Path,
        // others return its parent jsaddons folder and put the folder name in
        // Name. Prefer the latter when it is available, then fall back.
        if (name && !base.toLowerCase().endsWith("/" + name.toLowerCase()) &&
            !base.toLowerCase().endsWith("\\" + name.toLowerCase())) {
          return base + "/" + name + "/taskpane.html" + fragment;
        }
        return base + "/taskpane.html" + fragment;
      }
    } catch (_) {}
    return "taskpane.html" + fragment;
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

  function openPane(fragment, title) {
    const pane = application().CreateTaskPane(addinUrl(fragment), title);
    if (!pane) throw new Error("无法创建 WPS 任务窗格，请确认已启用 JS 加载项。");
    pane.Visible = true;
  }

  function runAsync(work) {
    Promise.resolve().then(work).catch(function (error) { tell(error && error.message ? error.message : error); });
  }

  function OnAddInLoad() {}
  function ShowCompatibilityStatus() { tell(capabilityText(), "WPS 图片原位替换兼容性"); }
  function OpenSingleFilePane() {
    runAsync(function () {
      const path = chooseImageFile("文件原位替换 - 选择新图片");
      if (!path) return;
      replacePictureKeepCrop(selectedPicture(), path);
      tell("文件原位替换完成。", "图片原位替换");
    });
  }
  function OpenBatchFilePane() {
    runAsync(async function () {
      const path = chooseImageFile("批量用文件替换 - 选择新图片");
      if (!path) return;
      tell(formatBatchResult(await replaceAllFromFile(path)), "图片原位替换");
    });
  }
  function ReplaceSelectedFromClipboard() { runAsync(async function () { await replaceSelectedFromClipboard(); tell("剪贴板原位替换完成。"); }); }
  function formatBatchResult(result) {
    return "批量替换完成：匹配 " + result.matched + " 张，成功 " + result.success + " 张，失败 " + result.failed + " 张。";
  }
  function ReplaceAllFromClipboard() { runAsync(async function () { tell(formatBatchResult(await replaceAllFromClipboard())); }); }

  global.OnAddInLoad = OnAddInLoad;
  global.OpenSingleFilePane = OpenSingleFilePane;
  global.OpenBatchFilePane = OpenBatchFilePane;
  global.ReplaceSelectedFromClipboard = ReplaceSelectedFromClipboard;
  global.ReplaceAllFromClipboard = ReplaceAllFromClipboard;
  global.ShowCompatibilityStatus = ShowCompatibilityStatus;
  global.WpsPictureReplace = {
    writeBrowserFile: writeBrowserFile,
    replaceSelectedFromFile: replaceSelectedFromFile,
    replaceAllFromFile: replaceAllFromFile,
    replaceSelectedFromClipboard: replaceSelectedFromClipboard,
    replaceAllFromClipboard: replaceAllFromClipboard,
    replacePictureKeepCrop: replacePictureKeepCrop,
    pasteReplacePictureKeepCrop: pasteReplacePictureKeepCrop,
    capabilityProbe: capabilityProbe,
    capabilityText: capabilityText,
    formatBatchResult: formatBatchResult,
    chooseImageFile: chooseImageFile,
    _math: {
      recoverNaturalSize: recoverNaturalSize,
      computeNewCrops: computeNewCrops,
      PREVIEW_PX: PREVIEW_PX
    }
  };
}(typeof window !== "undefined" ? window : globalThis));
