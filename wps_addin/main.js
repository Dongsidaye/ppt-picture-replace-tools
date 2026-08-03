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

  const THUMB_PX = 160;
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
    if (!isPicture(shape)) throw new Error("当前选中对象不是图片。");
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
  function yieldUI() {
    return new Promise(function (resolve) { setTimeout(resolve, 0); });
  }

  // Modal-style progress dialog (WPS Application.ShowDialog, official API).
  // The batch operation runs INSIDE the dialog page context, so the progress
  // bar is always visible and cancel works even if ShowDialog blocks the
  // caller. Returns false when the API is unavailable (caller falls back).
  function openProgressDialog(command) {
    try {
      const app = application();
      if (!hasMethod(app, "ShowDialog")) return false;
      const dpr = (typeof global.devicePixelRatio === "number" && global.devicePixelRatio > 0) ? global.devicePixelRatio : 1;
      app.ShowDialog(addinPageUrl("dialog_progress.html", "#" + encodeURIComponent(command)), "图片替换进度", Math.round(440 * dpr), Math.round(240 * dpr), false);
      return true;
    } catch (_) { return false; }
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

  // Decode PNG bytes to a 64-bit dHash via Image + canvas (CEF hosts).
  function dHashFromBinary(binary) {
    return new Promise(function (resolve) {
      if (!(global.Image && global.document && global.URL && global.Blob)) { resolve(""); return; }
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
            resolve(dHashFromPixels(data, 9));
          } catch (_) { global.URL.revokeObjectURL(url); resolve(""); }
        };
        img.onerror = function () { global.URL.revokeObjectURL(url); resolve(""); };
        img.src = url;
      } catch (_) { resolve(""); }
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

  // Export the full uncropped image once: return both a content fingerprint
  // (for same-source grouping) and a base64 thumbnail for the panel.
  async function contentFingerprintAndThumb(shape, scratch) {
    const path = tempPath("panel");
    try {
      const exported = exportUncroppedPreview(shape, scratch, path, false, false, THUMB_PX);
      if (!exported.ok) {
        return { fp: "", dataUrl: "", dHash: "", w: 0, h: 0, nw: 0, nh: 0, aspect: 0 };
      }
      const binary = fileSystem().readAsBinaryString(path);
      const fp = fnv1a(binary);
      const dataUrl = await fileToDataUrl(binary);
      const dHash = await dHashFromBinary(binary);
      // nw/nh = full uncropped image size before normalization.
      const nw = num(exported.fullW);
      const nh = num(exported.fullH);
      const aspect = nw > 0 && nh > 0 ? Math.round(1000 * nw / nh) : 0;
      return { fp: fp, dataUrl: dataUrl, dHash: dHash, w: THUMB_PX, h: THUMB_PX, nw: nw, nh: nh, aspect: aspect };
    } finally { removeFile(path); }
  }

  // =====================================================================
  // Batch thumbnail/fingerprint rendering. WPS JSAPI has no Shape.Export,
  // so the old path exported one scratch slide PER picture (each export and
  // each clipboard round-trip crosses the JSAPI bridge). Instead we paste up
  // to 64 pictures into an 8x8 grid on ONE scratch slide and export the grid
  // once; cells are then decoded in JS (dHash + fingerprint + thumbnail).
  // =====================================================================
  const BATCH_COLS = 8;
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

  function decodeBatchCells(binary, count) {
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
                const cell = global.document.createElement("canvas");
                cell.width = BATCH_CELL;
                cell.height = BATCH_CELL;
                const cctx = cell.getContext("2d");
                cctx.drawImage(full, col * cellPx, row * cellPx, cellPx, cellPx, 0, 0, BATCH_CELL, BATCH_CELL);
                const tiny = global.document.createElement("canvas");
                tiny.width = 9;
                tiny.height = 8;
                const tctx = tiny.getContext("2d");
                tctx.drawImage(cell, 0, 0, 9, 8);
                const dHash = dHashFromPixels(tctx.getImageData(0, 0, 9, 8).data, 9);
                const pixels = cctx.getImageData(0, 0, BATCH_CELL, BATCH_CELL).data;
                // Fingerprint over raw RGBA: encoder-independent and stable
                // across WPS versions/re-encodes (old bytes-based fingerprint
                // is kept only as a fallback for linked-cache reuse).
                const fp = fnvBytes(pixels);
                const dataUrl = cell.toDataURL("image/png");
                out[i] = { fp: fp, dHash: dHash, dataUrl: dataUrl };
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

  async function renderThumbnailsBatch(items, scratch) {
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
          naturals[i] = { nw: nw, nh: nh };
          pastedShape.Left = (idx % BATCH_COLS) * BATCH_CELL;
          pastedShape.Top = Math.floor(idx / BATCH_COLS) * BATCH_CELL;
          pastedShape.Width = BATCH_CELL;
          pastedShape.Height = BATCH_CELL;
          pasted += 1;
        } catch (_) {}
      }
      if (!pasted) continue;
      const path = tempPath("batch");
      removeFile(path);
      try {
        slide.Export(path, "PNG", BATCH_GRID, BATCH_GRID);
        if (!fileExists(path)) continue;
        const binary = fileSystem().readAsBinaryString(path);
        const decoded = await decodeBatchCells(binary, end - start);
        if (!decoded) continue;
        for (let i = start; i < end; i += 1) {
          const info = decoded[i - start];
          if (!info) continue;
          const nat = naturals[i];
          const nw = nat ? nat.nw : 0;
          const nh = nat ? nat.nh : 0;
          results[i] = {
            fp: info.fp,
            dataUrl: info.dataUrl,
            dHash: info.dHash,
            w: BATCH_CELL,
            h: BATCH_CELL,
            nw: nw,
            nh: nh,
            aspect: nw > 0 && nh > 0 ? Math.round(1000 * nw / nh) : 0
          };
        }
      } finally { removeFile(path); }
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
    try { current = fingerprintFile(src); } catch (_) { return "unreadable"; }
    return current === linked[0].fileFp ? "linked" : "modified";
  }

  // Recompute link status (source file hash) for all groups. Kept out of the
  // initial collect so the panel paints instantly; badges update afterwards.
  async function refreshLinkStates(groups, onProgress) {
    const states = [];
    for (let g = 0; g < groups.length; g += 1) {
      const state = await linkStateOf(groups[g]);
      try { groups[g].linkState = state; } catch (_) {}
      states.push(state);
      if (onProgress) { try { onProgress(g + 1, groups.length); } catch (_) {} }
    }
    return states;
  }
  // Collect every picture in the active presentation, group by content
  // fingerprint (same source image), and report link status per group.
  // onProgress(done, total) is invoked between instances.
  function documentKey(presentation) {
    try {
      const full = String(presentation.FullName || "");
      if (full) return "file:" + normalizePath(full);
    } catch (_) {}
    let key = "unsaved:";
    try { key += Number(presentation.Slides.Count) || 0; } catch (_) { key += "?"; }
    try {
      const sc = Number(presentation.Slides.Count) || 0;
      for (let i = 1; i <= sc; i += 1) {
        key += ":" + (Number(presentation.Slides.Item(i).Shapes.Count) || 0);
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
    for (let g = 0; g < groups.length; g += 1) {
      const state = await linkStateOf(groups[g]);
      try { groups[g].linkState = state; } catch (_) {}
      states.push(state);
      if (onProgress) { try { onProgress(g + 1, groups.length); } catch (_) {} }
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
  async function collectDeckImages(onProgress) {
    const presentation = activePresentation();
    const docKey = documentKey(presentation);
    const scratch = getSharedScratch();
    const pending = [];
    try {
      const slideCount = Number(presentation.Slides.Count) || 0;
      // per-slide layout name (for "applied to pages" of master/layout pics)
      const slideLayouts = {};
      for (let slideIndex = 1; slideIndex <= slideCount; slideIndex += 1) {
        try {
          const lo = presentation.Slides.Item(slideIndex).CustomLayout;
          slideLayouts[slideIndex] = lo ? String(lo.Name || "") : "";
        } catch (_) { slideLayouts[slideIndex] = ""; }
      }
      for (let slideIndex = 1; slideIndex <= slideCount; slideIndex += 1) {
        const slide = presentation.Slides.Item(slideIndex);
        const shapeCount = Number(slide.Shapes.Count) || 0;
        for (let shapeIndex = 1; shapeIndex <= shapeCount; shapeIndex += 1) {
          const shape = slide.Shapes.Item(shapeIndex);
          if (!isPicture(shape)) continue;
          const meta = parseLink(String(shape.AlternativeText || ""));
          pending.push({
            slideIndex: slideIndex,
            shapeIndex: shapeIndex,
            shape: shape,
            slide: slide,
            meta: meta,
            kind: "slide",
            layoutName: "",
            appliedPages: [slideIndex],
            name: meta && meta.name ? meta.name : String(shape.Name || "图片")
          });
        }
      }
      // SlideMaster pictures (applied to every slide)
      let master = null;
      try { master = presentation.SlideMaster; } catch (_) {}
      const layoutsMeta = [];
      if (master && master.Shapes) {
        const mc = Number(master.Shapes.Count) || 0;
        for (let mi = 1; mi <= mc; mi += 1) {
          const shape = master.Shapes.Item(mi);
          if (!isPicture(shape)) continue;
          const meta = parseLink(String(shape.AlternativeText || ""));
          pending.push({
            slideIndex: 0,
            shapeIndex: mi,
            shape: shape,
            slide: null,
            meta: meta,
            kind: "master",
            layoutName: "",
            appliedPages: slideCount > 0 ? Array.from({ length: slideCount }, function (_, k) { return k + 1; }) : [],
            name: meta && meta.name ? meta.name : String(shape.Name || "母版图片")
          });
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
                  if (!isPicture(shape)) continue;
                  const meta = parseLink(String(shape.AlternativeText || ""));
                  const applied = [];
                  for (let si2 = 1; si2 <= slideCount; si2 += 1) {
                    if (slideLayouts[si2] === loName) applied.push(si2);
                  }
                  pending.push({
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
                  });
                }
              }
            }
          }
        } catch (_) {}
      }
      const total = pending.length;
      const groups = [];
      const groupsByKey = new Map();
      const thumbCache = {};
      // Batch-render every picture that cannot be satisfied by a known
      // fingerprint: grid the shapes onto one scratch slide and export the
      // whole grid once per 36 pictures instead of exporting each picture.
      const renderQueue = [];
      for (let i = 0; i < pending.length; i += 1) {
        const item = pending[i];
        const cacheKey = item.meta && item.meta.contentFp ? item.meta.contentFp : "";
        if (cacheKey && thumbCache[cacheKey]) {
          item._info = thumbCache[cacheKey];
        } else {
          renderQueue.push(i);
        }
      }
      let batchResults = [];
      let fallbackCount = 0;
      if (renderQueue.length) {
        batchResults = await renderThumbnailsBatch(renderQueue.map(function (i) { return pending[i]; }), scratch);
      }
      for (let q = 0; q < renderQueue.length; q += 1) {
        const item = pending[renderQueue[q]];
        let info = batchResults[q];
        if (!info || !info.fp) {
          // Graceful fallback: single-shape export path (older WPS / no canvas).
          info = await contentFingerprintAndThumb(item.shape, scratch);
          fallbackCount += 1;
        }
        if (info && info.fp && !thumbCache[info.fp]) thumbCache[info.fp] = info;
        item._info = info || { fp: "", dataUrl: "", dHash: "", w: 0, h: 0, nw: 0, nh: 0, aspect: 0 };
      }
      for (let i = 0; i < pending.length; i += 1) {
        const item = pending[i];
        if (!item._info) {
          const cacheKey = item.meta && item.meta.contentFp ? item.meta.contentFp : "";
          item._info = (cacheKey && thumbCache[cacheKey]) || { fp: "", dataUrl: "", dHash: "", w: 0, h: 0, nw: 0, nh: 0, aspect: 0 };
        }
        const info = item._info;
        // Group by PERCEPTUAL hash when canvas decoding is available (same
        // source survives re-encoding/resampling), falling back to the strict
        // byte fingerprint. Aspect ratio is display-only and never splits.
        const groupKey = info.dHash ? "d:" + info.dHash : (info.fp ? "f:" + info.fp : "");
        let group = groupsByKey.get(groupKey);
        if (!group && info.dHash) {
          for (let gi = 0; gi < groups.length; gi += 1) {
            const gk = groups[gi].key;
            if (gk && gk.indexOf("d:") === 0 && dHashDistance(gk.slice(2), info.dHash) <= 2) {
              group = groups[gi];
              break;
            }
          }
        }
        if (!group && (info.fp || info.dHash)) {
          group = {
            key: groupKey,
            name: item.name,
            src: item.meta ? item.meta.src : "",
            fileFp: item.meta ? item.meta.fileFp : "",
            contentFp: info.fp,
            dHash: info.dHash,
            aspect: info.aspect,
            instances: []
          };
          if (groupKey) groupsByKey.set(groupKey, group);
          groups.push(group);
        }
        if (!group) {
          group = groupsByKey.get("");
          if (!group) {
            group = { key: "", name: "无法识别", src: "", fileFp: "", contentFp: "", aspect: 0, instances: [] };
            groupsByKey.set("", group);
            groups.push(group);
          }
        }
        const isTemplate = item.kind !== "slide";
        group.instances.push({
          uid: item.kind === "master" ? "M:" + item.shapeIndex : (item.kind === "layout" ? "L:" + item.layoutName + ":" + item.shapeIndex : item.slideIndex + ":" + item.shapeIndex),
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
          aspect: item.meta ? item.meta.aspect : info.aspect,
          userAlt: item.meta ? item.meta.userAlt : "",
          thumb: info.dataUrl,
          thumbW: info.w,
          thumbH: info.h
        });
        if (onProgress) { try { onProgress(i + 1, total); } catch (_) {} }
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
      Object.keys(bySlide).forEach(function (slideKey) {
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
          }
        }
      });
      for (let g = 0; g < groups.length; g += 1) {
        groups[g].linkState = "checking";
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
    return true;
  }

  function gotoSlide(slideIndex) {
    const windowObject = application().ActiveWindow;
    if (!windowObject || !windowObject.View) return false;
    try { windowObject.View.GotoSlide(Number(slideIndex)); return true; } catch (_) { return false; }
  }

  function gotoMasterView() {
    const windowObject = application().ActiveWindow;
    if (!windowObject) return false;
    // Verified against real WPS (2026-08): ViewType = 2 (ppViewSlideMaster)
    // switches to the slide-master view and reads back the new value;
    // Office's 11 (ppViewThumbnails) is a silent no-op in WPS. WPS does not
    // define global PpViewType constants in the add-in JS context, so use the
    // numeric value directly.
    try {
      windowObject.ViewType = 2;
      const deadline = Date.now() + 400;
      let after = Number(windowObject.ViewType);
      while (after !== 2 && Date.now() < deadline) after = Number(windowObject.ViewType);
      return after === 2;
    } catch (_) { return false; }
  }

  function selectMasterShape(shapeIndex) {
    const presentation = activePresentation();
    try {
      const master = presentation.SlideMaster;
      if (!master || !master.Shapes) return false;
      const shape = master.Shapes.Item(Number(shapeIndex));
      if (!hasMethod(shape, "Select")) return false;
      shape.Select();
      return true;
    } catch (_) { return false; }
  }

  function selectLayoutShape(layoutIndex, shapeIndex) {
    const presentation = activePresentation();
    try {
      const layouts = presentation.SlideMaster && presentation.SlideMaster.CustomLayouts;
      if (!layouts) return false;
      const layout = layouts.Item(Number(layoutIndex));
      let selected = false;
      try { if (hasMethod(layout, "Select")) { layout.Select(); selected = true; } } catch (_) {}
      try {
        const shape = layout.Shapes.Item(Number(shapeIndex));
        if (hasMethod(shape, "Select")) { shape.Select(); selected = true; }
      } catch (_) {}
      return selected;
    } catch (_) { return false; }
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

  function openPane(fragment, title) {
    const pane = application().CreateTaskPane(addinUrl(fragment), title);
    if (!pane) throw new Error("无法创建 WPS 任务窗格，请确认已启用 JS 加载项。");
    pane.Visible = true;
  }

  function runAsync(work) {
    Promise.resolve().then(work).catch(function (error) { tell(error && error.message ? error.message : error); });
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
    if (!spec || !spec.reportPath) {
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
  function OnGetFileImage() { return "icon_file.png"; }
  function OnGetFileAllImage() { return "icon_file_all.png"; }
  function OnGetClipboardImage() { return "icon_clipboard.png"; }
  function OnGetClipboardAllImage() { return "icon_clipboard_all.png"; }
  function OnGetInfoImage() { return "icon_info.png"; }
  var RIBBON_ICON_BY_ID = {
    OpenPicturePanelButton: "icon.png",
    CtxOpenPanel: "icon.png",
    ReplacePictureFile: "icon_file.png",
    ReplaceAllFile: "icon_file_all.png",
    CtxReplaceFile: "icon_file.png",
    CtxReplaceAllFile: "icon_file_all.png",
    ReplacePictureClipboard: "icon_clipboard.png",
    ReplaceAllClipboard: "icon_clipboard_all.png",
    CtxReplaceClipboard: "icon_clipboard.png",
    CtxReplaceAllClipboard: "icon_clipboard_all.png",
    PictureReplaceCompatibility: "icon_info.png"
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

  function OnAddInLoad() { runAsync(async function () { await maybeRunProfile(); await maybeRunViewProbe(); await maybeRunSelfTest(); }); }
  function OpenPicturePanel() {
    runAsync(function () { openPane("#panel", "图片清单"); });
  }
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
      if (openProgressDialog("replaceAllFromFile|" + encodeURIComponent(path))) return;
      const result = await runBatchWithProgress("批量文件替换", function (onProgress, cancelled) {
        return replaceAllFromFile(path, onProgress, cancelled);
      });
      tell(formatBatchResult(result) + (result && result.cancelled ? "（已取消）" : ""), "鍥剧墖鍘熶綅鏇挎崲");
    });
  }
  function ReplaceSelectedFromClipboard() { runAsync(async function () { await replaceSelectedFromClipboard(); tell("剪贴板原位替换完成。"); }); }
  function formatBatchResult(result) {
    return "批量替换完成：匹配 " + result.matched + " 张，成功 " + result.success + " 张，失败 " + result.failed + " 张。";
  }
  function ReplaceAllFromClipboard() {
    runAsync(async function () {
      if (openProgressDialog("replaceAllFromClipboard")) return;
      const result = await runBatchWithProgress("批量剪贴板替换", function (onProgress, cancelled) {
        return replaceAllFromClipboard(onProgress, cancelled);
      });
      tell(formatBatchResult(result) + (result && result.cancelled ? "（已取消）" : ""), "鍥剧墖鍘熶綅鏇挎崲");
    });
  }

  global.OnAddInLoad = OnAddInLoad;
  global.OpenPicturePanel = OpenPicturePanel;
  global.OnGetPicturePanelImage = OnGetPicturePanelImage;
  global.OnGetRibbonImage = OnGetRibbonImage;
  global.OnGetPanelImage = OnGetPanelImage;
  global.OnGetFileImage = OnGetFileImage;
  global.OnGetFileAllImage = OnGetFileAllImage;
  global.OnGetClipboardImage = OnGetClipboardImage;
  global.OnGetClipboardAllImage = OnGetClipboardAllImage;
  global.OnGetInfoImage = OnGetInfoImage;
  global.OpenSingleFilePane = OpenSingleFilePane;
  global.OpenBatchFilePane = OpenBatchFilePane;
  global.ReplaceSelectedFromClipboard = ReplaceSelectedFromClipboard;
  global.ReplaceAllFromClipboard = ReplaceAllFromClipboard;
  global.ShowCompatibilityStatus = ShowCompatibilityStatus;
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
    addinUrl: addinUrl,
    collectDeckImages: collectDeckImages,
    refreshLinkStates: refreshLinkStates,
    replaceInstances: replaceInstances,
    updateLinkedInstances: updateLinkedInstances,
    requestCancelTask: requestCancelTask,
    taskCancelled: taskCancelled,
    readTaskState: readTaskState,
    unlinkInstances: unlinkInstances,
    renameShape: renameShape,
    gotoSlide: gotoSlide,
    gotoMasterView: gotoMasterView,
    selectMasterShape: selectMasterShape,
    selectLayoutShape: selectLayoutShape,
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
