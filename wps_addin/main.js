/* Picture Replace Tools - WPS WPP JavaScript add-in.
 * The implementation uses only the WPP JSAPI documented by WPS:
 * Shapes.AddPicture, PictureFormat.Crop, Shape.SaveAsPicture and
 * Shapes/View.PasteSpecial. It intentionally does not depend on PPAM/VBA.
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
  const PREVIEW_SIZE = 32;

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

  function capabilityProbe() {
    const result = {
      host: "WPS WPP JSAPI",
      application: false,
      version: "",
      activePresentation: false,
      fileSystem: false,
      addPicture: false,
      crop: false,
      pasteSpecial: false,
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
        const slide = presentation.Slides.Item(1);
        result.addPicture = !!slide.Shapes && hasMethod(slide.Shapes, "AddPicture");
        result.pasteSpecial = !!slide.Shapes && hasMethod(slide.Shapes, "PasteSpecial");
        try {
          const shape = slide.Shapes.Item(1);
          result.crop = !!shape && !!shape.PictureFormat && !!shape.PictureFormat.Crop;
        } catch (_) {}
      }
    } catch (error) {
      result.errors.push(error && error.message ? String(error.message) : String(error));
    }
    result.ready = result.application && result.fileSystem && result.addPicture && result.crop;
    return result;
  }

  function capabilityText() {
    const c = capabilityProbe();
    const yes = value => value ? "可用" : "不可用";
    return [
      c.host + (c.version ? " " + c.version : ""),
      "FileSystem: " + yes(c.fileSystem),
      "AddPicture: " + yes(c.addPicture),
      "PictureFormat.Crop: " + yes(c.crop),
      "PasteSpecial: " + yes(c.pasteSpecial),
      "CreateTaskPane: " + yes(c.taskPane),
      "FileDialog: " + yes(c.fileDialog),
      c.currentAddInPath ? "AddIn.Path: " + c.currentAddInPath : "",
      c.currentAddInName ? "AddIn.Name: " + c.currentAddInName : "",
      "结论: " + (c.ready ? "核心替换 API 已就绪" : "当前环境不满足核心替换 API"),
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
    const shape = selection.ShapeRange.Item(1);
    try { void shape.PictureFormat; } catch (_) { throw new Error("当前选中对象不是图片。"); }
    return shape;
  }

  function slideOf(shape) {
    const slide = shape && shape.Parent;
    if (!slide || !slide.Shapes) throw new Error("无法取得图片所在幻灯片。");
    return slide;
  }

  function isPicture(shape) {
    try { void shape.PictureFormat; return true; } catch (_) { return false; }
  }

  function isTrue(value) {
    return value === true || value === MsoTrue || value === 1;
  }

  function captureState(shape) {
    const crop = shape.PictureFormat.Crop;
    return {
      left: Number(shape.Left), top: Number(shape.Top),
      width: Number(shape.Width), height: Number(shape.Height),
      rotation: Number(shape.Rotation) || 0,
      flipH: isTrue(shape.HorizontalFlip), flipV: isTrue(shape.VerticalFlip),
      lockAspectRatio: shape.LockAspectRatio,
      name: String(shape.Name || ""),
      alternativeText: String(shape.AlternativeText || ""),
      zOrder: Number(shape.ZOrderPosition) || 1,
      cropShapeW: Number(crop.ShapeWidth), cropShapeH: Number(crop.ShapeHeight),
      cropPicW: Number(crop.PictureWidth), cropPicH: Number(crop.PictureHeight),
      cropOffX: Number(crop.PictureOffsetX), cropOffY: Number(crop.PictureOffsetY)
    };
  }

  function clamp(value, minimum, maximum) {
    if (maximum < minimum) return 0;
    return Math.max(minimum, Math.min(maximum, value));
  }

  function applyPreservedCrop(shape, state, naturalW, naturalH) {
    if (!(naturalW > 0 && naturalH > 0)) throw new Error("新图片尺寸无效。");
    const oldAspect = state.cropPicW / state.cropPicH;
    const newAspect = naturalW / naturalH;
    const aspectChange = newAspect / oldAspect;
    let pictureW;
    let pictureH;
    let offsetX;
    let offsetY;

    if (aspectChange >= 0.8 && aspectChange <= 1.25) {
      pictureW = state.cropPicW;
      pictureH = state.cropPicH;
      offsetX = state.cropOffX;
      offsetY = state.cropOffY;
    } else {
      const frameAspect = state.cropShapeW / state.cropShapeH;
      if (newAspect >= frameAspect) {
        pictureH = state.cropShapeH;
        pictureW = pictureH * newAspect;
      } else {
        pictureW = state.cropShapeW;
        pictureH = pictureW / newAspect;
      }
      const oldBaseH = oldAspect >= frameAspect ? state.cropShapeH : state.cropShapeW / oldAspect;
      const oldBaseW = oldAspect >= frameAspect ? oldBaseH * oldAspect : state.cropShapeW;
      let zoom = Math.max(state.cropPicW / oldBaseW, state.cropPicH / oldBaseH, 1);
      zoom = Math.min(zoom, 100);
      pictureW *= zoom;
      pictureH *= zoom;
      offsetX = (state.cropOffX / state.cropPicW) * pictureW;
      offsetY = (state.cropOffY / state.cropPicH) * pictureH;
      offsetX = clamp(offsetX, -(pictureW - state.cropShapeW) / 2, (pictureW - state.cropShapeW) / 2);
      offsetY = clamp(offsetY, -(pictureH - state.cropShapeH) / 2, (pictureH - state.cropShapeH) / 2);
    }

    shape.LockAspectRatio = MsoFalse;
    const crop = shape.PictureFormat.Crop;
    crop.ShapeWidth = state.cropShapeW;
    crop.ShapeHeight = state.cropShapeH;
    crop.PictureWidth = pictureW;
    crop.PictureHeight = pictureH;
    crop.PictureOffsetX = offsetX;
    crop.PictureOffsetY = offsetY;
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

  function replacePictureKeepCrop(oldShape, imagePath) {
    const state = captureState(oldShape);
    const slide = slideOf(oldShape);
    let inserted = null;
    try {
      // Width/Height are intentionally omitted: WPS inserts at natural size.
      inserted = slide.Shapes.AddPicture(imagePath, MsoFalse, MsoTrue, state.left, state.top);
      const naturalW = Number(inserted.Width);
      const naturalH = Number(inserted.Height);
      applyPreservedCrop(inserted, state, naturalW, naturalH);
      if (state.flipH) inserted.Flip(MsoFlipHorizontal);
      if (state.flipV) inserted.Flip(MsoFlipVertical);
      inserted.Rotation = state.rotation;
      inserted.Left = state.left;
      inserted.Top = state.top;
      inserted.LockAspectRatio = state.lockAspectRatio;
      copyCosmetics(oldShape, inserted);
      let guard = 0;
      while (Number(inserted.ZOrderPosition) > state.zOrder + 1 && guard < 1000) {
        inserted.ZOrder(MsoSendBackward);
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

  function removeFile(path) {
    if (!path) return;
    try { fileSystem().unlinkSync(path); return; } catch (_) {}
    try { fileSystem().Remove(path); } catch (_) {}
  }

  function exportUncroppedPreview(shape, path, extraFlipH, extraFlipV) {
    let probe = null;
    try {
      probe = shape.Duplicate();
      const crop = probe.PictureFormat.Crop;
      const pictureW = Number(crop.PictureWidth);
      const pictureH = Number(crop.PictureHeight);
      if (!(pictureW > 0 && pictureH > 0)) throw new Error("图片原始尺寸无效。");
      crop.ShapeWidth = pictureW;
      crop.ShapeHeight = pictureH;
      crop.PictureOffsetX = 0;
      crop.PictureOffsetY = 0;
      if (isTrue(probe.HorizontalFlip)) probe.Flip(MsoFlipHorizontal);
      if (isTrue(probe.VerticalFlip)) probe.Flip(MsoFlipVertical);
      if (extraFlipH) probe.Flip(MsoFlipHorizontal);
      if (extraFlipV) probe.Flip(MsoFlipVertical);
      probe.Rotation = 0;
      probe.LockAspectRatio = MsoFalse;
      probe.Left = 0;
      probe.Top = 0;
      probe.Width = 128;
      probe.Height = 128;
      try { probe.Line.Visible = MsoFalse; } catch (_) {}
      try { probe.Shadow.Visible = MsoFalse; } catch (_) {}
      try { probe.Glow.Radius = 0; } catch (_) {}
      try { probe.SoftEdge.Radius = 0; } catch (_) {}
      removeFile(path);
      probe.SaveAsPicture(path);
      probe.Delete();
      probe = null;
      return fileSystem().Exists(path);
    } catch (error) {
      try { if (probe) probe.Delete(); } catch (_) {}
      removeFile(path);
      return false;
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
          canvas.width = PREVIEW_SIZE; canvas.height = PREVIEW_SIZE;
          const context = canvas.getContext("2d");
          context.drawImage(image, 0, 0, PREVIEW_SIZE, PREVIEW_SIZE);
          const pixels = context.getImageData(0, 0, PREVIEW_SIZE, PREVIEW_SIZE).data;
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

  async function signaturesForShape(shape, pathPrefix) {
    const signatures = [];
    const variants = [[false, false], [true, false], [false, true], [true, true]];
    for (let i = 0; i < variants.length; i += 1) {
      const path = tempPath(pathPrefix + i);
      if (!exportUncroppedPreview(shape, path, variants[i][0], variants[i][1])) { removeFile(path); continue; }
      signatures.push(await signatureFromPath(path));
      removeFile(path);
    }
    return signatures;
  }

  async function replaceAllMatching(referenceShape, imagePath) {
    const presentation = activePresentation();
    const referenceSignatures = new Set(await signaturesForShape(referenceShape, "reference"));
    if (!referenceSignatures.size) throw new Error("无法读取参考图片。");
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
        if (!exportUncroppedPreview(candidates[i], path, false, false)) continue;
        const signature = await signatureFromPath(path);
        if (referenceSignatures.has(signature)) matches.push(candidates[i]);
      } finally { removeFile(path); }
    }

    let success = 0;
    for (let i = 0; i < matches.length; i += 1) {
      try { if (replacePictureKeepCrop(matches[i], imagePath)) success += 1; } catch (_) {}
    }
    return { matched: matches.length, success: success, failed: matches.length - success };
  }

  function pasteClipboardAsPng(targetShape, path) {
    const slide = slideOf(targetShape);
    const before = Number(slide.Shapes.Count);
    let pasted = false;
    const formats = [PP_PASTE_PNG, PP_PASTE_BITMAP, PP_PASTE_JPG, PP_PASTE_GIF];
    for (let i = 0; i < formats.length && !pasted; i += 1) {
      try { slide.Shapes.PasteSpecial(formats[i]); pasted = true; } catch (_) {}
    }
    if (!pasted) throw new Error("剪贴板没有可粘贴的图片格式。");
    const after = Number(slide.Shapes.Count);
    if (after <= before) throw new Error("剪贴板粘贴没有生成图片对象。");
    try {
      const pastedShape = slide.Shapes.Item(after);
      pastedShape.SaveAsPicture(path);
      if (!fileSystem().Exists(path)) throw new Error("无法导出剪贴板图片。");
    } finally {
      for (let i = Number(slide.Shapes.Count); i > before; i -= 1) {
        try { slide.Shapes.Item(i).Delete(); } catch (_) {}
      }
    }
    return path;
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
    if (!fileSystem().Exists(path)) throw new Error("无法把图片写入 WPS 临时目录。");
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

  async function replaceSelectedFromClipboard() {
    const target = selectedPicture();
    const path = tempPath("clipboard_single");
    try { pasteClipboardAsPng(target, path); return replacePictureKeepCrop(target, path); }
    finally { removeFile(path); }
  }

  async function replaceAllFromClipboard() {
    const reference = selectedPicture();
    const path = tempPath("clipboard_batch");
    try {
      pasteClipboardAsPng(reference, path);
      const result = await replaceAllMatching(reference, path);
      if (!result.matched) throw new Error("没有找到匹配的原图实例。");
      return result;
    } finally { removeFile(path); }
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
    capabilityProbe: capabilityProbe,
    capabilityText: capabilityText,
    formatBatchResult: formatBatchResult,
    chooseImageFile: chooseImageFile
  };
}(typeof window !== "undefined" ? window : globalThis));
