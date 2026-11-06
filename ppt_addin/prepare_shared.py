"""Build the PowerPoint frontend from the checked WPS source without changing it.

Run with Python 3.10+: python ppt_addin/prepare_shared.py
Host adapters are maintained separately and are never overwritten by this script.
"""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
import re
import shutil
import xml.etree.ElementTree as ET


HERE = Path(__file__).resolve().parent
SOURCE = HERE.parent / "wps_addin"
DEST = HERE / "shared"
VERSION = "3.0.0"
EXPECTED_WPS_VERSION = "2.1.15"
OFFICE_NS = "http://schemas.microsoft.com/office/2009/07/customui"


def replace_once(text: str, old: str, new: str, label: str) -> str:
    count = text.count(old)
    if count != 1:
        raise ValueError(f"{label}: expected one exact source block, found {count}")
    return text.replace(old, new, 1)


def replace_function(text: str, name: str, expected_sha256: str, new: str) -> str:
    """Replace only a complete, hash-checked source block; reject source drift."""
    matches = list(re.finditer(r"^  (?:async )?function " + re.escape(name) + r"\(", text, re.M))
    if len(matches) != 1:
        raise ValueError(f"{name}: expected one function declaration")
    start = matches[0]
    following = re.search(r"^  (?:async )?function ", text[start.end():], re.M)
    if following is None:
        raise ValueError(f"{name}: missing next function boundary")
    end = start.end() + following.start()
    old = text[start.start():end]
    digest = hashlib.sha256(old.encode("utf-8")).hexdigest()
    if digest != expected_sha256:
        raise ValueError(f"{name}: source changed ({digest}); review the port before generating")
    return replace_once(text, old, new.strip("\n") + "\n\n", name)


TEXT_REPLACE = r'''
  function designReplaceTextRange(range, findText, replaceText, options) {
    const before = String(range.Text || "");
    const needle = String(findText || "");
    const replacement = String(replaceText == null ? "" : replaceText);
    const opts = options || {};
    if (!needle) return { count: 0, changed: false, manual: false, remaining: 0 };
    // COM TextRange.Replace has different positional arguments from WPS. Compute
    // literal matches once, then edit backwards to preserve unaffected text runs.
    // Unicode letters, marks, digits and connector characters are word members.
    const escaped = needle.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const pattern = new RegExp(escaped, opts.matchCase ? "gu" : "giu");
    const matches = [];
    let match = null;
    while ((match = pattern.exec(before)) !== null) {
      const at = match.index;
      const end = at + match[0].length;
      const whole = !opts.wholeWord ||
        (!/[\p{L}\p{N}\p{M}\p{Pc}]$/u.test(before.slice(0, at)) &&
         !/^[\p{L}\p{N}\p{M}\p{Pc}]/u.test(before.slice(end)));
      if (whole && match[0] !== replacement) matches.push({ at: at, length: match[0].length });
    }
    if (!matches.length) return { count: 0, changed: false, manual: false, remaining: 0 };
    if (hasMethod(range, "Characters")) {
      let count = 0;
      for (let i = matches.length - 1; i >= 0; i -= 1) {
        try {
          const slice = range.Characters(matches[i].at + 1, matches[i].length);
          if (!slice) break;
          slice.Text = replacement;
          count += 1;
        } catch (_) { break; }
      }
      // Never overwrite the entire paragraph after a partial slice failure.
      if (!count) throw new Error("已找到匹配文字，但 PowerPoint 拒绝修改文本。");
      return { count: count, changed: count > 0, manual: false, remaining: matches.length - count };
    }
    let after = before;
    for (let i = matches.length - 1; i >= 0; i -= 1) {
      const item = matches[i];
      after = after.slice(0, item.at) + replacement + after.slice(item.at + item.length);
    }
    try {
      range.Text = after;
      return { count: matches.length, changed: true, manual: true, remaining: 0 };
    } catch (_) { throw new Error("已找到匹配文字，但 PowerPoint 拒绝修改文本。"); }
  }
'''

SAFE_BLANK = r'''
  function designSlideHasVisibleContent(slide) {
    // Deliberately conservative: hidden objects are still user content. If any
    // inspection fails, keep the slide instead of treating an API error as empty.
    try {
      if (Number(slide.Shapes.Count) !== 0) return true;
      const notes = slide.NotesPage.Shapes;
      for (let i = 1; i <= Number(notes.Count); i += 1) {
        const shape = notes.Item(i);
        if (isTrue(shape.HasTextFrame) &&
            String(shape.TextFrame.TextRange.Text || "").trim()) return true;
      }
      const timeline = slide.TimeLine;
      if (Number(timeline.MainSequence.Count) !== 0) return true;
      const interactive = timeline.InteractiveSequences;
      if (Number(interactive.Count) !== 0) return true;
      return false;
    } catch (_) { return true; }
  }
'''

EXTRACT_SLIDES = r'''
  async function designExtractSlides(scope) {
    // Capture every source-dependent value before Add changes ActivePresentation.
    // No clipboard operations or asynchronous yields are used during extraction.
    const source = activePresentation();
    const slides = designSlides(scope || "selected");
    if (!slides.length) throw new Error("请先选择要提取的幻灯片。");
    const sourceCount = Number(source.Slides.Count);
    if (!Number.isInteger(sourceCount) || sourceCount < 1) throw new Error("无法读取源文稿页数，已中止提取。");
    const indexes = [];
    slides.forEach(function (slide) {
      const index = Number(slide.SlideIndex || slide.slideIndex);
      if (!Number.isInteger(index) || index < 1 || index > sourceCount) throw new Error("无法读取所选幻灯片的页码，已中止提取。");
      if (indexes.indexOf(index) < 0) indexes.push(index);
    });
    const width = Number(source.PageSetup.SlideWidth);
    const height = Number(source.PageSetup.SlideHeight);
    if (!(width > 0 && height > 0 && isFinite(width) && isFinite(height))) throw new Error("无法读取源文稿页面尺寸，已中止提取。");
    // Unknown macro state also uses a macro-capable snapshot, so saving a copy
    // never needs to discard source VBA or show a macro-loss confirmation.
    let macroCapable = true;
    try {
      const hasVba = source.HasVBProject;
      if (hasVba === false || hasVba === 0) macroCapable = false;
    } catch (_) {}
    let snapshot = "";
    for (let attempt = 0; attempt < 8; attempt += 1) {
      const candidate = tempPath("ExtractSlides").replace(/\.png$/i, macroCapable ? ".pptm" : ".pptx");
      if (!fileExists(candidate)) { snapshot = candidate; break; }
    }
    if (!snapshot) throw new Error("无法分配独立的提取临时文件。");
    const app = application();
    let target = null;
    let ownsTarget = false;
    let trackedTarget = false;
    let complete = false;
    try {
      source.SaveCopyAs(snapshot, macroCapable ? 25 : 24, MsoFalse);
      if (!fileExists(snapshot)) throw new Error("未能保存源文稿的临时副本，已中止提取。");
      target = app.Presentations.Add();
      if (!target || target === source) throw new Error("无法创建独立的目标演示文稿。");
      ownsTarget = true;
      global.__pptTrackScratch(target, true);
      trackedTarget = true;
      const targetSlides = target.Slides;
      if (!targetSlides || Number(targetSlides.Count) !== 0) throw new Error("新建目标文稿不是空文稿，已中止提取。");
      target.PageSetup.SlideWidth = width;
      target.PageSetup.SlideHeight = height;
      for (let i = 0; i < indexes.length; i += 1) {
        const before = Number(targetSlides.Count);
        const inserted = Number(targetSlides.InsertFromFile(snapshot, before, indexes[i], indexes[i]));
        if (inserted !== 1 || Number(targetSlides.Count) !== before + 1) throw new Error("第 " + indexes[i] + " 页导入失败，已撤销本次提取。");
      }
      global.__pptTrackScratch(target, false);
      trackedTarget = false;
      complete = true;
      return { ok: true, count: indexes.length, message: "已提取 " + indexes.length + " 页到新演示文稿，请手动保存。" };
    } finally {
      if (!complete && ownsTarget) {
        let closed = false;
        try { target.Saved = MsoTrue; } catch (_) {}
        try { target.Close(); closed = true; } catch (_) {}
        if (closed && trackedTarget) { try { global.__pptTrackScratch(target, false); } catch (_) {} }
      }
      // This unique path was absent before SaveCopyAs; partial saves are ours too.
      removeFile(snapshot);
    }
  }
'''

SAFE_SCRATCH = r'''
  function createScratchManager() {
    let presentation = null;
    let slide = null;
    let ownsPresentation = false;
    function ensure() {
      if (slide) return slide;
      const userPresentation = activePresentation();
      try {
        const presentations = application().Presentations;
        if (!presentations || !hasMethod(presentations, "Add")) throw new Error("no Presentations.Add");
        const created = presentations.Add(MsoFalse);
        if (!created || created === userPresentation) throw new Error("no separate scratch presentation");
        presentation = created;
        ownsPresentation = true;
        // The adapter unwraps its JS proxy before handing the owned COM document
        // to the native shutdown registry. Missing tracking fails this operation.
        global.__pptTrackScratch(presentation, true);
        const slides = presentation.Slides;
        slide = slides.Add(Number(slides.Count) + 1, 12);
        return slide;
      } catch (error) {
        dispose();
        throw new Error("无法创建独立的 PowerPoint 临时文稿；已中止，不修改当前文稿。 " + String(error && error.message || error));
      }
    }
    function clear() {
      if (!slide) return;
      const shapes = slide.Shapes;
      for (let i = Number(shapes.Count) || 0; i >= 1; i -= 1) {
        try { shapes.Item(i).Delete(); } catch (_) {}
      }
      // Pooled scratch documents stay open between thumbnail batches. Their
      // temporary edits must not produce a user-facing save prompt on exit.
      if (ownsPresentation && presentation) {
        try { presentation.Saved = MsoTrue; } catch (_) {}
      }
    }
    function dispose() {
      const owned = ownsPresentation;
      const p = presentation;
      presentation = null;
      slide = null;
      ownsPresentation = false;
      if (!p || !owned) return;
      // Only our independently created, hidden presentation can be discarded.
      try { p.Saved = MsoTrue; } catch (_) {}
      let closed = false;
      try { p.Close(); closed = true; } catch (_) {}
      // A failed Close remains registered for the native shutdown fallback.
      if (closed) { try { global.__pptTrackScratch(p, false); } catch (_) {} }
    }
    return { ensure: ensure, clear: clear, dispose: dispose };
  }
'''

CHECK_UPDATES = r'''
  async function checkForUpdates() {
    const result = { ok: false, current: ADDIN_VERSION, latest: "", hasUpdate: false, error: "", manifestUrl: UPDATE_MANIFEST_URL, downloadUrl: "", releasePage: UPDATE_RELEASE_PAGE };
    try {
      const response = await httpGetText(UPDATE_MANIFEST_URL, 20000);
      if (!response.ok) throw new Error(response.error || "无法读取 PowerPoint 更新清单");
      const manifest = JSON.parse(response.text);
      const latest = String(manifest && manifest.version || "");
      if (!manifest || manifest.name !== "picture-replace-tools-ppt" || !/^\d+\.\d+\.\d+$/.test(latest)) {
        throw new Error("PowerPoint 更新清单格式或产品标识不正确");
      }
      result.latest = latest;
      result.downloadUrl = releaseDownloadUrl(latest);
      result.releasePage = "https://github.com/Dongsidaye/ppt-picture-replace-tools/releases/tag/ppt-v" + latest;
      result.hasUpdate = compareVersions(latest, ADDIN_VERSION) > 0;
      result.ok = true;
    } catch (error) { result.error = String(error && error.message || error); }
    return result;
  }
'''

HOST_UPDATE = r'''
  async function updateAndRestart() {
    if (typeof global.PptHostUpdate !== "function") throw new Error("PowerPoint 更新服务尚未就绪。");
    return await global.PptHostUpdate();
  }
'''

HOST_LOAD = r'''
  function OnAddInLoad() {
    try { layerEnsureSelectionGuard(); } catch (_) {}
    try { bindPanelInventoryPreloadEvents(); } catch (_) {}
    try { schedulePanelInventoryPreload(650); } catch (_) {}
    // WPS diagnostic flag files must never trigger code against user PPT decks.
  }
'''


def brand(text: str) -> str:
    # Word boundaries preserve compatibility symbols such as CurrentWPSAddIn.
    return re.sub(r"(?<![A-Za-z0-9_])WPS(?![A-Za-z0-9_])", "PowerPoint", text)


def prepare_lifecycle(text: str) -> str:
    text = replace_once(text, '''  function layerList() {
    const selectionGuard = layerEnsureSelectionGuard();
    const context = layerCurrentContext();''', r'''  function pptOwnerDocumentKey() {
    try {
      if (typeof global.__pptDocumentKey === "function") {
        const key = String(global.__pptDocumentKey() || "");
        return key === "none" ? "" : key;
      }
      const hwnd = String(application().ActiveWindow.HWND);
      return /^-?\d+$/.test(hwnd) && Number(hwnd) !== 0 ? "window:" + hwnd : "";
    } catch (_) { return ""; }
  }

  function layerList() {
    const selectionGuard = layerEnsureSelectionGuard();
    const context = layerCurrentContext();
    const ownerDocumentKey = pptOwnerDocumentKey();''', "object document identity")
    text = replace_once(text, '''      items.push({
        shape: shape,
        shapeIndex: index,''', '''      items.push({
        ownerDocumentKey: ownerDocumentKey,
        shape: shape,
        shapeIndex: index,''', "bind listed object to document")
    text = replace_once(text, '''  function layerResolveShape(item) {
    if (item && typeof item === "object") {''', '''  function layerResolveShape(item) {
    if (item && typeof item === "object") {
      // Never fall back to a shape index when a row belongs to another window.
      if (!item.ownerDocumentKey || item.ownerDocumentKey !== pptOwnerDocumentKey()) return null;''', "reject stale object row")
    text = replace_once(text, '''  function runAsync(work) {
    Promise.resolve().then(work).catch(function (error) { tell(error && error.message ? error.message : error); });
  }''', '''  function runAsync(work) {
    return Promise.resolve().then(work).catch(function (error) { tell(error && error.message ? error.message : error); });
  }''', "return asynchronous completion")
    # These exact call forms are all Ribbon wrappers, not detached workers.
    for old, new, expected in (("    runAsync(", "    return runAsync(", 10), ("{ runAsync(", "{ return runAsync(", 11)):
        if text.count(old) != expected:
            raise ValueError("Ribbon asynchronous wrapper set changed; review return propagation")
        text = text.replace(old, new)
    text = replace_once(text, "function CheckForUpdates() { openUpdatePanel(); }", "function CheckForUpdates() { return openUpdatePanel(); }", "update wrapper completion")
    text = replace_once(text, "  let taskPane = null;\n  let lastTaskWrite = 0;", "  let taskPane = null;\n  let taskRunId = 0;\n  let lastTaskWrite = 0;", "progress generation")
    text = replace_once(text, "        running: !!state.running,", "        runId: taskRunId,\n        running: !!state.running,", "write progress generation")
    text = replace_once(text, "        running: !!parsed.running,", "        runId: Number(parsed.runId) || 0,\n        running: !!parsed.running,", "read progress generation")
    text = replace_once(text, '''      try { fileSystem().writeAsBinaryString(path, payload); } catch (_) { fileSystem().WriteFile(path, payload); }''', '''      fileSystem().WriteFile(path, payload);''', "UTF-8 task text")
    text = replace_once(text, '''      const raw = String(fileSystem().readAsBinaryString(path) || "");''', '''      const raw = String(fileSystem().ReadFile(path) || "");''', "UTF-8 task read")
    text = replace_once(text, "  function beginTask(title, total) {\n    clearCancelTask();", "  function beginTask(title, total) {\n    taskRunId += 1;\n    clearCancelTask();", "start task generation")
    text = replace_once(text, '''      if (taskPane) taskPane.Visible = true;
    } catch (_) {}
  }

  function reportTask''', '''      if (taskPane) { taskPane.__pptTaskRunId = taskRunId; taskPane.Visible = true; }
    } catch (_) {}
    return taskRunId;
  }

  function reportTask''', "return task generation")
    text = replace_once(text, '''  function openProgressPanel(title) {
    try {
      clearCancelTask();''', '''  function openProgressPanel(title) {
    try {
      taskRunId += 1;
      clearCancelTask();''', "panel task generation")
    text = replace_once(text, '''      if (!pane) return null;
      try { pane.DockPosition = 4; }''', '''      if (!pane) return null;
      pane.__pptTaskRunId = taskRunId;
      try { pane.DockPosition = 4; }''', "bind progress handle")
    text = replace_once(text, '''  function closeProgressPanel(pane) {
    try {''', '''  function closeProgressPanel(pane) {
    const finishedRunId = Number(pane && pane.__pptTaskRunId) || taskRunId;
    if (finishedRunId !== taskRunId) return;
    try {''', "panel completion ownership")
    text = replace_once(text, '''      setTimeout(function () {
        try { if (pane) pane.Visible = false; }''', '''      setTimeout(function () {
        if (finishedRunId !== taskRunId) return;
        try { if (pane) pane.Visible = false; }''', "ignore obsolete panel close")
    text = replace_once(text, '''  function endTask(label, cancelled) {
    writeTaskState''', '''  function endTask(label, cancelled, expectedRunId) {
    const finishedRunId = expectedRunId === undefined ? taskRunId : expectedRunId;
    if (finishedRunId !== taskRunId) return;
    writeTaskState''', "batch completion ownership")
    text = replace_once(text, '''    setTimeout(function () {
      try { if (taskPane) taskPane.Visible = false; }''', '''    setTimeout(function () {
      if (finishedRunId !== taskRunId) return;
      try { if (taskPane) taskPane.Visible = false; }''', "ignore obsolete batch close")
    text = replace_once(text, '    beginTask(title, 0);', '    const batchRunId = beginTask(title, 0);', "capture batch generation")
    text = replace_once(text, '      endTask(cancelled ? title + "已取消" : title + "完成", cancelled);', '      endTask(cancelled ? title + "已取消" : title + "完成", cancelled, batchRunId);', "complete own batch")
    text = replace_once(text, '      endTask("任务失败", false);', '      endTask("任务失败", false, batchRunId);', "fail own batch")
    return text


def prepare_file_ownership(text: str) -> str:
    text = replace_once(text, '''  async function writeBrowserFile(file) {''', '''  const pptOwnedBrowserFiles = new Set();

  async function writeBrowserFile(file) {''', "track browser-created temporary files")
    text = replace_once(text, '''    if (!fileExists(path)) throw new Error("无法把图片写入 WPS 临时目录。");
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
  }''', '''    if (!fileExists(path)) throw new Error("无法把图片写入 PowerPoint 临时目录。");
    pptOwnedBrowserFiles.add(path);
    return path;
  }

  function releaseBrowserFile(path) {
    // File-picker inputs belong to the user, even when the path is under TEMP.
    // Only a path allocated by writeBrowserFile in this engine may be released.
    if (!pptOwnedBrowserFiles.has(path)) return false;
    removeFile(path);
    if (fileExists(path)) return false;
    pptOwnedBrowserFiles.delete(path);
    return true;
  }

  async function replaceSelectedFromFile(path) {
    return replacePictureKeepCrop(selectedPicture(), path);
  }

  async function replaceAllFromFile(path, onProgress, cancelCheck) {
    const result = await replaceAllMatching(selectedPicture(), path, onProgress, cancelCheck);
    if (!result.matched) throw new Error("没有找到匹配的原图实例。");
    return result;
  }''', "preserve all caller-owned picture inputs")
    text = replace_once(text, '''    writeBrowserFile: writeBrowserFile,''', '''    writeBrowserFile: writeBrowserFile,
    releaseBrowserFile: releaseBrowserFile,''', "owned temporary release API")
    return text


def prepare_photoshop(text: str) -> str:
    text = replace_once(text, '''  function designPhotoshopOpen(explicitPath) {
    const shapes = designSelectedShapes(1, "请先选择一张图片。");
    const shape = shapes[0];''', '''  function designPhotoshopOpen(explicitPath) {
    const shapes = designSelectedShapes(1, "请先选择一张图片。");
    const shape = shapes[0];
    const ownerDocumentKey = pptOwnerDocumentKey();
    if (!ownerDocumentKey) throw new Error("无法确定 Photoshop 编辑任务所属文稿，请重新选择图片。");''', "bind Photoshop job to active document")
    text = replace_once(text, '''    ["CropLeft", "CropRight", "CropTop", "CropBottom"].forEach(function (prop) {
      try { if (Math.abs(designReadNumber(shape, [prop, prop.toLowerCase()], 0)) > 0.01) hasCrop = true; } catch (_) {}
    });''', '''    let cropReadFailed = false;
    ["CropLeft", "CropRight", "CropTop", "CropBottom"].forEach(function (prop) {
      try {
        const crop = Number(shape.PictureFormat[prop]);
        if (!isFinite(crop)) throw new Error("unreadable crop");
        if (Math.abs(crop) > 0.01) hasCrop = true;
      } catch (_) { cropReadFailed = true; }
    });
    if (cropReadFailed) throw new Error("无法读取图片裁剪状态，已中止 Photoshop 导出以保护原图。");''', "read PowerPoint PictureFormat crop")
    text = replace_once(text, '''    try {
      shape.Export(path, 2, Math.max(64, Math.round(dispW * 2)), Math.max(64, Math.round(dispH * 2)));
      exported = fileExists(path) && exportLooksFull(path);
    } catch (_) { exported = false; }
    if (!exported) {
      removeFile(path);
      try {
        shape.Export(path, "PNG", Math.max(64, Math.round(dispW * 2)), Math.max(64, Math.round(dispH * 2)));''', '''    // PowerPoint Shape.Export renders the visible crop. Never bake that crop
    // into a PS source that will later receive the original crop again.
    if (!hasCrop) {
      try {
        shape.Export(path, 2, Math.max(64, Math.round(dispW * 2)), Math.max(64, Math.round(dispH * 2)));
        exported = fileExists(path) && exportLooksFull(path);
      } catch (_) { exported = false; }
    }
    if (!exported && !hasCrop) {
      removeFile(path);
      try {
        shape.Export(path, "PNG", Math.max(64, Math.round(dispW * 2)), Math.max(64, Math.round(dispH * 2)));''', "avoid reapplying a baked Photoshop crop")
    text = replace_once(text,
        '    designPhotoshopJobs.unshift({ shape: shape, path: path, origPath: origSaved ? origPath : "", exe: exe, at: Date.now(), reloaded: null });',
        '    designPhotoshopJobs.unshift({ ownerDocumentKey: ownerDocumentKey, shape: shape, path: path, origPath: origSaved ? origPath : "", exe: exe, at: Date.now(), reloaded: null });',
        "store Photoshop job owner")
    text = replace_once(text, '''    if (!job) throw new Error("没有可载回的 Photoshop 编辑任务。");
    if (!fileExists(job.path))''', '''    if (!job) throw new Error("没有可载回的 Photoshop 编辑任务。");
    if (!job.ownerDocumentKey || job.ownerDocumentKey !== pptOwnerDocumentKey()) throw new Error("Photoshop 编辑任务属于其他文稿，请切回原文稿后载回。");
    if (!fileExists(job.path))''', "reject cross-document Photoshop reload")
    text = replace_once(text, '''    if (!job) throw new Error("没有可撤销的 Photoshop 载回任务。");
    if (!job.reloaded)''', '''    if (!job) throw new Error("没有可撤销的 Photoshop 载回任务。");
    if (!job.ownerDocumentKey || job.ownerDocumentKey !== pptOwnerDocumentKey()) throw new Error("Photoshop 编辑任务属于其他文稿，请切回原文稿后撤销载回。");
    if (!job.reloaded)''', "reject cross-document Photoshop undo")
    return text


def prepare_clipboard_safety(text: str) -> str:
    text = replace_once(text, '''      for (let attempt = 0; attempt < 3 && after <= before; attempt += 1) {
        try { shape.Copy(); } catch (_) {}
        try { result = shapes.Paste(); } catch (_) {}''', '''      for (let attempt = 0; attempt < 3 && after <= before; attempt += 1) {
        // A failed copy must not fingerprint whatever was previously on the clipboard.
        shape.Copy();
        try { result = shapes.Paste(); } catch (_) {}''', "fail closed on preview clipboard copy failure")
    text = replace_once(text, '''  async function pasteReplacePictureKeepCrop(target) {
    const owns = createScratchManager();
    try {''', '''  async function pasteReplacePictureKeepCrop(target) {
    const owns = createScratchManager();
    let inserted = null;
    let commitStarted = false;
    try {''', "track clipboard insertion for rollback")
    text = replace_once(text, '''      const inserted = asShape(shapes.Item(after));''', '''      inserted = asShape(shapes.Item(after));''', "retain new clipboard shape reference")
    text = replace_once(text, '''      const state = captureState(target, owns);
      const natural''', '''      const state = captureState(target, owns);
      if (state.captureFailed) throw new Error("无法读取原图片裁剪信息，已中止剪贴板替换并保留原图。");
      const natural''', "reject incomplete clipboard crop snapshot")
    text = replace_once(text, '''      target.Delete();
      try { inserted.Name = state.name; }''', '''      // If deletion itself throws ambiguously, keep the completed replacement:
      // removing it as well could leave no picture when the old one was deleted.
      commitStarted = true;
      target.Delete();
      try { inserted.Name = state.name; }''', "clipboard deletion commit boundary")
    text = replace_once(text, '''      return inserted;
    } finally {
      owns.dispose();
    }
  }

  async function replaceSelectedFromClipboard''', '''      return inserted;
    } catch (error) {
      if (!commitStarted && inserted) { try { inserted.Delete(); } catch (_) {} }
      throw error;
    } finally {
      owns.dispose();
    }
  }

  async function replaceSelectedFromClipboard''', "rollback failed clipboard insertion")
    return text


def prepare_main(source: str) -> str:
    replacements = (
        ("designReplaceTextRange", "fd35ec307ee7a104e098f350f7fa9016a289381e6a5afdd1714d13f9571fc834", TEXT_REPLACE),
        ("designSlideHasVisibleContent", "3576552a6ca6a1d9ee0825a541a0c9789a5d3e02713388ace5fcbe87c513a72d", SAFE_BLANK),
        ("designExtractSlides", "bb95285c061536ea3ad5d21b1c721bde8d25895b49235bd1056476e96b7f660d", EXTRACT_SLIDES),
        ("createScratchManager", "91c40f54b166c31b308a4bb08445b2ecd33a6b183ee25e04e4f3f0b1594cad92", SAFE_SCRATCH),
        ("checkForUpdates", "3c23a493bb956a772dc6a7869c06ccd6c695592a07049a9a5cc01ec558ba66eb", CHECK_UPDATES),
        ("updateAndRestart", "a6a7ea9557dbf87b03f2185ec0e173dbdf1ec7958820e3aaced62032beab3c85", HOST_UPDATE),
        ("OnAddInLoad", "d2aee8f0ef322061c31461e75ac03eb9340cbd152ff82f5dff04715ab774c20a", HOST_LOAD),
    )
    text = source
    for name, digest, body in replacements:
        text = replace_function(text, name, digest, body)
    text = replace_once(text, "    if (!panelInventoryPanelUsed) return;\n", "", "preload first-use gate")
    text = replace_once(text, '''  // Background preloads are only worth their cost (full deck walk + thumbnail
  // rendering through clipboard Copy/Paste) in sessions where the user has
  // actually opened the picture panel.  Until then every preload trigger is
  // ignored and the pane simply scans in the foreground on first open.
  let panelInventoryPanelUsed = false;''', '''  // PowerPoint warms the list on document-open events before the first panel use.
  // The debounce and yielding scan remain active to keep editing responsive.
  let panelInventoryPanelUsed = true;''', "preload initial state")
    text = replace_once(text, '''  function designReadColorTarget(target) {''', '''  // PowerPoint COM ColorFormat.RGB is OLE BGR; the HTML controls use RRGGBB.
  function designSwapRedBlue(value) {
    const color = Number(value) >>> 0;
    return ((color & 255) << 16) | (color & 0xff00) | ((color >>> 16) & 255);
  }

  function designReadColorTarget(target) {''', "OLE conversion helper")
    text = replace_once(text,
        "      if (value !== undefined && value !== null && isFinite(Number(value))) return Number(value);\n    }\n    return null;\n  }\n\n  function designWriteColorTarget(target, value) {",
        "      if (value !== undefined && value !== null && isFinite(Number(value)) && Number(value) >= 0 && Number(value) <= 0xffffff) return designSwapRedBlue(value);\n    }\n    return null;\n  }\n\n  function designWriteColorTarget(target, value) {",
        "OLE read conversion")
    text = replace_once(text, "      if (designWritePath(target.owner, target.paths[i], value)) return true;", "      if (designWritePath(target.owner, target.paths[i], designSwapRedBlue(value))) return true;", "OLE write conversion")
    text = replace_once(text, f'  const ADDIN_VERSION = "{EXPECTED_WPS_VERSION}";', f'  const ADDIN_VERSION = "{VERSION}";', "PPT version")
    text = replace_once(text, "/wps_addin/package.json\";", "/ppt_addin/package.json\";", "PPT manifest")
    text = replace_once(text,
        '  const UPDATE_RELEASE_PAGE = "https://github.com/Dongsidaye/ppt-picture-replace-tools/releases/latest";',
        f'  const UPDATE_RELEASE_PAGE = "https://github.com/Dongsidaye/ppt-picture-replace-tools/releases/tag/ppt-v{VERSION}";', "PPT release page")
    text = replace_once(text,
        '    return UPDATE_RELEASE_BASE + "v" + String(version || "") + "/PictureReplaceTools-WPS-" + String(version || "") + ".exe";',
        '    return UPDATE_RELEASE_BASE + "ppt-v" + String(version || "") + "/PictureReplaceTools-PPT-" + String(version || "") + ".exe";', "PPT download")
    text = text.replace("画布不能选中或移动该对象。", "自动取消选中，并在选择变化时恢复位置。")
    return brand(prepare_clipboard_safety(prepare_photoshop(prepare_file_ownership(prepare_lifecycle(text)))))


def prepare_taskpane_lifecycle(text: str) -> str:
    text = replace_once(text, '''          cfg.status.textContent = "正在读取图片并替换，请稍候…";
          try {
            var path = await W.writeBrowserFile(cfg.file.files[0]);''', '''          cfg.status.textContent = "正在读取图片并替换，请稍候…";
          var path = "";
          try {
            path = await W.writeBrowserFile(cfg.file.files[0]);''', "UI owns browser-file allocation")
    text = replace_once(text, '''          } finally {
            cfg.run.disabled = false;
          }
        });
      }
      Object.keys(LEGACY)''', '''          } finally {
            if (path && W.releaseBrowserFile) W.releaseBrowserFile(path);
            cfg.run.disabled = false;
          }
        });
      }
      Object.keys(LEGACY)''', "UI releases only its own browser-file allocation")
    text = replace_once(text, '''      async function toolRun(button, work) {
        if (toolState.busy) return;''', '''      function pptUiDocumentKey() {
        try { return String(window.__pptDocumentKey() || "none"); } catch (_) { return "none"; }
      }
      async function toolRun(button, work) {
        if (toolState.busy) return;
        var expectedDocument = pptUiDocumentKey();''', "capture UI command document")
    text = replace_once(text, '''        await yieldPane();
        try {
          var result = await work();''', '''        await yieldPane();
        try {
          if (expectedDocument === "none" || expectedDocument !== pptUiDocumentKey()) throw new Error("当前文稿已切换，请在目标文稿中重新执行。");
          var result = await work();''', "validate UI command after yield")
    text = replace_once(text, '''      function armButton(button, label, armedLabel, action) {
        var timer = null;
        button.addEventListener("click", function () {
          if (button.classList.contains("armed")) {
            clearTimeout(timer);
            button.classList.remove("armed");
            button.textContent = label;
            action();
          } else {
            button.classList.add("armed");
            button.textContent = armedLabel;
            timer = setTimeout(function () {
              button.classList.remove("armed");
              button.textContent = label;
            }, 3000);
          }
        });
      }''', '''      function armButton(button, label, armedLabel, action) {
        var timer = null;
        var armedDocument = "";
        function resetConfirmation() {
          if (timer) clearTimeout(timer);
          timer = null;
          armedDocument = "";
          button.classList.remove("armed");
          button.textContent = label;
        }
        button._pptResetConfirmation = resetConfirmation;
        button.addEventListener("click", function () {
          var currentDocument = pptUiDocumentKey();
          if (currentDocument === "none") { resetConfirmation(); return; }
          if (button.classList.contains("armed") && armedDocument === currentDocument) {
            resetConfirmation();
            action();
          } else {
            resetConfirmation();
            armedDocument = currentDocument;
            button.classList.add("armed");
            button.textContent = armedLabel;
            timer = setTimeout(resetConfirmation, 3000);
          }
        });
      }''', "document-scoped destructive confirmation")
    text = replace_once(text, '''        if (!el.replaceSelected._armed) {
          el.replaceSelected._armed = true;''', '''        if (!el.replaceSelected._armed || el.replaceSelected._pptArmedDocument !== pptUiDocumentKey()) {
          el.replaceSelected._pptArmedDocument = pptUiDocumentKey();
          el.replaceSelected._armed = true;''', "document-scoped picture confirmation")
    text = replace_once(text, '''        var cancelBtn = document.getElementById("taskCancel");
        if (cancelBtn && !cancelBtn._bound) {''', '''        var cancelBtn = document.getElementById("taskCancel");
        var pollingRunId = null;
        if (cancelBtn) cancelBtn.disabled = false;
        if (cancelBtn && !cancelBtn._bound) {''', "reset progress cancel button")
    text = replace_once(text, '''        function tick() {
          var s = null;
          try { if (W && W.readTaskState) s = W.readTaskState(); } catch (_) {}''', '''        function tick() {
          if (location.hash !== "#progress") return;
          var s = null;
          try { if (W && W.readTaskState) s = W.readTaskState(); } catch (_) {}
          if (s && s.runId !== pollingRunId) {
            pollingRunId = s.runId;
            if (cancelBtn) cancelBtn.disabled = false;
          }''', "progress generation transition")
    text = replace_once(text, '''        var hashParts = hash.split("/");
        var toolsView = document.getElementById("toolsView");''', '''        var hashParts = hash.split("/");
        if (hash !== "#progress") {
          if (window.__picrenewPoll) clearTimeout(window.__picrenewPoll);
          window.__picrenewPoll = null;
          var progressView = document.getElementById("taskPane");
          if (progressView) progressView.classList.add("hidden");
        }
        var toolsView = document.getElementById("toolsView");''', "leave progress view cleanly")
    text = replace_once(text, '''      window.addEventListener("hashchange", route);''', '''      window.addEventListener("ppt-document-changed", function (event) {
        // The WebView survives window migration; all old document UI is invalid.
        refreshSeq += 1;
        layerState.refreshSeq += 1;
        Array.prototype.forEach.call(document.querySelectorAll("button"), function (button) {
          if (button._pptResetConfirmation) button._pptResetConfirmation();
        });
        if (el.replaceSelected._timer) clearTimeout(el.replaceSelected._timer);
        el.replaceSelected._armed = false;
        el.replaceSelected._pptArmedDocument = "";
        el.replaceSelected.classList.remove("armed");
        el.replaceSelected.textContent = "替换选中图片…";
        if (zoomTimer) { clearTimeout(zoomTimer); zoomTimer = null; }
        zoomEndDisarm();
        zoomSessionId = null;
        zoomSetControlsDisabled(true);
        zoomSetStatus("文稿已切换，请选择对象后重新拾取。");
        layerState.result = null;
        layerState.items = [];
        if (layerEl.list) layerEl.list.textContent = "";
        state.groups = [];
        state.selected = {};
        state.docKey = "";
        state.scanning = false;
        if (el.groups) el.groups.textContent = "";
        hideProgress();
        if (!event.detail || event.detail.current === "none") return;
        var currentRoute = (location.hash || "#idle").toLowerCase();
        if (currentRoute === "#layers") layerRefreshList(true);
        else if (currentRoute === "#panel") refresh(false);
      });
      window.addEventListener("hashchange", route);''', "reset persistent UI on document switch")
    return text


def prepare_html(source: str, *, taskpane: bool) -> str:
    text = source
    if taskpane:
        text = replace_once(text, '''        if (layerEl.view) layerEl.view.classList.add("hidden");
        if (hashParts[0] === "#tools") {''', '''        if (layerEl.view) layerEl.view.classList.add("hidden");
        if (hash === "#idle") {
          el.header.classList.add("hidden");
          el.footer.classList.add("hidden");
          el.groups.classList.add("hidden");
          el.summary.classList.add("hidden");
          Object.keys(LEGACY).forEach(function (key) { LEGACY[key].view.classList.add("hidden"); });
          var idleUpdate = document.getElementById("updateView");
          if (idleUpdate) idleUpdate.classList.add("hidden");
          var idleProgress = document.getElementById("taskPane");
          if (idleProgress) idleProgress.classList.add("hidden");
          return;
        }
        if (hashParts[0] === "#tools") {''', "idle route")
        text = prepare_taskpane_lifecycle(text)
    marker = re.findall(r'  <script src="(?:\./)?main\.js"></script>', text)
    if len(marker) != 1:
        raise ValueError("HTML: expected exactly one main.js script")
    text = replace_once(text, marker[0], '  <script src="host-adapter.js"></script>\n' + marker[0], "adapter script order")
    text = replace_once(text, "</body>", '  <script src="host-bootstrap.js"></script>\n</body>', "bootstrap script order")
    text = re.sub(r"(?<=class=\"panel-version\">)v\d+\.\d+\.\d+", "v" + VERSION, text)
    return brand(text)


def read_icon_tables(main: str) -> dict[str, dict[str, str]]:
    tables = {}
    for name in ("TOOLS_ICON_BY_ID", "RIBBON_ICON_BY_ID"):
        match = re.search(r"  var " + name + r" = \{\n(.*?)\n  \};", main, re.S)
        if not match:
            raise ValueError(f"Missing source icon table {name}")
        tables[name] = dict(re.findall(r'([A-Za-z0-9_]+): "([^"]+)"', match.group(1)))
    return tables


def prepare_ribbon(source: str, main: str) -> tuple[str, dict[str, str], dict[str, str]]:
    root = ET.fromstring(source)
    tables = read_icon_tables(main)
    simple = dict(re.findall(r'function (OnGet\w+)\(\) \{ return "([^"]+)"; \}', main))
    actions: dict[str, str] = {}
    icons: dict[str, str] = {}
    original_ids = []
    for element in root.iter():
        local = element.tag.split("}")[-1]
        element.tag = "{" + OFFICE_NS + "}" + local
        if element.get("size") == "small":
            element.set("size", "normal")
        if element.get("size") not in (None, "normal", "large"):
            raise ValueError(f"Unsupported Office ribbon size: {element.get('size')}")
        control_id = element.get("id", "")
        if control_id:
            original_ids.append(control_id)
        if control_id == "PictureReplaceToolsWps":
            element.set("id", "PictureReplaceToolsPpt")
        if control_id == "AddonVersion":
            element.set("label", "v" + VERSION)
        callback = element.get("onAction")
        if callback:
            if not control_id:
                raise ValueError("Action control without an id")
            actions[control_id] = callback
            element.set("onAction", "OnAction")
        image_callback = element.get("getImage")
        if image_callback:
            if image_callback == "OnGetToolsImage":
                filename = tables["TOOLS_ICON_BY_ID"].get(control_id)
            elif image_callback == "OnGetRibbonImage":
                filename = tables["RIBBON_ICON_BY_ID"].get(control_id)
            else:
                filename = simple.get(image_callback)
            if not filename or not (SOURCE / filename).is_file():
                raise ValueError(f"Unresolved icon: {control_id} / {image_callback}")
            icons[control_id] = filename
            element.set("getImage", "GetImage")
        for key in ("label", "screentip", "supertip"):
            if element.get(key):
                element.set(key, brand(element.get(key, "")))
        if control_id == "CheckForUpdatesButton":
            element.set("supertip", "连接 GitHub 检查 PowerPoint 版更新；安装完成后，下次正常启动 PowerPoint 生效。")
    if len(original_ids) != len(set(original_ids)):
        raise ValueError("Duplicate source ribbon id")
    ET.register_namespace("", OFFICE_NS)
    ET.indent(root, space="  ")
    ribbon = '<?xml version="1.0" encoding="UTF-8"?>\n' + ET.tostring(root, encoding="unicode") + "\n"
    return ribbon, icons, actions


def write_text(path: Path, value: str) -> None:
    path.write_text(value, encoding="utf-8", newline="\n")


def main() -> None:
    metadata = json.loads((SOURCE / "package.json").read_text(encoding="utf-8-sig"))
    if metadata.get("version") != EXPECTED_WPS_VERSION:
        raise ValueError("WPS source version changed; review parity before generating")
    source_main = (SOURCE / "main.js").read_text(encoding="utf-8-sig")
    output = {"main.js": prepare_main(source_main)}
    for name in ("taskpane.html", "index.html", "dialog_progress.html"):
        output[name] = prepare_html((SOURCE / name).read_text(encoding="utf-8-sig"), taskpane=name == "taskpane.html")
    ribbon, icons, actions = prepare_ribbon((SOURCE / "ribbon.xml").read_text(encoding="utf-8-sig"), source_main)
    output["ribbon.xml"] = ribbon
    output["icon-map.json"] = json.dumps(icons, ensure_ascii=False, indent=2) + "\n"
    output["action-map.json"] = json.dumps(actions, ensure_ascii=False, indent=2) + "\n"
    assets = sorted(path for path in SOURCE.glob("icon*") if path.is_file() and path.suffix.lower() in {".png", ".svg", ".ico"})
    source_files = [SOURCE / name for name in ("main.js", "taskpane.html", "ribbon.xml", "index.html", "dialog_progress.html", "package.json")] + assets
    manifest = {
        "product": "picture-replace-tools-ppt",
        "version": VERSION,
        "sourceVersion": EXPECTED_WPS_VERSION,
        "sourceDirectory": "wps_addin",
        "files": {path.name: hashlib.sha256(path.read_bytes()).hexdigest() for path in source_files},
    }
    output["source-manifest.json"] = json.dumps(manifest, ensure_ascii=False, indent=2) + "\n"
    parity = HERE / "PARITY.md"
    parity_content = None
    if parity.exists():
        content = parity.read_text(encoding="utf-8")
        start, end = "<!-- SOURCE_HASHES_START -->", "<!-- SOURCE_HASHES_END -->"
        if content.count(start) != 1 or content.count(end) != 1:
            raise ValueError("PARITY.md source hash markers are missing or duplicated")
        block = "\n\n| WPS source | SHA-256 |\n|---|---|\n" + "".join(
            f"| `{name}` | `{manifest['files'][name]}` |\n" for name in ("main.js", "taskpane.html", "ribbon.xml")
        ) + "\n"
        content = content[:content.index(start) + len(start)] + block + content[content.index(end):]
        parity_content = content
    # All source and documentation assertions complete before anything is written.
    DEST.mkdir(parents=True, exist_ok=True)
    for name, contents in output.items():
        write_text(DEST / name, contents)
    for path in assets:
        shutil.copyfile(path, DEST / path.name)
    if parity_content is not None:
        write_text(parity, parity_content)
    print(f"Prepared PowerPoint {VERSION} from WPS {EXPECTED_WPS_VERSION}: {len(actions)} actions, {len(icons)} icon controls, {len(assets)} image assets.")


if __name__ == "__main__":
    main()
