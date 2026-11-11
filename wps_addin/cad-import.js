/* WPS DWG preview, optional OCR review, and asynchronous insertion. */
(function (global) {
  "use strict";
  const host = global.WpsCadBridge;
  const choose = document.getElementById("cadImportChoose");
  if (!host || !choose) return;

  const cancel = document.getElementById("cadImportCancel");
  const insert = document.getElementById("cadImportInsert");
  const file = document.getElementById("cadImportFile");
  const status = document.getElementById("cadImportStatus");
  const warning = document.getElementById("cadImportWarning");
  const preview = document.getElementById("cadImportPreview");
  const previewWrap = document.getElementById("cadImportPreviewWrap");
  const font = document.getElementById("cadImportFont");
  let sequence = 0;
  let requestId = "";
  let owner = "";
  let phase = "idle";
  let cancelRequested = false;
  let settlePreview = null;

  function currentDocument() {
    return String(host.documentKey() || "");
  }

  function current(token, documentKey) {
    return token === sequence && documentKey && documentKey === currentDocument();
  }

  function controls(nextPhase) {
    phase = nextPhase;
    choose.disabled = ["choosing", "ocr-choice", "converting", "preview", "cancelling", "inserting"].includes(phase);
    if (font) font.disabled = ["choosing", "ocr-choice", "converting", "preview", "cancelling", "inserting"].includes(phase);
    if (global.WpsCadOcr && global.WpsCadOcr.setDisabled) global.WpsCadOcr.setDisabled(["inserting", "cancelling", "inserted"].includes(phase));
    cancel.hidden = !requestId || !["converting", "preview", "cancelling", "ready", "inserting"].includes(phase);
    cancel.disabled = phase === "cancelling";
    insert.disabled = phase !== "ready" || !requestId;
  }

  function message(text, isError) {
    status.textContent = String(text || "");
    status.className = "tool-status" + (isError ? " error" : "");
  }

  function clearPreview() {
    if (global.WpsCadOcr) global.WpsCadOcr.clear();
    if (settlePreview) settlePreview();
    preview.onload = null;
    preview.onerror = null;
    preview.removeAttribute("src");
    previewWrap.hidden = true;
    warning.textContent = "";
    warning.hidden = true;
  }

  function cancelNative(id) {
    if (!id) return;
    try { host.CancelDwgImport(id); } catch (_) {}
  }

  function reset(text) {
    sequence += 1;
    cancelNative(requestId);
    requestId = "";
    owner = "";
    cancelRequested = false;
    file.textContent = "";
    file.hidden = true;
    clearPreview();
    controls("idle");
    message(text || "请选择 DWG 文件。");
  }

  function displayPreview(base64, token, documentKey) {
    return new Promise(function (resolve, reject) {
      if (!base64 || !/^[a-zA-Z0-9+/]+={0,2}$/.test(base64)) {
        reject(new Error("转换结果没有可显示的 SVG 预览。"));
        return;
      }
      let completed = false;
      const cancelPreview = function () { finish(false); };
      function finish(shown, error) {
        if (completed) return;
        completed = true;
        if (settlePreview === cancelPreview) {
          settlePreview = null;
          preview.onload = null;
          preview.onerror = null;
        }
        if (error) reject(error);
        else resolve(shown);
      }
      settlePreview = cancelPreview;
      preview.onload = function () {
        if (!current(token, documentKey)) { finish(false); return; }
        if (!preview.naturalWidth || !preview.naturalHeight) { finish(false, new Error("SVG 预览尺寸无效。")); return; }
        previewWrap.hidden = false;
        finish(true);
      };
      preview.onerror = function () { finish(false, new Error("SVG 预览无法显示，请重新选择 DWG 文件。")); };
      preview.src = "data:image/svg+xml;base64," + base64;
    });
  }

  async function convert(token, documentKey) {
    for (;;) {
      await new Promise(resolve => global.setTimeout(resolve, 150));
      if (!current(token, documentKey)) return;
      const result = JSON.parse(String(host.PollDwgImport(requestId)));
      if (result.state === "pending") {
        if (!cancelRequested) message(result.status || "正在将 DWG 转换为 SVG…");
        continue;
      }
      if (result.state !== "done") throw new Error("无法读取 DWG 转换结果。");
      if (cancelRequested || result.cancelled) {
        requestId = "";
        controls("idle");
        message("已取消导入。");
        return;
      }
      if (!result.ok || result.error) throw new Error(result.error || "DWG 转换失败。");
      controls("preview");
      const shown = await displayPreview(String(result.previewBase64 || ""), token, documentKey);
      if (!shown || !current(token, documentKey) || cancelRequested) return;
      warning.textContent = String(result.warning || "");
      warning.hidden = !warning.textContent;
      if (global.WpsCadOcr) global.WpsCadOcr.show(result.ocr, String(result.previewBase64 || ""));
      controls("ready");
      const count = Number(result.entityCount);
      message("转换完成，" + Number(result.textCount || 0) + " 行原文字可编辑。" + (result.ocr && result.ocr.lines.length ? "请勾选并核对要恢复的 OCR 文字。" : "检查预览后可插入。"));
      return;
    }
  }

  choose.addEventListener("click", async function () {
    if (choose.disabled) return;
    reset();
    const token = sequence;
    const documentKey = currentDocument();
    if (!documentKey) { message("请先打开文稿并选择要插入的幻灯片。", true); return; }
    owner = documentKey;
    controls("choosing");
    message("请选择要导入的 DWG 文件。");
    try {
      const selected = await host.chooseFile({ kind: "file", title: "选择要导入的 DWG 文件", filter: "*.dwg" });
      if (!current(token, documentKey)) return;
      if (!selected.ok || selected.cancelled) { controls("idle"); message("已取消选择文件。"); return; }
      controls("ocr-choice");
      const ocr = global.WpsCadOcr ? await global.WpsCadOcr.choose() : false;
      if (!current(token, documentKey)) return;
      if (ocr === null) { controls("idle"); message("已取消导入。"); return; }
      const path = String(selected.path || "");
      file.textContent = path.split(/[\\/]/).pop();
      file.hidden = false;
      controls("converting");
      message("正在准备 DWG 转换…");
      await new Promise(resolve => global.setTimeout(resolve, 0));
      if (!current(token, documentKey)) return;
      requestId = String(host.BeginDwgImport(JSON.stringify({ path: path, documentKey: documentKey, ocr: ocr, font: font ? font.value : "Microsoft YaHei" })));
      controls("converting");
      await convert(token, documentKey);
    } catch (error) {
      if (!current(token, documentKey)) return;
      cancelNative(requestId);
      requestId = "";
      clearPreview();
      controls("idle");
      message(error && error.message || error, true);
    }
  });

  cancel.addEventListener("click", function () {
    if (!requestId || cancel.disabled) return;
    cancelRequested = true;
    cancelNative(requestId);
    clearPreview();
    if (phase === "ready" || phase === "preview") {
      requestId = "";
      controls("idle");
      message("已取消导入。");
    } else {
      controls("cancelling");
      message("正在取消导入…");
    }
  });

  insert.addEventListener("click", async function () {
    if (insert.disabled || phase !== "ready" || !requestId) return;
    const token = sequence;
    const documentKey = owner;
    if (!current(token, documentKey)) { reset("文稿已切换，请在当前文稿中重新选择 DWG 文件。"); return; }
    controls("inserting");
    message("正在插入 SVG…");
    try {
      await new Promise(resolve => global.setTimeout(resolve, 0));
      if (!current(token, documentKey)) return;
      const options = { font: font ? font.value : "Microsoft YaHei", ocrSelections: global.WpsCadOcr ? global.WpsCadOcr.selections() : [] };
      let result;
      if (typeof host.BeginDwgInsertion === "function") {
        result = JSON.parse(String(host.BeginDwgInsertion(requestId, JSON.stringify(options))));
        if (!result.ok) throw new Error(result.error || "开始插入失败。");
        do {
          await new Promise(resolve => global.setTimeout(resolve, 100));
          if (!current(token, documentKey)) return;
          result = JSON.parse(String(host.PollDwgInsertion(requestId)));
          if (result.state === "pending" && !cancelRequested) message("正在插入可编辑文字 " + (result.done || 0) + " / " + (result.total || 0) + "…");
        } while (result.state === "pending");
      } else result = JSON.parse(String(host.InsertDwgImport(requestId)));
      if (cancelRequested || result.cancelled) {
        requestId = "";
        clearPreview();
        controls("idle");
        message("已取消导入。");
        return;
      }
      if (!result.ok || result.error) throw new Error(result.error || "SVG 插入失败。");
      requestId = "";
      controls("inserted");
      message("已插入图形和 " + (Number(result.textCount || 0) + Number(result.ocrCount || 0)) + " 处可编辑文字。" + (result.fallbackCount ? result.fallbackCount + " 组文字保留 SVG 原样。" : ""));
    } catch (error) {
      if (!current(token, documentKey)) return;
      cancelNative(requestId);
      requestId = "";
      clearPreview();
      controls("idle");
      message(error && error.message || error, true);
    }
  });

  global.addEventListener("wps-cad-document-changed", function () {
    reset("文稿已切换，请在当前文稿中重新选择 DWG 文件。");
  });

  controls("idle");
}(window));
