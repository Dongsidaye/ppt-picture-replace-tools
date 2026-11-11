(function (global) {
  "use strict";
  const host = global.__pptHost;
  const W = global.WpsPictureReplace;
  const actionMap = fetch("action-map.json").then(r => {
    if (!r.ok) throw new Error("功能区命令映射加载失败。");
    return r.json();
  });
  let busy = false;
  global.__pptDispatch = async function (id, expectedDocumentKey) {
    try {
      const map = await actionMap;
      if (!Object.prototype.hasOwnProperty.call(map, id)) throw new Error("未知功能区命令。");
      const callback = map[id];
      const command = global[callback];
      if (typeof command !== "function") throw new Error("功能入口未加载：" + map[id]);
      if (expectedDocumentKey && expectedDocumentKey !== global.__pptDocumentKey()) throw new Error("文稿已切换，请在当前文稿中重新点击此命令。");
      // These entries only navigate the persistent pane. Keep them available
      // during an asynchronous operation without releasing its mutation mutex.
      const navigation = ["OpenPicturePanel", "OpenDesignToolsPane", "OpenSelectionPane", "OpenCadImportPane"].includes(callback);
      const choosingFile = global.__pptFileDialogPending && global.__pptFileDialogPending();
      if (navigation) { await command({ Id: id }); return; }
      if (choosingFile) throw new Error("请先完成或取消当前文件选择，再执行此操作。");
      if (busy) return;
      busy = true;
      try { await command({ Id: id }); } finally { busy = false; }
    } catch (error) {
      host.Log("ribbon: " + String(error));
      global.alert(String(error.message || error));
    }
  };
  // A PPT release never executes the WPS updater. Installers are platform-scoped.
  global.PptHostUpdate = async function () {
    const info = await W.checkForUpdates();
    if (!info.ok) { global.alert("检查更新失败：" + info.error); return false; }
    if (!info.hasUpdate) { global.alert("当前已是最新版 v" + info.current); return false; }
    if (!/^\d+\.\d+\.\d+$/.test(info.latest)) throw new Error("更新版本格式无效。");
    if (!global.confirm("下载并安装 PowerPoint 版 v" + info.latest + "？安装不会强制关闭未保存的文稿。")) return false;
    const releaseResponse = await fetch("https://raw.githubusercontent.com/Dongsidaye/ppt-picture-replace-tools/agent/wps-adaptation-1-1-1/ppt_addin/release.json", { cache: "no-store" });
    if (!releaseResponse.ok) throw new Error("无法获取安装包校验清单。");
    const release = await releaseResponse.json();
    if (release.platform !== "PowerPoint" || release.version !== info.latest || release.downloadUrl !== info.downloadUrl || !/^[a-f0-9]{64}$/i.test(release.sha256)) throw new Error("安装包校验清单不匹配。");
    const path = String(host.UpdateRoot) + "\\PictureReplaceTools-PPT-" + info.latest + ".exe";
    const result = await W.downloadInstaller(info.downloadUrl, path);
    if (!result.ok) { global.alert("下载失败：" + result.error); return false; }
    const binary = global.wps.FileSystem.readAsBinaryString(path);
    const bytes = Uint8Array.from(binary, char => char.charCodeAt(0));
    const digest = Array.from(new Uint8Array(await crypto.subtle.digest("SHA-256", bytes))).map(b => b.toString(16).padStart(2, "0")).join("");
    if (digest.toLowerCase() !== release.sha256.toLowerCase()) {
      host.RemoveFile(path);
      throw new Error("安装包 SHA-256 校验失败，已停止安装。");
    }
    return Boolean(host.ShellExecute(path, ""));
  };
  // Hidden scratch documents emit NewPresentation/PresentationClose too. Only a
  // real change of the active document window invalidates the user's session.
  let documentKey = global.__pptDocumentKey();
  global.addEventListener("ppt-host-event", event => {
    if (["PresentationClose", "WindowActivate", "PresentationOpen", "NewPresentation"].includes(event.detail)) {
      const nextKey = global.__pptDocumentKey();
      if (nextKey === documentKey) return;
      const previousKey = documentKey;
      documentKey = nextKey;
      try { W.smartZoomEnd(); } catch (_) {}
      try { if (W.readTaskState && W.readTaskState().running) W.requestCancelTask(); } catch (_) {}
      global.dispatchEvent(new CustomEvent("ppt-document-changed", { detail: { previous: previousKey, current: nextKey } }));
    }
  });
  try {
    global.OnAddInLoad();
    global.__pptEngineReady = true;
    host.Log("PowerPoint engine ready v" + W.addinVersion);
  } catch (error) { global.__pptEngineError = String(error); host.Log("startup: " + String(error)); }
}(window));
