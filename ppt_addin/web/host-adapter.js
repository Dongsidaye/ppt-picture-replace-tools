/* PowerPoint's IDispatch object model, projected through a trusted local WebView2. */
(function (global) {
  "use strict";
  const host = global.chrome.webview.hostObjects.sync.pptHost;
  // WebView2 represents IDispatch objects as callable proxies. Office default
  // properties (DISPID 0: TextRange.Text, ColorFormat.RGB, ...) are otherwise
  // mistaken for functions. Resolve properties through the CLR COM binder.
  const wrapped = new WeakMap();
  const originals = new WeakMap();
  let fileDialogPending = false;
  function assertNoFileDialog() {
    if (fileDialogPending) throw new Error("请先完成或取消当前文件选择，再操作文稿。");
  }
  function unwrap(value) {
    if (Array.isArray(value)) return value.map(unwrap);
    return value && (typeof value === "object" || typeof value === "function") && originals.has(value) ? originals.get(value) : value;
  }
  function wrap(raw) {
    if (!raw || (typeof raw !== "object" && typeof raw !== "function")) return raw;
    if (Array.isArray(raw)) return raw.map(wrap);
    if (wrapped.has(raw)) return wrapped.get(raw);
    const proxy = new Proxy({}, {
      get(_, key) {
        if (key === "then" || typeof key !== "string") return undefined;
        try { return wrap(host.ReadMember(raw, key)); }
        catch (propertyError) {
          let candidate;
          try { candidate = raw[key]; }
          catch (memberError) {
            // Optional WPS-style capability probes must behave like ordinary
            // JS property access when PowerPoint has no such COM member.
            if (/0x8002000[36]|unknown name|member not found|未知名称|找不到成员/i.test(String(memberError))) return undefined;
            throw propertyError;
          }
          if (typeof candidate !== "function") throw propertyError;
          return function () {
            assertNoFileDialog();
            const args = Array.from(arguments).map(unwrap);
            return wrap(Function.prototype.apply.call(candidate, raw, args));
          };
        }
      },
      set(_, key, value) { assertNoFileDialog(); host.WriteMember(raw, String(key), unwrap(value)); return true; }
    });
    wrapped.set(raw, proxy); originals.set(proxy, raw);
    return proxy;
  }
  const native = wrap(global.chrome.webview.hostObjects.sync.pptApplication);
  const listeners = new Map();
  const temp = String(host.TempFolder);
  const fs = {
    tmpdir: () => temp,
    Exists: path => Boolean(host.FileExists(String(path))),
    existsSync: path => Boolean(host.FileExists(String(path))),
    readAsBinaryString: path => global.atob(String(host.ReadBinaryBase64(String(path)))),
    writeAsBinaryString: (path, bytes) => host.WriteBinaryBase64(String(path), global.btoa(String(bytes))),
    ReadFile: path => String(host.ReadText(String(path))),
    WriteFile: (path, text) => host.WriteText(String(path), String(text)),
    Remove: path => host.RemoveFile(String(path)),
    unlinkSync: path => host.RemoveFile(String(path)),
    statSync: path => {
      const stat = JSON.parse(String(host.StatJson(String(path))));
      stat.mtime = new Date(stat.mtimeMs);
      return stat;
    }
  };
  let currentPane = null;
  function pane(url, title) {
    const parsed = new URL(String(url), location.href);
    if (parsed.origin !== location.origin || !parsed.pathname.endsWith("taskpane.html")) {
      throw new Error("任务窗格地址不受信任。");
    }
    const hash = parsed.hash || "#panel";
    let visible = false;
    let previousPane = null;
    let previousHash = "#idle";
    const instance = {
      get Visible() { return visible; },
      set Visible(value) {
        visible = Boolean(value);
        if (visible) {
          if (hash === "#progress" && location.hash !== "#progress") {
            previousPane = currentPane;
            previousHash = location.hash || "#idle";
          }
          currentPane = instance;
          host.ShowPane(String(title || "东四大爷的工具箱"));
          if (location.hash === hash) global.dispatchEvent(new Event("hashchange"));
          else location.hash = hash;
        } else if (currentPane === instance) {
          currentPane = null;
          if (hash === "#progress" && previousPane && previousPane !== instance) {
            previousPane.Visible = true;
            if (location.hash !== previousHash) location.hash = previousHash;
          } else {
            location.hash = hash === "#progress" ? previousHash : "#idle";
            host.HidePane();
          }
          previousPane = null;
        }
      },
      get Width() { return 430; },
      set Width(_) {},
      Delete() { instance.Visible = false; }
    };
    return instance;
  }
  const extensions = {
    // Synchronous Office modal dialogs inside a WebView2 host callback can
    // deadlock the script queue. All callers must use __pptChooseFile instead.
    FileDialog() { throw new Error("此文件选择入口需要更新，请使用插件的异步文件选择。"); },
    FileSystem: fs,
    Env: { GetTempPath: () => temp },
    CurrentWPSAddIn: { Path: location.origin + "/", Name: "PowerPointTools" },
    CreateTaskPane: pane,
    ApiEvent: {
      AddApiEventListener(name, callback) {
        if (typeof callback !== "function") throw new TypeError("事件需要函数。");
        if (!listeners.has(name)) listeners.set(name, new Set());
        listeners.get(name).add(callback);
      },
      RemoveApiEventListener(name, callback) {
        if (listeners.has(name)) listeners.get(name).delete(callback);
      }
    },
    OAAssist: {
      ShellExecute(path, args) {
        if (!host.ShellExecute(String(path), String(args || ""))) throw new Error("无法打开指定的文件或链接。");
      }
    },
    alert: message => global.alert(String(message))
  };
  global.wps = new Proxy(extensions, {
    get(target, key) {
      if (key in target) return target[key];
      if (typeof key !== "string" || key === "then") return undefined;
      return native[key];
    },
    set(target, key, value) {
      if (key in target) throw new Error("宿主扩展接口不可覆盖。");
      native[key] = value;
      return true;
    }
  });
  global.__pptHost = host;
  global.__pptReadSourcePreview = path => global.atob(String(host.ReadPicturePreviewBase64(String(path))));
  let pictureSourcesPromise = null;
  global.__pptLoadPictureSources = function () {
    if (pictureSourcesPromise) return pictureSourcesPromise;
    pictureSourcesPromise = (async function () {
      const key = global.__pptDocumentKey();
      const id = String(host.BeginPictureSnapshot());
      const started = Date.now();
      for (;;) {
        await new Promise(resolve => global.setTimeout(resolve, 60));
        const result = JSON.parse(String(host.PollPictureSnapshot(id)));
        if (result.state === "pending") {
          if (Date.now()-started>120000) throw new Error("图片源读取超时，请稍后刷新清单。");
          continue;
        }
        if (result.state !== "done") throw new Error("无法读取图片源结果。");
        if (result.error) throw new Error(String(result.error));
        if (result.documentKey !== key || global.__pptDocumentKey() !== key) throw new Error("文稿已切换，请重新打开图片清单。");
        return result.entries || {};
      }
    }()).finally(function () { pictureSourcesPromise = null; });
    return pictureSourcesPromise;
  };
  global.__pptFileDialogPending = function () { return fileDialogPending; };
  global.__pptChooseFile = async function (options) {
    if (fileDialogPending) throw new Error("请先完成或取消当前文件选择。");
    const input = options || {};
    const kind = input.kind || "file";
    if (kind !== "file" && kind !== "folder") throw new Error("不支持的文件选择类型。");
    const documentKey = global.__pptDocumentKey();
    fileDialogPending = true;
    try {
      const id = String(host.BeginFileDialog(JSON.stringify({
        kind, title: String(input.title || (kind === "folder" ? "选择文件夹" : "选择文件")),
        filter: String(input.filter || "*.*"), documentKey
      })));
      for (;;) {
        // Yield the JavaScript stack before native Show starts, and keep the
        // renderer free to paint/navigation while the user's dialog is open.
        await new Promise(resolve => global.setTimeout(resolve, 80));
        const result = JSON.parse(String(host.PollFileDialog(id)));
        if (result.state === "pending") continue;
        if (result.state !== "done") throw new Error("无法读取文件选择结果，请重试。");
        if (result.error) throw new Error(String(result.error));
        if (documentKey !== global.__pptDocumentKey()) throw new Error("文稿已切换，已取消本次操作，请重新选择。");
        return { ok: result.ok === true, cancelled: result.cancelled === true, path: result.ok === true ? String(result.path || "") : "" };
      }
    } finally { fileDialogPending = false; }
  };
  global.__pptTrackScratch = function (presentation, add) {
    if (add) host.TrackScratch(unwrap(presentation));
    else host.ReleaseScratch(unwrap(presentation));
  };
  global.__pptSameObject = function (first, second) {
    try { return !!host.SameObject(unwrap(first), unwrap(second)); } catch (_) { return false; }
  };
  global.__pptDocumentKey = function () {
    try {
      const key = String(host.DocumentKey || "");
      if (/^document:[a-f0-9]{32}:window:-?[1-9]\d*$/.test(key)) return key;
    } catch (_) {}
    // A window handle or file name alone can be reused by another document.
    return "";
  };
  global.__pptEmit = function (name) {
    if (global.__pptTesting) return;
    const handlers = listeners.get(String(name));
    if (handlers) handlers.forEach(fn => {
      try { fn(); } catch (error) { host.Log("event " + name + ": " + String(error)); }
    });
    global.dispatchEvent(new CustomEvent("ppt-host-event", { detail: String(name) }));
  };
  global.addEventListener("error", event => host.Log("script: " + event.message));
  global.addEventListener("unhandledrejection", event => host.Log("promise: " + String(event.reason)));
}(window));
