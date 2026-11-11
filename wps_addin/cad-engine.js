/* Local CAD jobs and transactional insertion in the current WPS JS context. */
(function (global) {
  "use strict";
  const jobs = new Map();
  const cancelledConversions = new Map();
  let owner = null;
  let closed = false;

  function app() {
    const value = global.wps || global.Application;
    if (!value) throw new Error("WPS 接口尚未就绪。");
    return value;
  }

  function fs() {
    if (!app().FileSystem) throw new Error("当前 WPS 未提供本地文件接口。");
    return app().FileSystem;
  }

  function randomId() {
    if (!global.crypto || !global.crypto.getRandomValues) throw new Error("当前 WPS 未提供安全随机数接口，请升级 WPS。");
    return Array.from(global.crypto.getRandomValues(new Uint8Array(16)), value => value.toString(16).padStart(2, "0")).join("");
  }

  function normalized(path) { return String(path || "").replace(/\\/g, "/").toLowerCase(); }
  function exists(path) {
    try { return !!fs().Exists(path); } catch (_) {}
    try { return !!fs().existsSync(path); } catch (_) { return false; }
  }
  function remove(path) {
    try { fs().unlinkSync(path); return; } catch (_) {}
    try { fs().Remove(path); } catch (_) {}
  }
  function write(path, value) {
    const text = JSON.stringify(value).replace(/[\u007f-\uffff]/g, character => "\\u" + character.charCodeAt(0).toString(16).padStart(4, "0"));
    if (typeof fs().WriteFile === "function") fs().WriteFile(path, text);
    else fs().writeAsBinaryString(path, text);
  }
  function read(path) {
    const text = String(typeof fs().ReadFile === "function" ? fs().ReadFile(path) : fs().readAsBinaryString(path));
    if (text.length > 32 * 1024 * 1024) throw new Error("CAD 结果过大，请分区导入。");
    return JSON.parse(text.replace(/^\uFEFF/, ""));
  }
  function pause(milliseconds) { return new Promise(resolve => global.setTimeout(resolve, milliseconds)); }
  function truth(value) { return value === true || Number(value) === -1 || Number(value) === 1; }
  function windowHandle(value) {
    try { return Number(value.HWND) || 0; } catch (_) { return 0; }
  }

  function ownerActive(value) {
    if (!value || value.invalid) return false;
    try {
      const current = app().ActivePresentation;
      if (!current || normalized(current.FullName || current.Name) !== value.name) return false;
      if (!truth(value.window.Active) || Number(value.presentation.Windows.Count) < 1) return false;
      if (value.hwnd && windowHandle(app().ActiveWindow) !== value.hwnd) return false;
      return normalized(value.presentation.FullName || value.presentation.Name) === value.name;
    } catch (_) { return false; }
  }

  function documentKey() {
    if (closed) return "";
    if (ownerActive(owner)) return owner.key;
    try {
      const presentation = app().ActivePresentation;
      const windowObject = app().ActiveWindow;
      if (!presentation || !windowObject || !truth(windowObject.Active)) return "";
      owner = { key: randomId(), presentation: presentation, window: windowObject,
        name: normalized(presentation.FullName || presentation.Name), hwnd: windowHandle(windowObject), invalid: false };
      return owner.key;
    } catch (_) { return ""; }
  }

  function requireJob(job) {
    if (!job || job.cancelled || closed) throw new Error("已取消导入。");
    if (!ownerActive(job.owner) || documentKey() !== job.owner.key) throw new Error("文稿已切换或关闭，请重新导入。");
    return job;
  }

  function workerPath() {
    const current = app().CurrentWPSAddIn;
    const roots = [];
    if (current && current.Path) {
      const base = String(current.Path).replace(/[\\/]+$/, "");
      roots.push(base);
      if (current.Name) roots.push(base + "\\" + String(current.Name).replace(/^[\\/]+|[\\/]+$/g, ""));
    }
    if (global.location && /^file:/i.test(String(global.location.href || ""))) {
      roots.push(decodeURIComponent(String(global.location.href).replace(/^file:\/\/\//i, "").split(/[?#]/)[0]).replace(/\/[^/]*$/, "").replace(/\//g, "\\"));
    }
    for (const root of roots) {
      const path = root + "\\cad\\WpsCadBridge.exe";
      if (exists(path)) return path;
    }
    throw new Error("CAD 转换组件缺失，请安装完整的 WPS 插件更新。");
  }

  function launch(job, phase, options) {
    const request = job.prefix + "." + phase + ".json";
    const status = job.prefix + "." + phase + ".status.json";
    remove(status);
    write(request, Object.assign({ id: job.id, phase: phase }, options || {}));
    const executable = workerPath();
    const checksumsPath = executable.replace(/[\\/]cad[\\/]WpsCadBridge\.exe$/i, "\\checksums.json");
    const workerHash = String(read(checksumsPath)["cad/WpsCadBridge.exe"] || "").toLowerCase();
    if (!/^[a-f0-9]{64}$/.test(workerHash)) throw new Error("CAD 组件的安装校验信息缺失，请重新安装完整插件。");
    write(job.prefix + "." + phase + ".ready", { exe: executable, request: request, workerHash: workerHash });
    const heartbeatPath = job.prefix.replace(/[\\/][^\\/]+$/, "") + "\\DongsidayeWpsCad." + workerHash + ".server.json";
    let running = false;
    try {
      const heartbeat = read(heartbeatPath);
      const age = Date.now() - Number(heartbeat.heartbeatAt);
      running = heartbeat.protocol === 1 && heartbeat.workerHash === workerHash && normalized(heartbeat.exe) === normalized(executable) && age >= -2000 && age < 8000;
    } catch (_) {}
    // An installed session reuses the local worker. WPS can ask for confirmation
    // when starting it again after its idle timeout or a Windows restart.
    if (!running && phase !== "cleanup") {
      const shell = app().OAAssist;
      if (!shell || typeof shell.ShellExecute !== "function") throw new Error("当前 WPS 无法启动 CAD 转换组件。");
      shell.ShellExecute(executable, "--serve");
    }
    return status;
  }

  async function runPhase(job, phase, options) {
    const path = launch(job, phase, options);
    const started = Date.now();
    let seen = false;
    for (;;) {
      await pause(120);
      let status = null;
      try { if (exists(path)) { status = read(path); seen = true; } } catch (_) {}
      if (status && status.state === "done") {
        if (!status.ok) {
          const error = new Error(status.error || (status.cancelled ? "已取消导入。" : "CAD 转换失败。"));
          error.cancelled = !!status.cancelled;
          throw error;
        }
        if (phase === "cleanup") return status;
        const names = { convert: "converted.json", prepare: "plan.json", finalize: "finalized.json" };
        const output = job.folder + "\\" + names[phase];
        if (normalized(status.output) !== normalized(output)) throw new Error("CAD 结果路径无效。");
        return read(output);
      }
      if (status && status.message) job.status = String(status.message);
      if (phase !== "cleanup" && job.cancelledAt && Date.now() - job.cancelledAt > 15000) {
        const error = new Error("已取消导入。");
        error.cancelled = true;
        throw error;
      }
      if ((!seen && Date.now() - started > 15000) || Date.now() - started > 210000) {
        signalCancel(job);
        throw new Error(seen ? "CAD 转换超时，请分区导入。" : "CAD 转换组件没有启动，请重新安装插件后重试。");
      }
    }
  }

  function signalCancel(job) {
    if (!job || job.finished) return;
    job.cancelled = true;
    if (!job.cancelledAt) job.cancelledAt = Date.now();
    try { write(job.prefix + ".cancel.flag", { cancel: true }); } catch (_) {}
  }

  async function cleanup(job) {
    if (!job || job.cleaning) return;
    job.cleaning = true;
    if (job.cancelled && job.conversionDone) {
      cancelledConversions.set(job.id, { state: "done", ok: false, cancelled: true });
      while (cancelledConversions.size > 8) cancelledConversions.delete(cancelledConversions.keys().next().value);
    }
    try {
      await runPhase(job, "cleanup", {});
      remove(job.prefix + ".cleanup.status.json");
    } catch (_) {}
    jobs.delete(job.id);
  }

  function begin(optionsJson) {
    const options = JSON.parse(optionsJson);
    if (!options.documentKey || options.documentKey !== documentKey()) throw new Error("请在当前文稿中重新选择 DWG。");
    if (!/\.dwg$/i.test(String(options.path || ""))) throw new Error("请选择 DWG 文件。");
    if (Array.from(jobs.values()).some(job => job.inserting || !job.conversionDone)) throw new Error("上一项 CAD 导入正在结束，请稍后再试。");
    for (const previous of jobs.values()) { signalCancel(previous); cleanup(previous); }
    const id = randomId();
    let base = "";
    try { base = String(fs().tmpdir() || ""); } catch (_) {}
    if (!base) base = String(app().Env.GetTempPath() || "");
    if (!base || /["\r\n]/.test(base)) throw new Error("WPS 临时目录无效。");
    const prefix = base.replace(/[\\/]+$/, "") + "\\DongsidayeWpsCad_" + id;
    const job = { id: id, prefix: prefix, folder: prefix, owner: owner, cancelled: false, conversionDone: false,
      status: "正在转换模型空间…", done: 0, total: 0, inserting: false, inserted: false };
    jobs.set(id, job);
    runPhase(job, "convert", { input: String(options.path), ocr: options.ocr === true }).then(result => {
      job.result = Object.assign({ state: "done", ok: true }, result);
    }).catch(error => {
      job.result = { state: "done", ok: false, cancelled: job.cancelled || !!error.cancelled, error: String(error.message || error) };
    }).finally(() => {
      job.conversionDone = true;
      if (job.cancelled) cleanup(job);
    });
    return id;
  }

  function getJob(id) {
    const job = jobs.get(String(id));
    if (!job) throw new Error("CAD 导入已结束，请重新选择文件。");
    return job;
  }

  function asShape(value) {
    if (value && value.Type !== undefined) return value;
    if (value && typeof value.Item === "function" && Number(value.Count) === 1) return value.Item(1);
    return value;
  }

  function samePageSize(job, width, height) {
    return Math.abs(Number(job.owner.presentation.PageSetup.SlideWidth) - width) < 0.01 &&
      Math.abs(Number(job.owner.presentation.PageSetup.SlideHeight) - height) < 0.01;
  }

  function rollbackJob(job) {
    let clean = true;
    if (job.group) {
      try { job.group.Delete(); job.group = null; job.created.length = 0; }
      catch (_) { clean = false; }
    } else {
      for (const shape of job.created.slice().reverse()) {
        try { shape.Delete(); job.created.splice(job.created.indexOf(shape), 1); }
        catch (_) { clean = false; }
      }
    }
    return clean;
  }

  async function insertJob(job, options) {
    const created = [];
    job.created = created;
    const fallback = new Set();
    let group = null;
    let nativeCount = 0;
    let ocrCount = 0;
    try {
      requireJob(job);
      const view = Number(job.owner.window.ViewType);
      if (view !== 1 && view !== 9) throw new Error("请切换到普通幻灯片页面后插入。");
      const slide = job.owner.window.View.Slide;
      const slideId = Number(slide.SlideID);
      if (!slideId) throw new Error("请先选择要插入的幻灯片。");
      const pageWidth = Number(job.owner.presentation.PageSetup.SlideWidth);
      const pageHeight = Number(job.owner.presentation.PageSetup.SlideHeight);
      const plan = await runPhase(job, "prepare", { font: options.font || "Microsoft YaHei", pageWidth: pageWidth,
        pageHeight: pageHeight, ocrSelections: options.ocrSelections || [] });
      requireJob(job);
      if (!samePageSize(job, pageWidth, pageHeight)) throw new Error("幻灯片尺寸已变化，请重新导入。");
      if (!Array.isArray(plan.texts) || plan.texts.length > 20500) throw new Error("CAD 文字数量无效。");
      job.total = plan.texts.length;
      const sourceGroups = new Map();
      for (const entry of plan.texts) {
        const key = String(entry.item.SourceId || "");
        if (!key) throw new Error("CAD 文字身份缺失。");
        if (!sourceGroups.has(key)) sourceGroups.set(key, []);
        sourceGroups.get(key).push(entry);
      }
      for (const entries of sourceGroups.values()) {
        const current = [];
        let failed = false;
        for (const entry of entries) {
          requireJob(job);
          if (Number(slide.SlideID) !== slideId) throw new Error("目标幻灯片已关闭。");
          if (!failed) {
            try {
              const text = global.WpsCadText.add(slide, entry, options.font || "Microsoft YaHei");
              if (text) { created.push(text); current.push(text); text.Name = "DWG_Text_" + randomId(); }
            } catch (error) {
              if (error.createdShape) { created.push(error.createdShape); current.push(error.createdShape); }
              if (entry.ocr) throw error;
              failed = true;
              fallback.add(String(entry.item.SourceId));
              for (const text of current) { text.Delete(); created.splice(created.indexOf(text), 1); }
              current.length = 0;
            }
          }
          job.done++;
          if (job.done % 8 === 0) await pause(8);
        }
        if (entries[0].ocr) ocrCount += current.length;
        else nativeCount += current.length;
      }
      requireJob(job);
      const geometry = await runPhase(job, "finalize", { fallbackIds: Array.from(fallback) });
      requireJob(job);
      if (!samePageSize(job, pageWidth, pageHeight)) throw new Error("幻灯片尺寸已变化，请重新导入。");
      if (geometry.hasGeometry) {
        if (normalized(geometry.path) !== normalized(job.folder + "\\insertion.svg")) throw new Error("CAD 图形路径无效。");
        const picture = asShape(slide.Shapes.AddPicture(geometry.path, 0, -1, plan.left, plan.top, plan.width, plan.height));
        created.push(picture);
        picture.Name = "DWG_Geometry_" + randomId();
        for (let n = 0; n < created.length - 1; n++) {
          picture.ZOrder(3);
          if (n % 16 === 15) { await pause(8); requireJob(job); }
        }
      }
      requireJob(job);
      if (!created.length) throw new Error("图纸没有可插入的可见内容。");
      group = created.length === 1 ? created[0] : asShape(slide.Shapes.Range(created.map(shape => String(shape.Name))).Group());
      job.group = group;
      job.inserted = true;
      try { group.Name = "DWG_Import"; group.AlternativeText = "DWG 模型空间 · 可编辑文字 " + (nativeCount + ocrCount); group.Select(); } catch (_) {}
      job.insertionResult = { state: "done", ok: true, textCount: nativeCount, ocrCount: ocrCount, fallbackCount: fallback.size };
      try { if (global.WpsPictureReplace.clearDeckImageCache) global.WpsPictureReplace.clearDeckImageCache(); } catch (_) {}
    } catch (error) {
      const clean = rollbackJob(job);
      if (!clean) job.inserted = true;
      job.insertionResult = { state: "done", ok: false, cancelled: job.cancelled || !!error.cancelled,
        error: clean ? String(error.message || error) : "导入中断且部分对象未能清理，请检查当前页面。" };
    } finally {
      job.inserting = false;
      job.finished = true;
      // Keep the small result available for the panel while discarding native job files.
      cleanup(job);
    }
  }

  const completed = new Map();
  const api = {
    documentKey: documentKey,
    chooseFile: async function () {
      const dialog = app().FileDialog(3);
      dialog.Title = "选择要导入的 DWG 文件";
      dialog.AllowMultiSelect = false;
      try { dialog.Filters.Clear(); dialog.Filters.Add("CAD 图纸", "*.dwg"); } catch (_) {}
      if (Number(dialog.Show()) !== -1 || Number(dialog.SelectedItems.Count) < 1) return { ok: false, cancelled: true };
      return { ok: true, path: String(dialog.SelectedItems.Item(1)) };
    },
    BeginDwgImport: begin,
    PollDwgImport: function (id) {
      if (cancelledConversions.has(String(id))) return JSON.stringify(cancelledConversions.get(String(id)));
      const job = getJob(id);
      return JSON.stringify(job.conversionDone ? job.result : { state: "pending", status: job.status });
    },
    CancelDwgImport: function (id) {
      const job = jobs.get(String(id));
      if (!job) return;
      signalCancel(job);
      if (job.conversionDone && !job.inserting) cleanup(job);
    },
    BeginDwgInsertion: function (id, optionsJson) {
      const job = requireJob(getJob(id));
      if (!job.conversionDone || !job.result.ok || job.inserting || job.inserted || job.finished) throw new Error("CAD 导入尚未就绪或已结束。");
      const options = JSON.parse(optionsJson || "{}");
      if (!options || typeof options !== "object" || Array.isArray(options)) throw new Error("CAD 插入选项无效。");
      job.inserting = true;
      insertJob(job, options).then(() => {
        completed.set(job.id, job.insertionResult);
        while (completed.size > 8) completed.delete(completed.keys().next().value);
      });
      return JSON.stringify({ ok: true, pending: true });
    },
    PollDwgInsertion: function (id) {
      if (completed.has(String(id))) return JSON.stringify(completed.get(String(id)));
      const job = getJob(id);
      return JSON.stringify(job.insertionResult || { state: "pending", done: job.done, total: job.total });
    }
  };
  global.WpsCadBridge = api;

  function notifyChanged() {
    for (const job of jobs.values()) if (!ownerActive(job.owner)) signalCancel(job);
    if (typeof global.dispatchEvent === "function") global.dispatchEvent(new Event("wps-cad-document-changed"));
  }
  try {
    const events = app().ApiEvent;
    events.AddApiEventListener("PresentationClose", function (presentation) {
      let name = "";
      try { name = normalized(presentation.FullName || presentation.Name); } catch (_) {}
      if (owner && (!name || name === owner.name)) { owner.invalid = true; notifyChanged(); }
    });
    events.AddApiEventListener("WindowActivate", function () {
      if (owner && !ownerActive(owner)) notifyChanged();
    });
  } catch (_) {}
  if (typeof global.addEventListener === "function") {
    global.addEventListener("beforeunload", function () {
      closed = true;
      for (const job of jobs.values()) {
        signalCancel(job);
        if (job.inserting && !job.inserted) rollbackJob(job);
        if (job.conversionDone) cleanup(job);
      }
    });
    global.setInterval(function () {
      if (owner && !ownerActive(owner)) { owner.invalid = true; owner = null; notifyChanged(); }
    }, 350);
  }
}(typeof window !== "undefined" ? window : globalThis));
