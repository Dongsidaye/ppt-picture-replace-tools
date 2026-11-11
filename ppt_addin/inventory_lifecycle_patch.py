"""Fail closed when image inventory work outlives its document/revision."""
from source_patch import once


def prepare_inventory_lifecycle(text: str) -> str:
    text = once(text, 'liveResult = panelInventoryHydrate(panelInventoryMemoryCache.snapshot, presentation);', 'liveResult = panelInventoryHydrate(panelInventoryMemoryCache.snapshot, presentation, panelInventoryMemoryCache.result);', 'retain proven template references for memory cache')
    text = once(text, '  function panelInventoryHydrate(snapshot, presentation) {', '  function panelInventoryHydrate(snapshot, presentation, liveResult) {', 'optional live cache references')
    text = once(text, '''        const cached = cachedInstances[i];
        const shape = panelInventoryResolvedShape(presentation, cached);''', '''        const cached = Object.assign({},cachedInstances[i]);
        if (cached.kind!=="slide" && liveResult) {
          const live=[];
          (liveResult.groups||[]).forEach(function(group){(group.instances||[]).forEach(function(entry){
            if(entry.uid===cached.uid && entry.sourceKey===cached.sourceKey && Number(entry.shapeIdentity)===Number(cached.shapeIdentity)) live.push(entry);
          });});
          if(live.length!==1) return null;
          cached.containerReference=live[0].containerReference;cached.shape=live[0].shape;
        }
        const shape = panelInventoryResolvedShape(presentation, cached);''', 'template memory references are revalidated not serialized')
    text = once(text, '  async function collectDeckImagesCached(onProgress, onPartial, options) {\n    await panelPreparePictureSources();', '  async function collectDeckImagesCached(onProgress, onPartial, options) {', 'join scan before creating another snapshot')
    text = once(text, '''    if (!force) {
      if (panelInventoryMemoryCache''', '''    await panelPreparePictureSources();
    // Another caller can have started the scan while both awaited the same
    // native snapshot. Join it instead of starting a second export loop.
    if (panelInventoryScan) return collectDeckImagesCached(onProgress, onPartial, options);
    if (!force) {
      if (panelInventoryMemoryCache''', 'coalesce concurrent preparation')
    structure = '''  function panelInventoryStructure(presentation) {
    try {
      const parts=[];
      function add(kind, container) {
        const ids=[],shapes=container.Shapes;
        for(let i=1;i<=Number(shapes.Count);i+=1) {const shape=shapes.Item(i);ids.push(Number(shape.Id)+":"+Number(shape.Type));}
        parts.push(kind+"["+ids.sort().join(",")+"]");
      }
      for(let s=1;s<=Number(presentation.Slides.Count);s+=1) {
        const slide=presentation.Slides.Item(s);add("s:"+Number(slide.SlideID),slide);
      }
      const designs=presentation.Designs;
      for(let d=1;d<=Number(designs.Count);d+=1) {
        const master=designs.Item(d).SlideMaster;add("m:"+d,master);
        for(let l=1;l<=Number(master.CustomLayouts.Count);l+=1) add("l:"+d+":"+l,master.CustomLayouts.Item(l));
      }
      return parts.sort().join("|");
    } catch (_) {return "";}
  }

'''
    text = once(text, '  function panelInventorySnapshot(result) {', structure+'  function panelInventorySnapshot(result) {', 'inventory coverage signature')
    text = once(text, '      version: PANEL_CACHE_VERSION,\n      createdAt:', '      version: PANEL_CACHE_VERSION,\n      structure: panelInventoryStructure(activePresentation()),\n      createdAt:', 'cache complete shape coverage')
    text = once(text, '    if (String(snapshot.docKey || "") !== documentKey(presentation)) return null;', '    if (String(snapshot.docKey || "") !== documentKey(presentation)) return null;\n    if (!snapshot.structure || snapshot.structure !== panelInventoryStructure(presentation)) return null;', 'cache rejects added unknown objects')
    text = once(text, '  function cropWindowCompatible(a, b) {\n    return cropWindowCoverage(a, b) >= CROP_GROUP_MIN_COVERAGE;\n  }', '''  function cropWindowCompatible(a, b) {
    if (!a || !b || !a.known || !b.known) return false;
    return ["l","t","r","b"].every(function (key) {
      return Number.isFinite(Number(a[key])) && Number.isFinite(Number(b[key])) && Math.abs(Number(a[key])-Number(b[key])) <= 0.00001;
    });
  }''', 'exact visible crop identity')
    guard = '''  function panelInventoryGuard(owner, epoch) {
    if (epoch !== panelInventoryEpoch || (global.__pptDocumentKey && owner !== global.__pptDocumentKey())) {
      const error = new Error("文稿或图片已变化，已丢弃过期清单，请重新刷新。");
      error.pptInventoryStale = true;
      throw error;
    }
  }

'''
    text = once(text, '  async function panelInventoryRunScan(presentation, docKey, onProgress, onPartial) {\n    const epoch = panelInventoryEpoch;', guard + '''  async function panelInventoryRunScan(presentation, docKey, onProgress, onPartial) {
    const epoch = panelInventoryEpoch;
    const owner = global.__pptDocumentKey ? global.__pptDocumentKey() : "";
    const structure = panelInventoryStructure(presentation);
    const check = function () { panelInventoryGuard(owner, epoch); };
    const progress = onProgress ? function () { check(); return onProgress.apply(null, arguments); } : null;
    const partial = onPartial ? function () { check(); return onPartial.apply(null, arguments); } : null;''', 'scan generation ownership')
    text = once(text, '        const result = await collectDeckImages(onProgress, onPartial);\n        let current = null;', '''        const result = await collectDeckImages(progress, partial);
        check();
        // Re-read the original media after asynchronous preview generation.
        // Edits made while a scan was running must not bless an older view.
        if (global.__pptLoadPictureSources) await panelPreparePictureSources();
        if (structure !== panelInventoryStructure(presentation)) invalidatePanelInventoryCache();
        check();
        let current = null;''', 'scan final revision verification')
    text = once(text, '    panelInventoryScan = { docKey: docKey, promise: promise };', '    panelInventoryScan = { docKey: docKey, promise: promise, epoch: epoch, owner: owner };', 'scan token metadata')
    text = once(text, '''        } catch (_) {
          // A best-effort background preload may fail''', '''        } catch (error) {
          if (error && error.pptInventoryStale) {
            // A caller explicitly invalidated an older preload. Retry once,
            // only for the same live document; never publish that old result.
            const retries=Number(options && options._staleRetries)||0;
            if (documentKey(activePresentation())!==docKey || retries>=1) throw error;
            return collectDeckImagesCached(onProgress,onPartial,Object.assign({},options===true?{force:true}:(options||{}),{_staleRetries:retries+1}));
          }
          // A best-effort background preload may fail''', 'do not retry stale scan on another document')
    text = once(text, '''  async function panelInventoryDeliverCached(result, onProgress, onPartial, source) {
    await yieldUI();''', '''  async function panelInventoryDeliverCached(result, onProgress, onPartial, source) {
    const owner = global.__pptDocumentKey ? global.__pptDocumentKey() : "";
    const epoch = panelInventoryEpoch;
    await yieldUI();
    panelInventoryGuard(owner, epoch);''', 'cached delivery owner')
    text = once(text, '    return panelInventoryResultView(result, true, source);', '    panelInventoryGuard(owner, epoch);\n    return panelInventoryResultView(result, true, source);', 'cached delivery final guard')
    text = once(text, '''  async function renderNativePictureThumbnails(items, onChunk) {
    const result''', '''  async function renderNativePictureThumbnails(items, onChunk) {
    const sources = panelPictureSources;
    const owner = panelPictureSourceDocument;
    const epoch = panelInventoryEpoch;
    const result''', 'immutable preview source context')
    text = once(text, '      if (global.__pptDocumentKey && panelPictureSourceDocument !== global.__pptDocumentKey()) throw new Error("文稿已切换，已停止图片读取。");', '      panelInventoryGuard(owner, epoch);', 'preview starting guard')
    text = once(text, '        const source=panelPictureSources[panelPictureSourceKey(item)];', '        const source=sources[panelPictureSourceKey(item)];', 'preview immutable lookup')
    text = once(text, '      if (onChunk) await onChunk(i+1,items.length,result);', '      panelInventoryGuard(owner, epoch);\n      if (onChunk) await onChunk(i+1,items.length,result);', 'preview publication guard')
    # Session identity is orthogonal to file path and shape counts. A closed
    # document reopened under the same name cannot inherit old COM instances.
    text = once(text, '  function documentKey(presentation) {\n    try {', '''  function documentKey(presentation) {
    if (global.__pptDocumentKey) return String(global.__pptDocumentKey() || "");
    try {''', 'native session document cache identity')
    return text


def prepare_inventory_pane(text: str) -> str:
    return once(text, '''          state.scanning = false;
          hideProgress();
          setStatus("刷新失败："''', '''          state.scanning = false;
          state.groups = [];
          state.selected = {};
          state.docKey = "";
          render();
          hideProgress();
          setStatus("刷新失败："''', 'clear invalid partial rows on failed refresh')
