"""Native PowerPoint source identity and clipboard-free inventory previews."""

def once(text, old, new, label):
    if text.count(old) != 1:
        raise ValueError(label + ': source drift')
    return text.replace(old, new, 1)


def prepare_sources(text):
    helpers = r'''  let panelPictureSources = {};
  let panelPictureSourceDocument = "";
  let panelPictureSourceRevision = "";

  async function panelPreparePictureSources() {
    if (!global.__pptLoadPictureSources) return;
    const document = global.__pptDocumentKey();
    let entries = {};
    try { entries = await global.__pptLoadPictureSources(); }
    catch (error) {
      if (document !== global.__pptDocumentKey()) throw error;
      // Native previews still work; unverified identities remain separate.
      entries = {};
    }
    if (document !== global.__pptDocumentKey()) throw new Error("文稿已切换，请重新读取图片。");
    const keys = Object.keys(entries).sort();
    const revision = keys.length ? JSON.stringify(keys.map(function (key) {
      return [key, entries[key].fingerprint, entries[key].crop, entries[key].display];
    })) : "unverified:" + Date.now();
    if (revision !== panelPictureSourceRevision || document !== panelPictureSourceDocument) invalidatePanelInventoryCache();
    panelPictureSources = entries;
    panelPictureSourceDocument = document;
    panelPictureSourceRevision = revision;
  }

  function panelPictureSourceKey(item) {
    try {
      const id = Number(item.shape.Id);
      if (item.kind === "slide" || !item.kind) {
        const slide = activePresentation().Slides.Item(Number(item.slideIndex));
        return "slide:" + Number(slide.SlideID) + ":" + id;
      }
      // The inherited scanner only handles one master. Do not guess the
      // owning design when a document has more than one.
      if (Number(activePresentation().Designs.Count) !== 1) return "";
      if (item.kind === "master") return "master:1:" + id;
      if (item.kind === "layout") return "layout:1:" + Number(item.layoutIndex) + ":" + id;
    } catch (_) {}
    return "";
  }

  function panelNativeVisibleThumbnail(binary, source) {
    return new Promise(function (resolve, reject) {
      const url = global.URL.createObjectURL(new global.Blob([binaryToBytes(binary)], {type:"image/png"}));
      const img = new global.Image();
      img.onload = function () {
        try {
          const canvas = global.document.createElement("canvas");
          canvas.width = THUMB_PX; canvas.height = THUMB_PX;
          if (source) {
            const crop=source.crop;
            canvas.getContext("2d").drawImage(img,crop.l*img.width,crop.t*img.height,(crop.r-crop.l)*img.width,(crop.b-crop.t)*img.height,0,0,THUMB_PX,THUMB_PX);
            resolve(panelDisplayThumbnail(canvas,source.display));
          } else {
            const fit = THUMB_PX / Math.max(img.width, img.height);
            const width = img.width * fit, height = img.height * fit;
            canvas.getContext("2d").drawImage(img, (THUMB_PX-width)/2, (THUMB_PX-height)/2, width, height);
            resolve(canvas.toDataURL("image/png"));
          }
        } catch (error) { reject(error); }
        finally { global.URL.revokeObjectURL(url); }
      };
      img.onerror = function () { global.URL.revokeObjectURL(url); reject(new Error("无法解码图片预览。")); };
      img.src = url;
    });
  }

  async function renderNativePictureThumbnails(items, onChunk) {
    const result = new Array(items.length).fill(null);
    for (let i=0; i<items.length; i+=1) {
      if (global.__pptDocumentKey && panelPictureSourceDocument !== global.__pptDocumentKey()) throw new Error("文稿已切换，已停止图片读取。");
      const item=items[i], shape=item.shape, path=tempPath("ppt-native-panel");
      try {
        const source=panelPictureSources[panelPictureSourceKey(item)];
        let dataUrl="", fingerprint="";
        if (source && source.display) {
          try {
            // Preview pixels and source identity come from the same immutable
            // snapshot, even if the user continues editing the live slide.
            const binary=global.__pptReadSourcePreview ? global.__pptReadSourcePreview(source.path) : fileSystem().readAsBinaryString(source.path);
            dataUrl=await panelNativeVisibleThumbnail(binary,source);
            fingerprint=source.fingerprint;
          } catch (_) {}
        }
        if (!dataUrl) {
          // Unsupported media gets a native preview of this exact object, but
          // no identity authorizing an automatic multi-object replacement.
          shape.Export(path, 2);
          if (!fileExists(path)) throw new Error("图片预览导出失败。");
          dataUrl=await panelNativeVisibleThumbnail(fileSystem().readAsBinaryString(path));
        }
        const crop=source && source.crop ? normalizeCropWindow(Object.assign({known:true},source.crop)) : cropWindowForShape(shape,0,0);
        result[i]={fp:fingerprint,dataUrl:dataUrl,dHash:"",aHash:"",cropWindow:crop,
          w:THUMB_PX,h:THUMB_PX,nw:num(shape.Width),nh:num(shape.Height),aspect:Math.round(1000*num(shape.Width)/Math.max(0.01,num(shape.Height)))};
      } catch (_) { result[i]=emptyPanelInfo(); }
      finally { removeFile(path); }
      if (onChunk) await onChunk(i+1,items.length,result);
      await yieldUI();
    }
    return result;
  }

  async function panelValidateInstanceSources(instances) {
    if (!global.__pptLoadPictureSources) return;
    await panelPreparePictureSources();
    for (let i=0;i<instances.length;i+=1) {
      const instance=instances[i], key=panelPictureSourceKey(instance);
      if (key!==instance.sourceKey || Number(instance.shape.Id)!==Number(instance.shapeIdentity) ||
          (instance.sourceIdentity && instance.sourceIdentity!==((panelPictureSources[key]||{}).fingerprint||""))) {
        throw new Error("图片已变化，请刷新清单后重新选择，已停止替换。");
      }
    }
  }

  function panelSourceIdentityForShape(shape) {
    try {
      const key="slide:"+Number(shape.Parent.SlideID)+":"+Number(shape.Id);
      return (panelPictureSources[key]||{}).fingerprint || "";
    } catch (_) { return ""; }
  }

'''
    text=once(text,'  async function collectDeckImagesCached(onProgress, onPartial, options) {',helpers+'  async function collectDeckImagesCached(onProgress, onPartial, options) {\n    await panelPreparePictureSources();','source snapshot before cache')
    text=once(text,'  async function collectDeckImages(onProgress, onPartial) {\n    const presentation', '  async function collectDeckImages(onProgress, onPartial) {\n    if (global.__pptLoadPictureSources && panelPictureSourceDocument !== global.__pptDocumentKey()) await panelPreparePictureSources();\n    const presentation','direct inventory source preparation')
    text=once(text,'  const PANEL_CACHE_VERSION = 3;','  const PANEL_CACHE_VERSION = 4;','discard poisoned picture caches')
    text=once(text,'        if (panelInventoryMemoryCache.presentation !== presentation) {','        if (panelInventoryMemoryCache.snapshot) {','validate memory instances too')
    text=once(text,'''      const group = Object.assign({}, cachedGroup);
      group.instances = instances;''','''      const identity=String(cachedGroup.contentFp || "");
      if (instances.length>1 && !/^sha256:[a-f0-9]{64}$/.test(identity)) return null;
      if (identity && instances.some(function (entry) { return entry.sourceIdentity!==identity; })) return null;
      for (let a=0;a<instances.length;a+=1) for(let b=a+1;b<instances.length;b+=1) {
        if (!cropWindowCompatible(instances[a].cropWindow,instances[b].cropWindow)) return null;
      }
      const group = Object.assign({}, cachedGroup);
      group.instances = instances;''','validate cached group membership')
    text=once(text,'      "aspect", "userAlt", "thumb", "thumbW", "thumbH", "cropWindow"','      "aspect", "userAlt", "thumb", "thumbW", "thumbH", "cropWindow", "sourceKey", "sourceIdentity", "shapeIdentity", "rotation", "flipH", "flipV"','persist instance identity')
    text=once(text,'      uid: panelInstanceUid(item),','''      uid: panelInstanceUid(item),
      sourceKey: panelPictureSourceKey(item),
      sourceIdentity: (panelPictureSources[panelPictureSourceKey(item)] || {}).fingerprint || "",
      shapeIdentity: Number(item.shape.Id),
      rotation: num(item.shape.Rotation), flipH: isTrue(item.shape.HorizontalFlip), flipV: isTrue(item.shape.VerticalFlip),''','snapshot actual shape binding')
    text=once(text,'    if (!shape || !cached || !isPicture(shape)) return false;','''    if (!shape || !cached || !isPicture(shape)) return false;
    if (Number(shape.Id) !== Number(cached.shapeIdentity)) return false;
    if (panelPictureSourceKey(Object.assign({},cached,{shape:shape})) !== cached.sourceKey) return false;
    if (!panelInventoryNear(shape.Rotation,cached.rotation) || isTrue(shape.HorizontalFlip)!==cached.flipH || isTrue(shape.VerticalFlip)!==cached.flipV) return false;
    if (cached.sourceIdentity !== ((panelPictureSources[cached.sourceKey] || {}).fingerprint || "")) return false;''','validate source and shape identity')
    text=once(text,'''        const scratch = getSharedScratch();
        await yieldUI();
        let lastBatchDone = 0;
        batchResults = await renderThumbnailsBatch(renderQueue.map(function (i) { return pending[i]; }), scratch, async function (done, count, liveResults) {''','''        await yieldUI();
        let lastBatchDone = 0;
        batchResults = await renderNativePictureThumbnails(renderQueue.map(function (i) { return pending[i]; }), async function (done, count, liveResults) {''','native exact-object previews')
    start=text.index('          if (!info || !info.fp) {\n            // Graceful fallback:')
    end=text.index('          const thumbKey',start)
    text=text[:start]+'''          if (!info) { info = emptyPanelInfo(); fallbackCount += 1; }
'''+text[end:]
    text=once(text,'  async function replaceInstances(instances, imagePath, docKey, onProgress, cancelCheck) {','  async function replaceInstances(instances, imagePath, docKey, onProgress, cancelCheck) {\n    await panelValidateInstanceSources(instances);','validate sources before batch replacement')
    text=once(text,'  async function updateLinkedInstances(instances, docKey, onProgress, cancelCheck) {','  async function updateLinkedInstances(instances, docKey, onProgress, cancelCheck) {\n    await panelValidateInstanceSources(instances);','validate sources before linked update')
    start=text.index('      const referenceSignatures = new Set(await signaturesForShape(')
    end=text.index('      const candidates = [];',start)
    text=text[:start]+'''      await panelPreparePictureSources();
      const referenceIdentity=panelSourceIdentityForShape(referenceShape);
      if (!/^sha256:[a-f0-9]{64}$/.test(referenceIdentity)) throw new Error("无法确认原图片身份，已停止批量替换；请使用单图替换。");
'''+text[end:]
    start=text.index('        const path = tempPath("candidate");',start)
    end=text.index('\n      }',start)
    text=text[:start]+'''        if (panelSourceIdentityForShape(candidates[i]) === referenceIdentity) matches.push(candidates[i]);'''+text[end:]
    import re
    for name in ('replaceAllMatching','replaceInstances','updateLinkedInstances'):
        begin=text.index('  async function '+name+'(')
        following=re.search(r'^  (?:async )?function ',text[begin+3:],re.M)
        end=begin+3+following.start()
        block=text[begin:end]
        brace=block.index('{')+1
        block=block[:brace]+'''
    const operationOwner=global.__pptDocumentKey ? global.__pptDocumentKey() : "";
    function verifyOwner() {
      if (operationOwner && global.__pptDocumentKey()!==operationOwner) throw new Error("文稿已切换，已停止批量操作。");
    }
'''+block[brace:]
        block=block.replace('await yieldUI();','await yieldUI(); verifyOwner();')
        block=block.replace('await panelPreparePictureSources();','await panelPreparePictureSources(); verifyOwner();')
        block=block.replace('await panelValidateInstanceSources(instances);','await panelValidateInstanceSources(instances); verifyOwner();')
        block=block.replace('const info = await contentFingerprintAndThumb(newShape, scratch);','const info = await contentFingerprintAndThumb(newShape, scratch); verifyOwner();')
        block=block.replace('for (let i = 0; i < targets.length; i += 1) {','for (let i = 0; i < targets.length; i += 1) {\n        verifyOwner();')
        text=text[:begin]+block+text[end:]
    return text
