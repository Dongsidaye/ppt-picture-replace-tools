"""Bind PowerPoint inventory actions to document/container/shape identities."""
import re


def once(text, old, new, label):
    if text.count(old) != 1:
        raise ValueError(label + ': source drift')
    return text.replace(old, new, 1)


def function(text, name, body):
    found = list(re.finditer(r'^  (?:async )?function ' + re.escape(name) + r'\(', text, re.M))
    if len(found) != 1:
        raise ValueError(name + ': expected one function')
    start = found[0].start()
    following = re.search(r'^  (?:async )?function ', text[found[0].end():], re.M)
    if not following:
        raise ValueError(name + ': missing function boundary')
    end = found[0].end() + following.start()
    return text[:start] + body.strip('\n') + '\n\n' + text[end:]


HELPERS = r'''  function panelSameObject(left, right) {
    if(!left || !right) return false;
    try {return global.__pptSameObject ? global.__pptSameObject(left,right)===true : left===right;}catch(_){return false;}
  }

  function panelFindShapeById(container, identity) {
    if (!container || !(identity > 0)) return null;
    let match=null, index=0;
    try {
      const shapes=container.Shapes, count=Number(shapes.Count)||0;
      for(let i=1;i<=count;i+=1) {
        const shape=shapes.Item(i);
        if(Number(shape.Id)!==identity) continue;
        if(match) return null;
        if(!isPicture(shape)) return null;
        match=shape; index=i;
      }
    } catch (_) { return null; }
    return match ? {shape:match,shapeIndex:index,container:container} : null;
  }

  function panelFindSlideById(presentation, identity) {
    if(!presentation || !(identity>0)) return null;
    try {
      if(hasMethod(presentation.Slides,"FindBySlideID")) {
        const slide=presentation.Slides.FindBySlideID(identity);
        if(slide && Number(slide.SlideID)===identity) return slide;
      }
      const count=Number(presentation.Slides.Count)||0;
      for(let i=1;i<=count;i+=1) {
        const slide=presentation.Slides.Item(i);
        if(Number(slide.SlideID)===identity) return slide;
      }
    } catch (_) {}
    return null;
  }

  function panelResolveInstance(presentation, instance) {
    if(!instance || !instance.ownerDocumentKey || instance.ownerDocumentKey!==pptOwnerDocumentKey()) return null;
    const shapeIdentity=Number(instance.shapeIdentity), kind=String(instance.kind||"slide");
    if(!(shapeIdentity>0)) return null;
    if(kind==="slide") {
      const source=/^slide:(\d+):(\d+)$/.exec(String(instance.sourceKey||""));
      const slideIdentity=Number(instance.slideIdentity)||(source ? Number(source[1]) : 0);
      if(!source || Number(source[1])!==slideIdentity || Number(source[2])!==shapeIdentity) return null;
      const slide=panelFindSlideById(presentation,slideIdentity);
      const found=panelFindShapeById(slide,shapeIdentity);
      if(!found) return null;
      found.slideIdentity=slideIdentity; found.slideIndex=Number(slide.SlideIndex);
      found.sourceKey="slide:"+slideIdentity+":"+shapeIdentity;
      return found;
    }
    // No persisted ordinal is proof of a master/layout container's identity.
    // Keep these actions available only while the original live container can
    // be verified in this presentation. COM wrappers that cannot prove identity
    // fail closed and require a fresh list/native selection.
    const container=instance.containerReference;
    if(!container || !instance.shape) return null;
    let owned=false,masterIndex=0,layoutIndex=0;
    try {
      if(!panelSameObject(instance.shape.Parent,container)) return null;
      const designs=presentation.Designs;
      for(let d=1;d<=Number(designs.Count);d+=1) {
        const master=designs.Item(d).SlideMaster;
        if(kind==="master" && panelSameObject(master,container)) {owned=true;masterIndex=d;}
        if(kind==="layout") {
          const layouts=master.CustomLayouts;
          for(let l=1;l<=Number(layouts.Count);l+=1) if(panelSameObject(layouts.Item(l),container)) {owned=true;masterIndex=d;layoutIndex=l;}
        }
      }
    } catch (_) { return null; }
    if(!owned) return null;
    const found=panelFindShapeById(container,shapeIdentity);
    if(found) {
      found.slideIndex=0;found.masterIndex=masterIndex;found.layoutIndex=layoutIndex;
      found.sourceKey=kind==="master" ? "master:"+masterIndex+":"+shapeIdentity : "layout:"+masterIndex+":"+layoutIndex+":"+shapeIdentity;
    }
    return found;
  }

  function panelRefreshResolvedInstance(instance, found) {
    instance.shape=found.shape;instance.shapeIndex=found.shapeIndex;
    instance.ownerDocumentKey=pptOwnerDocumentKey();
    if(found.slideIdentity) {
      instance.slideIdentity=found.slideIdentity;instance.slideIndex=found.slideIndex;
      instance.appliedPages=[found.slideIndex];instance.sourceKey=found.sourceKey;
    } else {
      instance.masterIndex=found.masterIndex;instance.layoutIndex=found.layoutIndex;
      instance.containerReference=found.container;instance.sourceKey=found.sourceKey;
    }
    return instance;
  }

  function panelInstanceFromLiveShape(shape) {
    try {
      const slide=slideOf(shape), id=Number(shape.Id), sid=Number(slide.SlideID);
      if(!(id>0 && sid>0)) return null;
      return {ownerDocumentKey:pptOwnerDocumentKey(),kind:"slide",slideIdentity:sid,
        shapeIdentity:id,sourceKey:"slide:"+sid+":"+id,shape:shape};
    } catch (_) { return null; }
  }

  async function locatePictureInstance(instance) {
    const found=panelResolveInstance(activePresentation(),instance);
    if(!found) throw new Error("图片已不存在或所属文稿已变化，请刷新清单。");
    panelRefreshResolvedInstance(instance,found);
    if((instance.kind||"slide")==="slide") return locateSlideShape(found.slideIndex,found.shapeIndex,found.shape,instance);
    if(instance.kind==="master") return locateMasterShape(found.shapeIndex,found.shape,instance);
    if(instance.kind==="layout") return locateLayoutShape(instance.layoutIndex,found.shapeIndex,found.shape,instance);
    return false;
  }

  function panelInstanceFromTemplateReference(kind, shape) {
    try {return {ownerDocumentKey:pptOwnerDocumentKey(),kind:kind,shape:shape,
      shapeIdentity:Number(shape.Id),containerReference:shape.Parent};}catch(_){return null;}
  }

  async function locateTemplatePictureInstance(instance) {
    let found=panelResolveInstance(activePresentation(),instance);
    if(!found) return false;
    const owner=instance.ownerDocumentKey,windowObject=application().ActiveWindow;
    if(!windowObject || !await requestWindowViewType(windowObject,2) || owner!==pptOwnerDocumentKey()) return false;
    found=panelResolveInstance(activePresentation(),instance);
    if(!found || layerGuardShapeLocked(found.shape)) return false;
    panelRefreshResolvedInstance(instance,found);
    try {if(hasMethod(found.container,"Select")) {activateWindow(windowObject);found.container.Select(MsoTrue);}}catch(_){return false;}
    await yieldUI(70);
    if(owner!==pptOwnerDocumentKey()) return false;
    found=panelResolveInstance(activePresentation(),instance);
    if(!found) return false;
    panelRefreshResolvedInstance(instance,found);
    return selectCanvasShapeAsync(windowObject,found.shape,found.shapeIndex);
  }

  function renamePictureInstances(instances, name) {
    const targets=resolveTargets(activePresentation(),instances);
    let count=0;
    for(let i=0;i<targets.length;i+=1) if(renameShape(targets[i].shape,name)) count+=1;
    return count;
  }

'''


def prepare_targets(text):
    text=once(text,'  function panelInventoryResolvedShape(presentation, instance) {',HELPERS+'  function panelInventoryResolvedShape(presentation, instance) {','stable instance helpers')
    text=function(text,'panelInventoryResolvedShape',r'''  function panelInventoryResolvedShape(presentation, instance) {
    const found=panelResolveInstance(presentation,instance);
    return found ? found.shape : null;
  }''')
    text=function(text,'resolveTargets',r'''  function resolveTargets(presentation, instances) {
    const targets=[], seen=new Set();
    for(let i=0;i<instances.length;i+=1) {
      const instance=instances[i], found=panelResolveInstance(presentation,instance);
      if(!found) throw new Error("图片已不存在或所属位置已变化，请刷新清单后重试。");
      const identity=String(found.sourceKey||instance.sourceKey)+":"+Number(instance.shapeIdentity);
      if(seen.has(identity)) continue;
      seen.add(identity);panelRefreshResolvedInstance(instance,found);
      targets.push({shape:found.shape,inst:instance});
    }
    return targets;
  }''')
    text=function(text,'panelValidateInstanceSources',r'''  async function panelValidateInstanceSources(instances) {
    if(global.__pptLoadPictureSources) await panelPreparePictureSources();
    const targets=resolveTargets(activePresentation(),instances);
    for(let i=0;i<targets.length;i+=1) {
      const instance=targets[i].inst, source=panelPictureSources[instance.sourceKey]||{};
      if(!instance.sourceIdentity || !instance.sourceAppearance) {
        throw new Error("无法验证此图片的原始身份，请在幻灯片正文中选中后使用单图替换。");
      }
      if(instance.sourceIdentity!==(source.fingerprint||"") ||
         instance.sourceAppearance!==(source.appearance||"")) {
        throw new Error("图片内容或裁剪效果已变化，请刷新清单后重新选择，已停止替换。");
      }
    }
  }''')
    text=once(text,'      uid: panelInstanceUid(item),','''      uid: panelInstanceUid(item),
      ownerDocumentKey: pptOwnerDocumentKey(),
      slideIdentity: item.kind==="slide" ? Number(item.slide.SlideID) : 0,
      containerReference: item.kind!=="slide" ? item.shape.Parent : null,''','bind runtime inventory owner')
    text=once(text,'    out.slideIndex = Number(out.slideIndex) || 0;','''    out.ownerDocumentKey=String(instance && instance.ownerDocumentKey || "");
    out.slideIdentity=Number(instance && instance.slideIdentity)||0;
    out.slideIndex = Number(out.slideIndex) || 0;''','persist stable owner fields')
    text=once(text,'        instance.shape = shape;','''        const found=panelResolveInstance(presentation,instance);
        if(!found) return null;
        panelRefreshResolvedInstance(instance,found);''','refresh cached positions')
    text=once(text,'        const slide = activePresentation().Slides.Item(Number(item.slideIndex));\n        return "slide:" + Number(slide.SlideID) + ":" + id;','''        const slide = slideOf(item.shape);
        return "slide:" + Number(slide.SlideID) + ":" + id;''','source key from real parent')
    text=function(text,'selectionMatchesShape',r'''  function selectionMatchesShape(windowObject, shape) {
    try {
      const range=windowObject && windowObject.Selection && windowObject.Selection.ShapeRange;
      if(!range) return null;
      const id=Number(shape.Id);if(!(id>0)) return false;
      for(let i=1;i<=Number(range.Count);i+=1) {
        const selected=range.Item(i);
        if(Number(selected.Id)!==id) continue;
        // IDs are local to a container: when available, verify slide identity.
        let targetSlide=0,selectedSlide=0;
        try {targetSlide=Number(slideOf(shape).SlideID)||0;selectedSlide=Number(slideOf(selected).SlideID)||0;}catch(_){}
        if(targetSlide && selectedSlide) {if(targetSlide!==selectedSlide) continue;}
        else if(!panelSameObject(shape.Parent,selected.Parent)) continue;
        return true;
      }
      return false;
    } catch (_) { return null; }
  }''')
    text=function(text,'selectCanvasShapeAsync',r'''  async function selectCanvasShapeAsync(windowObject, shape, shapeIndex) {
    if(!shape || !hasMethod(shape,"Select") || layerGuardShapeLocked(shape)) return false;
    const owner=pptOwnerDocumentKey();
    try {activateWindow(windowObject);shape.Select(MsoTrue);}catch(_){return false;}
    await yieldUI(55);
    if(!owner || owner!==pptOwnerDocumentKey()) return false;
    // An old index or duplicate name can select a different live picture.
    // Never retry through those ambiguous selectors.
    return selectionMatchesShape(windowObject,shape)===true;
  }''')
    text=function(text,'locateSlideShape',r'''  async function locateSlideShape(slideIndex, shapeIndex, shapeReference, instance) {
    const target=instance || panelInstanceFromLiveShape(shapeReference);
    const found=panelResolveInstance(activePresentation(),target);
    if(!found) return false;
    panelRefreshResolvedInstance(target,found);
    const owner=target.ownerDocumentKey, windowObject=application().ActiveWindow;
    if(!windowObject || !windowObject.View || layerGuardShapeLocked(found.shape)) return false;
    if(!await requestWindowViewType(windowObject,NORMAL_VIEW_TYPE) || owner!==pptOwnerDocumentKey()) return false;
    const current=panelResolveInstance(activePresentation(),target);
    if(!current) return false;
    activateWindow(windowObject);
    try {windowObject.View.GotoSlide(current.slideIndex);}catch(_){return false;}
    const readable=readWindowSlideIndex(windowObject);
    if(readable===null) await yieldUI(90);
    else if(!await waitForWindowState(function(){return readWindowSlideIndex(windowObject);},current.slideIndex,1500)) return false;
    if(owner!==pptOwnerDocumentKey()) return false;
    const live=panelResolveInstance(activePresentation(),target);
    if(!live || live.slideIndex!==current.slideIndex) return false;
    return selectCanvasShapeAsync(windowObject,live.shape,live.shapeIndex);
  }''')
    text=function(text,'locateMasterShape',r'''  async function locateMasterShape(shapeIndex, shapeReference, instance) {
    return locateTemplatePictureInstance(instance || panelInstanceFromTemplateReference("master",shapeReference));
  }''')
    text=function(text,'locateLayoutShape',r'''  async function locateLayoutShape(layoutIndex, shapeIndex, shapeReference, instance) {
    return locateTemplatePictureInstance(instance || panelInstanceFromTemplateReference("layout",shapeReference));
  }''')
    text=once(text,'    renameShape: renameShape,','''    renameShape: renameShape,
    renamePictureInstances: renamePictureInstances,
    locatePictureInstance: locatePictureInstance,''','export safe instance actions')
    return text


def prepare_target_pane(text):
    start=text.index('      async function locateInstance(inst) {')
    end=text.index('      // Show any runtime error',start)
    text=text[:start]+'''      async function locateInstance(inst) {
        if(!W || !inst || !W.locatePictureInstance) return;
        try {
          const selected=await W.locatePictureInstance(inst);
          flash(selected ? (inst.kind==="slide" ? "已在第 "+inst.slideIndex+" 页正文画布中选中图片" : "已在模板画布中选中图片") : "未能确认图片选中，请刷新清单后重试");
        } catch(error) {flash(error.message||String(error));}
      }

'''+text[end:]
    text=once(text,'''                var n = 0;
                group.instances.forEach(function (inst) {
                  if (inst.shape && W.renameShape(inst.shape, value)) n += 1;
                });
                group.name = value;
                setStatus("已重命名 " + n + " 个实例");''','''                try {
                  var n = W.renamePictureInstances(group.instances,value);
                  if(n>0) group.name=value;
                  setStatus("已重命名 " + n + " 个实例");
                } catch(error) {setStatus(error.message||String(error));}''','rename by stable instance')
    return text
