"use strict";
const fs=require("node:fs"),path=require("node:path"),vm=require("node:vm"),assert=require("node:assert/strict");
let main=fs.readFileSync(path.join(__dirname,"../shared/main.js"),"utf8");
function fixture() {
  const c={console,setTimeout,clearTimeout,key:"window:A",writes:[],gotos:[],selections:[],currentSlide:1};c.window=c;
  c.__pptDocumentKey=()=>c.key;
  c.__pptSameObject=(a,b)=>!!a&&!!b&&(a===b || (a.comIdentity&&a.comIdentity===b.comIdentity));
  c.__pptLoadPictureSources=()=>{};
  const slides=[];
  function slide(id) {
    const objects=[];
    const value={SlideID:id,objects,get SlideIndex(){return slides.indexOf(value)+1;},Select(){c.currentSlide=value.SlideIndex;}};
    value.Shapes={get Count(){return objects.length;},Item(index){const result=typeof index==="string"?objects.find(s=>s.Name===index):objects[index-1];if(!result)throw Error("Shape missing");return result;}};
    slides.push(value);return value;
  }
  const a=slide(100),b=slide(200);
  function shape(parent,id,name="Picture 2") {
    let alt='PICRENEW|{"src":"input.png","userAlt":"original accessibility"}',label=name;
    const value={Type:13,Id:id,Parent:parent,Left:0,Top:0,Width:100,Height:100,Rotation:0,HorizontalFlip:0,VerticalFlip:0,Visible:-1,
      PictureFormat:{CropLeft:0,CropRight:0,CropTop:0,CropBottom:0},
      get Name(){return label;},set Name(next){c.writes.push({op:"rename",id,active:c.key});label=next;},
      get AlternativeText(){return alt;},set AlternativeText(next){c.writes.push({op:"unlink",id,active:c.key});alt=next;},
      Select(){c.selections.push(id);}};
    parent.objects.push(value);return value;
  }
  const target=shape(a,2),other=shape(a,3),otherPage=shape(b,2);
  const deck={Designs:{Count:0},Slides:{get Count(){return slides.length;},Item(index){return slides[index-1];},FindBySlideID(id){return slides.find(s=>s.SlideID===id);}}};
  a.Parent=deck;b.Parent=deck;c.deck=deck;
  c.nativeWindow={ViewType:9,View:{GotoSlide(index){c.gotos.push(index);c.currentSlide=index;}}};
  c.instance={kind:"slide",slideIndex:1,slideIdentity:100,shapeIndex:1,shapeIdentity:2,shape:target,sourceKey:"slide:100:2",sourceIdentity:"",ownerDocumentKey:"window:A"};
  const hooks=`
    activePresentation=()=>global.deck;application=()=>({ActiveWindow:global.nativeWindow,ActivePresentation:global.deck});
    assertDocument=()=>{};invalidatePanelInventoryCache=()=>{};yieldUI=async()=>{if(global.onYield)global.onYield();};
    requestWindowViewType=async()=>{if(global.onViewRequest)global.onViewRequest();return true;};layerGuardShapeLocked=()=>false;
    readWindowSlideIndex=()=>global.currentSlide;waitForWindowState=async(read,target)=>read()===target;
    selectCanvasShapeAsync=async(_,shape)=>{global.selections.push(shape.Id);return !shape.Parent.SlideIndex || shape.Parent.SlideIndex===global.currentSlide;};
    panelPreparePictureSources=async()=>{if(global.prepareSource)await global.prepareSource();panelPictureSources=global.sources||{};};
    global.resolve=resolveTargets;global.cachedResolve=panelInventoryResolvedShape;global.unlink=unlinkInstances;
    global.locate=locateSlideShape;global.matches=selectionMatchesShape;global.rename=renameShape;
    global.renameInstances=renamePictureInstances;global.locateInstance=locatePictureInstance;
    global.validateSources=panelValidateInstanceSources;
    global.setSources=(sources)=>{panelPictureSources=sources;global.sources=sources;};
    global.snapshot=panelInventorySnapshot;global.hydrate=panelInventoryHydrate;global.keyFor=documentKey;
  `;
  vm.createContext(c);vm.runInContext(main.replace("  global.WpsPictureReplace = {",hooks+"\n  global.WpsPictureReplace = {"),c);
  function templates() {
    function container(token) {
      const result={comIdentity:token,objects:[],Select(){c.currentTemplate=token;}};
      result.Shapes={get Count(){return result.objects.length;},Item:i=>result.objects[i-1]};return result;
    }
    const master=container("master1"),one=container("layout1"),two=container("layout2"),layouts=[one,two];
    master.CustomLayouts={get Count(){return layouts.length;},Item:i=>Object.assign({},layouts[i-1])};
    deck.Designs={Count:1,Item:()=>({SlideMaster:Object.assign({},master)})};
    const picture=shape(one,12);
    const instance={kind:"layout",layoutIndex:1,shapeIndex:1,shapeIdentity:12,shape:picture,
      containerReference:Object.assign({},one),sourceKey:"layout:1:1:12",ownerDocumentKey:"window:A"};
    return {master,one,two,layouts,picture,instance};
  }
  return {c,slides,a,b,target,other,otherPage,deck,templates};
}

const failures=[];
async function test(name,run){try{await run();console.log("PASS "+name);}catch(error){failures.push({name,error:String(error.message||error)});console.error("FAIL "+name+": "+String(error.message||error));}}
(async()=>{
  await test("cached and action targets resolve by stable IDs after object insertion",()=>{
    const f=fixture();f.a.objects.reverse();const inst=Object.assign({},f.c.instance,{shape:null});
    assert.equal(f.c.cachedResolve(f.deck,inst),f.target,"cache binding must use Shape.Id rather than old collection position");
    const resolved=f.c.resolve(f.deck,[inst]);assert.equal(resolved.length,1);assert.equal(resolved[0].shape,f.target);
  });
  await test("targets survive slide reordering without borrowing the same local Shape.Id",()=>{
    const f=fixture();f.slides.reverse();const inst=Object.assign({},f.c.instance,{shape:null});
    const resolved=f.c.resolve(f.deck,[inst]);assert.equal(resolved.length,1);assert.equal(resolved[0].shape,f.target,"Shape.Id is only unique inside its original SlideID");
  });
  await test("deleted target never unlinks the object now occupying its old index",async()=>{
    const f=fixture();f.a.objects.shift();
    const stale={Id:2,Parent:f.a,get PictureFormat(){throw Error("Deleted COM shape");}};
    const inst=Object.assign({},f.c.instance,{shape:stale});
    let failure=null;try{await f.c.unlink([inst],"fixture");}catch(error){failure=error;}
    assert.equal(f.c.writes.length,0,"failed stable resolution must fail closed instead of unlinking the next picture");
  });
  await test("locating a live object uses its current slide index",async()=>{
    const f=fixture();f.slides.reverse();const ok=await f.c.locate(1,1,f.target);
    assert.equal(f.c.gotos[f.c.gotos.length-1],2,"the cached page number must not override the resolved slide identity");
    assert.equal(ok,true);
  });
  await test("duplicate names never override a known different shape ID",()=>{
    const f=fixture();const window={Selection:{ShapeRange:{Count:1,Item:()=>f.other}}};
    assert.equal(f.c.matches(window,f.target),false,"selection must not claim the wrong same-named shape is the target");
  });
  await test("rename targets the original IDs and deduplicates repeated inputs",()=>{
    const f=fixture();f.a.objects.reverse();f.slides.reverse();
    assert.equal(f.c.renameInstances([f.c.instance,Object.assign({},f.c.instance)],"Renamed"),1);
    assert.equal(f.target.Name,"Renamed");assert.equal(f.other.Name,"Picture 2");assert.equal(f.otherPage.Name,"Picture 2");
    assert.equal(f.c.writes.filter(w=>w.op==="rename").length,1);
  });
  await test("invalid owner rejects every rename before the first write",()=>{
    const f=fixture();assert.throws(()=>f.c.renameInstances([f.c.instance,Object.assign({},f.c.instance,{ownerDocumentKey:"window:B"})],"Wrong"));
    assert.equal(f.c.writes.length,0);
    f.c.key="window:B";assert.equal(f.c.cachedResolve(f.deck,f.c.instance),null);
  });
  await test("source verification resolves live IDs and rejects changed appearance",async()=>{
    const f=fixture();f.a.objects.reverse();
    const inst=Object.assign({},f.c.instance,{shape:f.other,sourceIdentity:"media",sourceAppearance:"original"});
    f.c.setSources({"slide:100:2":{fingerprint:"media",appearance:"original"}});
    await f.c.validateSources([inst]);assert.equal(inst.shape,f.target);
    f.c.setSources({"slide:100:2":{fingerprint:"media",appearance:"cropped"}});
    await assert.rejects(()=>f.c.validateSources([inst]),/裁剪/);assert.equal(f.c.writes.length,0);
  });
  await test("switching documents during source verification rejects stale instances",async()=>{
    const f=fixture();f.c.prepareSource=async()=>{f.c.key="window:B";};
    await assert.rejects(()=>f.c.validateSources([f.c.instance]));assert.equal(f.c.writes.length,0);
  });
  await test("unknown to known, known to unknown, and fully unknown identities cannot authorize stale replacement",async()=>{
    for(const mode of ["unknown-to-known","known-to-unknown","unknown-to-unknown"]) {
      const f=fixture(),known=mode==="known-to-unknown";
      const inst=Object.assign({},f.c.instance,{sourceIdentity:known?"media":"",sourceAppearance:known?"appearance":""});
      f.c.setSources(mode==="unknown-to-known"?{"slide:100:2":{fingerprint:"new-media",appearance:"new-look"}}:{});
      await assert.rejects(()=>f.c.validateSources([inst]),/身份|变化/);assert.equal(f.c.writes.length,0);
    }
  });
  await test("cache hydration refreshes current page and object positions",()=>{
    const f=fixture(),hash="sha256:"+"1".repeat(64);
    const inst=Object.assign({},f.c.instance,{sourceIdentity:hash,sourceAppearance:"appearance",shapeName:f.target.Name,
      left:0,top:0,width:100,height:100,rotation:0,flipH:false,flipV:false,visible:true,cropWindow:{l:0,t:0,r:1,b:1,known:true,raw:[0,0,0,0]}});
    f.c.setSources({"slide:100:2":{fingerprint:hash,appearance:"appearance"}});
    const snapshot=f.c.snapshot({docKey:f.c.keyFor(f.deck),groups:[{contentFp:hash,instances:[inst]}],total:1,slideCount:2});
    f.a.objects.reverse();f.slides.reverse();
    const result=f.c.hydrate(snapshot,f.deck);assert.ok(result);const restored=result.groups[0].instances[0];
    assert.equal(restored.shape,f.target);assert.equal(restored.slideIndex,2);assert.equal(restored.shapeIndex,2);
    f.c.key="window:B";assert.equal(f.c.hydrate(snapshot,f.deck),null);
  });
  await test("template ownership uses native identity and follows reordered layouts",async()=>{
    const f=fixture(),t=f.templates();t.layouts.reverse();
    const targets=f.c.resolve(f.deck,[t.instance]);assert.equal(targets[0].shape,t.picture);
    assert.equal(t.instance.layoutIndex,2);assert.equal(t.instance.sourceKey,"layout:1:2:12");
    assert.equal(await f.c.locateInstance(t.instance),true);assert.equal(f.c.currentTemplate,"layout1");
  });
  await test("missing template live identity never falls back to an old ordinal",()=>{
    const f=fixture(),t=f.templates();t.instance.containerReference=null;
    assert.throws(()=>f.c.renameInstances([t.instance],"Wrong"));assert.equal(f.c.writes.length,0);
  });
  await test("switching documents during template navigation never selects stale shape",async()=>{
    const f=fixture(),t=f.templates();f.c.onViewRequest=()=>{f.c.key="window:B";};
    assert.equal(await f.c.locateInstance(t.instance),false);assert.equal(f.c.selections.length,0);
  });
  await test("equal shape IDs on different template containers do not confirm selection",()=>{
    const f=fixture(),t=f.templates();
    const wrong={Id:12,Name:t.picture.Name,Parent:t.two};
    assert.equal(f.c.matches({Selection:{ShapeRange:{Count:1,Item:()=>wrong}}},t.picture),false);
  });
  assert.equal(failures.length,0,JSON.stringify(failures,null,2));
})().catch(error=>{console.error(error.message||error);process.exitCode=1;});
