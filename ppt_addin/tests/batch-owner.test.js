"use strict";
const fs=require("node:fs"),path=require("node:path"),vm=require("node:vm"),assert=require("node:assert/strict");
const main=fs.readFileSync(path.join(__dirname,"../shared/main.js"),"utf8");

function fixture(options={}) {
  const c={console,setTimeout,clearTimeout,key:"A",writes:[],links:[],yields:0,disposed:0,options};
  c.window=c;c.__pptDocumentKey=()=>c.key;
  const shapes=[1,2].map(id=>({Type:13,Id:id,Name:"A"+id,AlternativeText:"",Parent:{SlideID:100}}));
  const slide={Shapes:{Count:2,Item:i=>shapes[i-1]}};
  c.deck={Slides:{Count:1,Item:()=>slide}};
  c.instances=shapes.map(shape=>({shape,linked:true,src:"replacement.png",uid:shape.Name}));
  const hooks=`
    activePresentation=()=>global.deck;
    createScratchManager=()=>({dispose(){global.disposed++;}});
    panelPreparePictureSources=async()=>{if(global.options.switchDuringPrepare)global.key="B";};
    panelValidateInstanceSources=async()=>{if(global.options.switchDuringPrepare)global.key="B";};
    panelPictureSources={"slide:100:1":{fingerprint:"sha256:"+"a".repeat(64)},"slide:100:2":{fingerprint:"sha256:"+"a".repeat(64)}};
    yieldUI=async()=>{global.yields++;if(global.yields===global.options.switchAtYield)global.key="B";};
    replacePictureKeepCrop=shape=>{global.writes.push({shape:shape.Name,active:global.key});return shape;};
    contentFingerprintAndThumb=async()=>{if(global.options.switchDuringFingerprint)global.key="B";return {fp:"fixture",aspect:1000};};
    attachLink=shape=>global.links.push({shape:shape.Name,active:global.key});
    resolveTargets=(_,instances)=>instances.map(inst=>({inst,shape:inst.shape}));
    fingerprintFile=()=>"fixture";fileExists=()=>true;assertDocument=()=>{};invalidatePanelInventoryCache=()=>{};
    global.matching=replaceAllMatching;global.instancesReplace=replaceInstances;global.linkedUpdate=updateLinkedInstances;
  `;
  vm.createContext(c);vm.runInContext(main.replace("  global.WpsPictureReplace = {",hooks+"\n  global.WpsPictureReplace = {"),c);
  c.call=name=>name==="matching" ? c.matching(shapes[0],"replacement.png")
    : name==="instancesReplace" ? c.instancesReplace(c.instances,"replacement.png","fixture")
    : c.linkedUpdate(c.instances,"fixture");
  return c;
}

(async()=>{
  for(const name of ["matching","instancesReplace","linkedUpdate"]){
    const stable=fixture();await stable.call(name);
    assert.equal(stable.writes.length,2,name+" must still replace both sources in the owner document");
    assert(stable.writes.every(entry=>entry.active==="A"));

    const preparing=fixture({switchDuringPrepare:true});
    await assert.rejects(preparing.call(name),/文稿已切换/);
    assert.equal(preparing.writes.length,0,name+" must make zero writes when source preparation enters another document");

    const between=fixture({switchAtYield:name==="matching"?4:1});
    await assert.rejects(between.call(name),/文稿已切换/);
    assert.equal(between.writes.length,1,name+" must stop before the next target after switching");
    assert.equal(between.writes.filter(entry=>entry.active==="B").length,0,name+" must not write the old document while another is active");
    assert.equal(between.disposed,1,"owned scratch must be disposed on owner rejection");
  }
  for(const switchAtYield of [1,3]){
    const c=fixture({switchAtYield});await assert.rejects(c.call("matching"),/文稿已切换/);
    assert.equal(c.writes.length,0,"switch during scanning or immediately before replacement must produce zero writes");
  }
  for(const name of ["instancesReplace","linkedUpdate"]){
    const c=fixture({switchDuringFingerprint:true});await assert.rejects(c.call(name),/文稿已切换/);
    assert.equal(c.writes.length,1,"only the operation completed before the switch may remain");
    assert.equal(c.links.length,0,"source metadata must not be attached after switching documents");
  }
  console.log("PASS batch owner isolation: all three replacement flows stop after document changes, with zero writes in the wrong active document");
})().catch(error=>{console.error(error);process.exitCode=1;});
