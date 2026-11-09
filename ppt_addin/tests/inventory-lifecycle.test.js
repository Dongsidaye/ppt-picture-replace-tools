"use strict";
const fs=require("node:fs"),path=require("node:path"),vm=require("node:vm"),assert=require("node:assert/strict");
const main=fs.readFileSync(path.join(__dirname,"../shared/main.js"),"utf8");
function context() {
  const c={console,setTimeout,clearTimeout,setInterval,clearInterval,Uint8Array};c.window=c;c.owner="doc-A";c.__pptDocumentKey=()=>c.owner;
  c.deck={Slides:{Count:1}};c.events=[];c.prepareCalls=0;c.fakeCollect=async()=>({groups:[],docKey:c.owner,total:0});c.fakeYield=async()=>{};
  const hooks=`
    activePresentation=()=>global.deck; documentKey=()=>global.owner;
    panelInventoryWriteBusy=()=>{};panelInventoryClearBusy=()=>{};panelInventoryTouchBusy=()=>{};
    panelInventoryRemovePersisted=()=>{};panelInventoryPersist=()=>{};
    collectDeckImages=(...args)=>global.fakeCollect(...args);yieldUI=()=>global.fakeYield();
    panelPreparePictureSources=async()=>{global.prepareCalls++;};
    global.api={scan:panelInventoryRunScan,invalidate:invalidatePanelInventoryCache,deliver:panelInventoryDeliverCached,
      cached:()=>panelInventoryMemoryCache,crop:cropWindowCompatible,structure:panelInventoryStructure,collect:collectDeckImagesCached};
  `;
  vm.createContext(c);vm.runInContext(main.replace("  global.WpsPictureReplace = {",hooks+"\n  global.WpsPictureReplace = {"),c);return c;
}
(async()=>{
  let failures=0;
  async function test(name,fn){try{await fn();console.log("PASS",name);}catch(e){failures++;console.error("FAIL",name,e.message);}}
  await test("invalidated scan must reject, never return stale rows",async()=>{
    const c=context();c.fakeCollect=async()=>{c.api.invalidate();return {groups:[{name:"stale"}],docKey:c.owner};};
    await assert.rejects(c.api.scan(c.deck,c.owner));assert.equal(c.api.cached(),null);
  });
  await test("late partial callbacks cannot publish invalidated rows",async()=>{
    const c=context();c.fakeCollect=async(progress,partial)=>{c.api.invalidate();partial({groups:[{name:"wrong"}]});return {groups:[],docKey:c.owner};};
    try{await c.api.scan(c.deck,c.owner,null,p=>c.events.push(p));}catch(_){}assert.equal(c.events.length,0);
  });
  await test("cached delivery must reject after document switches during yield",async()=>{
    const c=context();c.fakeYield=async()=>{c.owner="doc-B";};
    await assert.rejects(c.api.deliver({groups:[],docKey:"doc-A"},null,p=>c.events.push(p),"memory"));assert.equal(c.events.length,0);
  });
  await test("overlapping but different visible crops stay separate",async()=>{
    const c=context();assert.equal(c.api.crop({l:0,t:0,r:0.8,b:1,known:true},{l:0.2,t:0,r:1,b:1,known:true}),false);
    assert.equal(c.api.crop({l:0,t:0,r:1,b:1,known:true},{l:0,t:0,r:1,b:1,known:true}),true);
  });
  await test("new unsupported picture during a scan invalidates the entire result",async()=>{
    const c=context(),shapes=[{Id:1,Type:13}];
    const slide={SlideID:100,Shapes:{get Count(){return shapes.length;},Item:i=>shapes[i-1]}};
    c.deck={Slides:{Count:1,Item:()=>slide},Designs:{Count:0}};
    c.fakeCollect=async()=>{shapes.push({Id:2,Type:28});return {groups:[],docKey:c.owner};};
    await assert.rejects(c.api.scan(c.deck,c.owner));assert.equal(c.api.cached(),null);
  });
  await test("reopening during a scan joins it without creating another source snapshot",async()=>{
    const c=context();let finish;
    c.fakeCollect=()=>new Promise(resolve=>{finish=resolve;});
    const scan=c.api.scan(c.deck,c.owner),reopen=c.api.collect();
    assert.equal(c.prepareCalls,0,"joining must not retire source files still being rendered");
    finish({groups:[],docKey:c.owner,total:0});await scan;
    const result=await reopen;assert.equal(result.cacheSource,"preload");assert.equal(c.prepareCalls,0);
  });
  await test("opening after an invalidation retries the old scan once for the same document",async()=>{
    const c=context();let finish;
    c.fakeCollect=()=>new Promise(resolve=>{finish=resolve;});
    const old=c.api.scan(c.deck,c.owner).catch(()=>null);
    c.api.invalidate();c.fakeCollect=async()=>({groups:[{name:"current"}],docKey:c.owner,total:1});
    const reopened=c.api.collect();finish({groups:[{name:"outdated"}],docKey:c.owner,total:1});await old;
    const result=await reopened;assert.equal(result.groups[0].name,"current");
  });
  if(failures)process.exitCode=1;
})();
