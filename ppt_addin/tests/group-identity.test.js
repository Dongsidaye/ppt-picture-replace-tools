"use strict";
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");
const assert = require("node:assert/strict");
const child = require("node:child_process");
const root = path.resolve(__dirname, "..");
let main = fs.readFileSync(path.join(root, "shared/main.js"), "utf8");
// Supports checking the generator patch before the root regenerates shared.
if (process.env.PPT_IDENTITY_TEST_PATCH === "1") {
  main = child.execFileSync("python", ["-c", "import sys,pathlib;sys.path.insert(0,sys.argv[1]);from identity_patch import prepare_identity;sys.stdout.write(prepare_identity(pathlib.Path(sys.argv[2]).read_text(encoding='utf-8')))", root, path.join(root, "shared/main.js")], {encoding:"utf8"});
}
const full = {l:0,t:0,r:1,b:1,known:true,raw:[0,0,0,0]};
const strong = value => "sha256:" + crypto.createHash("sha256").update(value).digest("hex");

async function inventory(infos, metadata) {
  const c = {console,setTimeout,clearTimeout,Uint8Array}; c.window=c;
  const shapes=infos.map((info,index)=>({
    Type:13, Name:"Picture "+index, Visible:-1, Left:index*120, Top:0, Width:100, Height:100,
    AlternativeText:metadata ? "PICRENEW|"+JSON.stringify(metadata) : "",
    PictureFormat:{CropLeft:0,CropRight:0,CropTop:0,CropBottom:0}
  }));
  const slide={Shapes:{Count:shapes.length,Item:index=>shapes[index-1]}};
  c.deck={Slides:{Count:1,Item:()=>slide}};
  c.infos=infos.map(info=>Object.assign({dataUrl:"fixture",aspect:1000,cropWindow:full},info));
  const hooks=`
    activePresentation=()=>global.deck; documentKey=()=>"identity-fixture"; yieldUI=async()=>{};
    renderNativePictureThumbnails=async()=>global.infos;
    global.pixelHashes=(pixels)=>({dHash:dHashFromPixels(pixels,9),aHash:aHashFromPixels(pixels,9)});
  `;
  vm.createContext(c);
  vm.runInContext(main.replace("  global.WpsPictureReplace = {",hooks+"\n  global.WpsPictureReplace = {"),c);
  const partials=[];
  const result=await c.WpsPictureReplace.collectDeckImages(null,partial=>partials.push(partial));
  return {result,partials,hashes:c.pixelHashes};
}

function flatPixels(r,g,b) {
  const pixels=new Uint8Array(9*8*4);
  for(let i=0;i<pixels.length;i+=4){pixels[i]=r;pixels[i+1]=g;pixels[i+2]=b;pixels[i+3]=255;}
  return pixels;
}

(async()=>{
  const base=await inventory([]);
  const red=flatPixels(255,0,0),blue=flatPixels(0,0,255);
  const redHashes=base.hashes(red),blueHashes=base.hashes(blue);
  assert.equal(redHashes.dHash,blueHashes.dHash,"fixture must collide in dHash");
  assert.equal(redHashes.aHash,blueHashes.aHash,"fixture must collide in aHash");
  const redInfo=Object.assign({fp:strong(red)},redHashes);
  const blueInfo=Object.assign({fp:strong(blue)},blueHashes);
  assert.notEqual(redInfo.fp,blueInfo.fp);
  const distinct=await inventory([redInfo,blueInfo]);
  assert.equal(distinct.result.groups.length,2,"red and blue with identical grayscale hashes must remain separate replacement groups");
  assert(distinct.result.groups.every(group=>group.key.startsWith("f:sha256:")),"group keys must use strong identity, not low-resolution perceptual hashes");

  const identical=await inventory([redInfo,redInfo]);
  assert.equal(identical.result.groups.length,1,"identical strong identities and crop windows should group");
  assert.equal(identical.result.groups[0].instances.length,2);

  const near=await inventory([redInfo,Object.assign({},blueInfo,{dHash:"0"+blueInfo.dHash.slice(1),aHash:"0"+blueInfo.aHash.slice(1)})]);
  assert.equal(near.result.groups.length,2,"near perceptual hashes must not override different strong identities");

  const unknown=await inventory([{fp:"",...redHashes},{fp:"",...redHashes},{fp:"legacy-fnv32",...redHashes},{fp:"legacy-fnv32",...redHashes}]);
  assert.equal(unknown.result.groups.length,4,"missing or legacy identities must each remain independent");

  const top={l:0,t:0,r:1,b:0.45,known:true,raw:[0,0,0,0]};
  const bottom={l:0,t:0.55,r:1,b:1,known:true,raw:[0,0,0,0]};
  const crops=await inventory([Object.assign({},redInfo,{cropWindow:top}),Object.assign({},redInfo,{cropWindow:bottom}),redInfo]);
  assert.equal(crops.result.groups.length,3,"disjoint crops and the broad full image must not form a transitive replacement group");

  const oldMetadata=await inventory([redInfo,blueInfo],{name:"stale source",contentFp:strong("old embedded content")});
  const structural=oldMetadata.partials.find(partial=>partial.groups);
  assert.equal(structural.groups.length,2,"unverified link metadata must not merge provisional rows");
  assert.equal(oldMetadata.result.groups.length,2,"old metadata cannot override freshly rendered identities");
  console.log("PASS strong source identity, grayscale collisions, legacy/unknown isolation, crop compatibility, and provisional metadata isolation");
})().catch(error=>{console.error(error);process.exitCode=1;});
