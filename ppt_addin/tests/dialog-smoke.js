/* Opt-in host test: cancel two pickers, then choose a test export folder. */
(async function () {
  const host=window.__pptHost, app=window.wps, W=window.WpsPictureReplace;
  const original=app.ActiveWindow, results=[], files=[]; let deck=null;
  window.__pptTesting=true;
  try {
    deck=app.Presentations.Add(-1); deck.Slides.Add(1,12);
    await window.__pptDispatch('DesignExportToolsOpen');
    const exported=await W.designExportSlides('current','png',72);
    if(!exported.cancelled)throw Error('Cancel the export picker for this smoke test.');
    await window.__pptDispatch('OpenPicturePanelButton');
    if(location.hash!=='#panel')throw Error('Image panel did not reopen after export cancellation.');
    results.push({name:'export-cancel-then-image-panel',ok:true});
    const picked=await window.__pptChooseFile({kind:'file',title:'PowerPoint test: cancel this image picker',filter:'*.png;*.jpg'});
    if(!picked.cancelled)throw Error('Cancel the image picker for this smoke test.');
    await window.__pptDispatch('OpenPicturePanelButton');
    if(location.hash!=='#panel'||window.__pptFileDialogPending())throw Error('Picker lock or navigation was not released.');
    results.push({name:'image-picker-cancel-then-image-panel',ok:true});
    const source=String(host.TempFolder)+'\\PPTDialogSmoke_'+Date.now()+'.pptx';
    deck.SaveAs(source,24);files.push(source);
    const completed=await W.designExportSlides('current','png',72);
    if(!completed.ok||completed.count!==1)throw Error('Expected one exported test slide.');
    files.push(...completed.files);
    await window.__pptDispatch('OpenPicturePanelButton');
    if(location.hash!=='#panel')throw Error('Image panel did not reopen after successful export.');
    results.push({name:'export-success-then-image-panel',ok:true});
  } catch(error) { results.push({name:'dialog-smoke',ok:false,error:String(error.stack||error)}); }
  finally {
    try{if(deck){deck.Saved=-1;deck.Close();}}catch(error){results.push({name:'cleanup',ok:false,error:String(error)});}
    files.forEach(path=>{try{host.RemoveFile(path);}catch(_){}});
    try{original.Activate();}catch(_){}
    window.__pptTesting=false;
    host.Report(JSON.stringify({passed:results.filter(x=>x.ok).length,failed:results.filter(x=>!x.ok).length,tests:results},null,2));
  }
}());
