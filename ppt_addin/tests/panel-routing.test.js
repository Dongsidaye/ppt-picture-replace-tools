const fs=require('fs'),vm=require('vm'),assert=require('node:assert/strict'),path=require('path');
const main=fs.readFileSync(path.join(__dirname,'../shared/main.js'),'utf8');
const c={console,URL,setTimeout,clearTimeout};c.window=c;
c.location=new URL('https://ppt-toolbox.local/taskpane.html#idle');
c.document={get location(){return c.location;}};
c.wps={CreateTaskPane(url){return {set Visible(value){if(value)c.location=new URL(url);}};}};
vm.createContext(c);vm.runInContext(main,c);
(async()=>{
  for(const tool of ['DesignExportToolsOpen','DesignColorToolsOpen','DesignExportToolsOpen']){
    await c.OpenDesignToolsPane({Id:tool});
    assert.equal(c.location.pathname,'/taskpane.html');
    assert(!c.location.hash.includes('taskpane.html'),'a route must not become part of the URL directory');
    await c.OpenPicturePanel();assert.equal(c.location.hash,'#panel','picture panel must reopen after tool route');
  }
  console.log('PASS actual engine routing: export/color to image list without hash path accumulation');
})().catch(e=>{console.error(e);process.exitCode=1;});
