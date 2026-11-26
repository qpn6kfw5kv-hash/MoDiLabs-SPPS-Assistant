const {chromium}=require(process.env.PLAYWRIGHT_MODULE || 'C:/Users/franc/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright');
const http=require('node:http'),fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict');
const root=path.resolve(__dirname,'..');
const server=http.createServer((req,res)=>{
 const file=path.resolve(root,'.'+new URL(req.url,'http://localhost').pathname.replace(/\/$/,'/index.html'));
 if(!file.startsWith(root+path.sep))return res.writeHead(403).end();
 fs.readFile(file,(err,data)=>{res.writeHead(err?404:200,{'Content-Type':{'.js':'text/javascript','.html':'text/html','.css':'text/css','.png':'image/png'}[path.extname(file)]||'text/plain'});res.end(err?'':data);});
});
(async()=>{
 let browser;await new Promise(r=>server.listen(0,'127.0.0.1',r));
 try{
  browser=await chromium.launch({channel:'msedge',headless:true});
  for(const width of [1440,390]){
   const page=await browser.newPage({viewport:{width,height:900}});
   await page.route('https://**/*',r=>r.abort());
   await page.goto('http://127.0.0.1:'+server.address().port);
   await page.waitForFunction(()=>document.body.classList.contains('intro-running'));
   const samples=await page.evaluate(()=>{
    const mark=document.querySelector('.site-intro-mark');
    const animations=document.getAnimations();animations.forEach(a=>a.pause());
    return [0,1472,2176,3100].map(t=>{
     animations.forEach(a=>a.currentTime=t);
     const rect=mark.getBoundingClientRect();
     return {width:rect.width,left:rect.left,right:rect.right,opacity:parseFloat(getComputedStyle(mark).opacity)};
    });
   });
   assert(samples[1].width>samples[0].width*5);
   assert(samples[2].width<samples[1].width*1.1);
   assert(samples[3].width>samples[2].width*5);
   assert(samples[1].left>=0 && samples[1].right<=width);
   await page.evaluate(()=>document.getAnimations().forEach(a=>a.currentTime=1600));
   fs.mkdirSync(path.join(root,'test-results'),{recursive:true});
   await page.screenshot({path:path.join(root,'test-results',`intro-${width}.png`)});
   await page.evaluate(()=>document.getAnimations().forEach(a=>{a.currentTime=3100;a.play();}));
   await page.waitForFunction(()=>document.querySelector('.site-intro').classList.contains('is-complete'));
   assert.equal(await page.locator('#app').evaluate(el=>getComputedStyle(el).opacity),'1');
   await page.reload();assert(await page.locator('.site-intro').evaluate(el=>el.classList.contains('is-skipped')));
   await page.close();
  }
  const page=await browser.newPage({reducedMotion:'reduce'});
  await page.route('https://**/*',r=>r.abort());await page.goto('http://127.0.0.1:'+server.address().port);
  assert(await page.locator('.site-intro').evaluate(el=>el.classList.contains('is-skipped')));
  console.log('PASS: depth zoom, readable pause, forward fly-through, homepage reveal, desktop/mobile framing, once per session, reduced motion.');
 }finally{if(browser)await browser.close();await new Promise(r=>server.close(r));}
})().catch(e=>{console.error(e);process.exitCode=1;});
