const {chromium}=require('C:/Users/franc/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright');
const http=require('node:http'),fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict');
const root=path.resolve(__dirname,'..');
const server=http.createServer((req,res)=>{
 const file=path.resolve(root,'.'+new URL(req.url,'http://localhost').pathname.replace(/\/$/,'/index.html'));
 if(!file.startsWith(root+path.sep))return res.writeHead(403).end();
 fs.readFile(file,(err,data)=>{res.writeHead(err?404:200,{'Content-Type':{'.js':'text/javascript','.html':'text/html','.css':'text/css','.wasm':'application/wasm','.png':'image/png'}[path.extname(file)]||'text/plain'});res.end(err?'':data);});
});
const fixture={id:'structure-qa',name:'Peptide 2D',sequence:'KR*CVQRC*KDFLR',resinType:'Rink Amide',resinLoading:.6,scale:.05,equivalents:3,status:'in-progress',dateStarted:'2026-11-02',customMWs:{},notes:'',couplingReagents:[{name:'DIC',eq:3},{name:'Oxyma',eq:3},{name:'DIPEA',eq:6}],aaReagents:{}};
(async()=>{
 let browser;await new Promise(r=>server.listen(0,'127.0.0.1',r));
 try{
  browser=await chromium.launch({channel:'msedge',headless:true});
  const page=await browser.newPage({viewport:{width:1440,height:1000},reducedMotion:'reduce'}),errors=[];
  page.on('pageerror',e=>errors.push(e.message));await page.route('https://**/*',route=>route.abort());
  await page.addInitScript(f=>{sessionStorage.setItem('spps-intro-played-this-session','1');if(!localStorage.getItem('structure-seeded')){localStorage.setItem('structure-seeded','1');localStorage.setItem('spps-lab-assistant-data',JSON.stringify([f]));}},fixture);
  await page.goto('http://127.0.0.1:'+server.address().port);await page.evaluate(()=>navigate('detail','structure-qa'));
  await page.locator('[data-detail-tab="synthesis"]').click();
  const panel=page.locator('.peptide-structure-panel');await panel.locator('svg path.bond-0').first().waitFor({timeout:45000});
  assert((await panel.locator('.structure-metadata').innerText()).includes('1 ponte'));
  assert(await panel.locator('.structure-drawing svg path').count()>100);
  const initial=await panel.locator('.structure-drawing').evaluate(el=>el.clientWidth);
  await panel.locator('[data-structure-action="in"]').click();assert(await panel.locator('.structure-drawing').evaluate(el=>el.clientWidth)>initial);
  await panel.locator('[data-structure-action="fit"]').click();assert.equal(await panel.locator('.structure-drawing').evaluate(el=>el.clientWidth),initial);
  assert.equal(await panel.locator('.structure-drawing').evaluate(el=>el.clientHeight),await panel.locator('.structure-viewport').evaluate(el=>el.clientHeight));
  await panel.locator('[data-structure-action="full"]').click();assert(await page.evaluate(()=>!!document.fullscreenElement));await page.evaluate(()=>document.exitFullscreen());
  const downloadEvent=page.waitForEvent('download');await panel.locator('[data-structure-action="svg"]').click();const download=await downloadEvent;assert(download.suggestedFilename().endsWith('.svg'));
  const molEvent=page.waitForEvent('download');await panel.locator('[data-structure-action="mol"]').click();assert((await molEvent).suggestedFilename().endsWith('.mol'));
  fs.mkdirSync(path.join(root,'test-results'),{recursive:true});
  for(const width of [1440,390]){
   await page.setViewportSize({width,height:1000});await panel.scrollIntoViewIfNeeded();
   assert(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1));
   await page.screenshot({path:path.join(root,'test-results',`peptide-structure-${width}.png`)});
  }
  await page.setViewportSize({width:1440,height:1000});await page.evaluate(()=>{initTheme();document.documentElement.setAttribute('data-theme','dark');});
  await page.screenshot({path:path.join(root,'test-results','peptide-structure-dark.png')});
  await page.evaluate(async()=>{window.print=()=>{};await printSynthesis('structure-qa');});
  assert.equal(await page.locator('#print-area .print-peptide-structure svg').count(),1);
  await page.emulateMedia({media:'print'});await page.screenshot({path:path.join(root,'test-results','peptide-structure-print.png'),fullPage:true});await page.emulateMedia({media:'screen'});
  await page.evaluate(()=>{const s=syntheses.find(s=>s.id==='structure-qa');s.sequence='*KAGD*';render();});await page.locator('[data-detail-tab="synthesis"]').click();
  await page.waitForFunction(()=>document.querySelector('.structure-metadata')?.textContent.includes('C15H26N6O5'));
  await page.evaluate(()=>{const s=syntheses.find(s=>s.id==='structure-qa');s.sequence='A(Unknown)G';render();});
  assert((await page.locator('.structure-status').innerText()).includes('non definita'));
  assert.equal(await page.locator('.structure-drawing svg').count(),0);
  assert.deepEqual(errors,[]);console.log('PASS: real WASM rendering, disulfide/lactam structures, zoom, fit, SVG export, print, mobile/desktop, dark theme and missing-structure state.');
 }finally{if(browser)await browser.close();await new Promise(r=>server.close(r));}
})().catch(e=>{console.error(e);process.exitCode=1;});
