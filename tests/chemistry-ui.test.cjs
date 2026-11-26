const {chromium}=require(process.env.PLAYWRIGHT_MODULE||'C:/Users/franc/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright');
const http=require('node:http'),fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict');
const root=path.resolve(__dirname,'..');
const server=http.createServer((req,res)=>{
 const file=path.resolve(root,'.'+new URL(req.url,'http://localhost').pathname.replace(/\/$/,'/index.html'));
 if(!file.startsWith(root+path.sep))return res.writeHead(403).end();
 fs.readFile(file,(err,data)=>{res.writeHead(err?404:200,{'Content-Type':{'.js':'text/javascript','.html':'text/html','.css':'text/css','.png':'image/png'}[path.extname(file)]||'text/plain'});res.end(err?'':data);});
});
const fixture={id:'chemistry-qa',name:'Chemical QA',sequence:'A(Aib)*KAGD*G',resinType:'Rink Amide',resinLoading:.6,scale:.05,equivalents:3,status:'in-progress',dateStarted:'2026-10-06',customMWs:{},notes:'',couplingReagents:[{name:'DIC',eq:3},{name:'Oxyma',eq:3},{name:'DIPEA',eq:6}],aaReagents:{}};
(async()=>{
 let browser;await new Promise(r=>server.listen(0,'127.0.0.1',r));
 try{
  browser=await chromium.launch({channel:'msedge',headless:true});
  const page=await browser.newPage({viewport:{width:1440,height:1000},reducedMotion:'reduce'}),errors=[],queries=[];
  page.on('pageerror',e=>errors.push(e.message));
  await page.route('https://**/*',async route=>{
   const url=decodeURIComponent(route.request().url());
   if(!url.includes('/rest/pug/'))return route.abort();
   queries.push(url);
   if(url.includes('/synonyms/'))return route.fulfill({json:{InformationList:{Information:[{Synonym:['Fmoc-Nva-OH']}]}}});
   const nva=url.includes('Fmoc-Nva-OH');
   return route.fulfill({json:{PropertyTable:{Properties:[{CID:nva?1000:1001,MolecularWeight:nva?'339.39':'103.12',MolecularFormula:nva?'C20H21NO4':'C4H9NO2'}]}}});
  });
  await page.addInitScript(f=>{sessionStorage.setItem('spps-intro-played-this-session','1');if(!localStorage.getItem('chemistry-seeded')){localStorage.setItem('chemistry-seeded','1');localStorage.setItem('spps-lab-assistant-data',JSON.stringify([f,{...f,id:'legacy',name:'Legacy',sequence:'AG',couplingReagents:undefined,activators:['DIC/Oxyma'],aaActivators:{}}]));}},fixture);
  await page.goto('http://127.0.0.1:'+server.address().port);
  await page.evaluate(()=>navigate('detail','chemistry-qa'));
  await page.locator('[data-detail-tab="synthesis"]').click();
  assert.deepEqual(await page.locator('.sequence-progress-strip > button').evaluateAll(nodes=>nodes.map(n=>Number(n.dataset.sequenceIndex))),[6,5,4,3,2,1,0]);
  assert.equal(await page.locator('[data-order]').count(),0);
  assert.deepEqual(await page.locator('.sequence-progress-index').allTextContents(),['1','2','3','4','5','6','7']);
  const lys=page.locator('[data-step-index="2"]');assert((await lys.innerText()).includes('67.88'));
  const aib=page.locator('[data-step-index="1"]');assert((await aib.innerText()).includes('325.36'));assert((await aib.innerText()).includes('48.80'));
  assert.equal(await lys.locator('.coupling-dose').count(),3);
  for(const value of ['23.26 µL','21.32 mg','52.25 µL'])assert((await lys.innerText()).includes(value));
  await lys.locator('summary').click();await lys.locator('input[value="Oxyma"]').uncheck();
  assert.equal(await page.locator('[data-step-index="2"] .coupling-dose').count(),2);
  await page.locator('.detail-header [data-ui-action="lab"]').click();
  for(const name of ['DIC','Oxyma','DIPEA'])assert((await page.locator('#lab-mode-view').innerText()).includes(name));
  await page.locator('[data-ui-action="exit-lab"]').click();
  const printHtml=await page.evaluate(async()=>{window.print=()=>{};await printSynthesis('chemistry-qa');return document.querySelector('#print-area').innerHTML;});
  assert(printHtml.includes('52.25 µL')&&printHtml.includes('23.26 µL')&&printHtml.includes('21.32 mg'));
  assert.equal((printHtml.match(/1\. Fmoc-Gly-OH/g)||[]).length,1);
  assert(printHtml.includes('325.36')&&printHtml.includes('48.80 mg'));
  await page.evaluate(()=>document.querySelector('#print-area')?.remove());
  await page.locator('#btn-edit-params').click();
  assert.equal(await page.locator('#edit-activator-1 option[value="DIC/Oxyma"]').count(),0);
  await page.locator('#edit-reagent-concentration-1').fill('.5');
  await page.locator('#edit-reagent-eq-2').fill('4');await page.locator('#edit-modal-save').click();
  await page.locator('[data-detail-tab="synthesis"]').click();
  assert((await page.locator('[data-step-index="0"]').innerText()).includes('300.00 µL'));
  await page.reload();await page.evaluate(()=>navigate('detail','chemistry-qa'));await page.locator('[data-detail-tab="synthesis"]').click();
  assert((await page.locator('[data-step-index="0"]').innerText()).includes('300.00 µL'));
  await page.evaluate(()=>navigate('detail','legacy'));await page.locator('[data-detail-tab="synthesis"]').click();
  assert.equal(await page.locator('[data-step-index="0"] .coupling-dose').count(),2);
  await page.locator('#btn-edit-params').click();await page.locator('#edit-modal-save').click();
  assert(await page.evaluate(()=>!syntheses.find(s=>s.id==='legacy').couplingReagents));
  await page.evaluate(()=>{const s=syntheses.find(s=>s.id==='chemistry-qa');s.sequence='G(Nva)(Unknown)';s.customMWs.Unknown=103.12;saveSyntheses(syntheses);navigate('detail',s.id);});
  await page.waitForFunction(()=>syntheses.find(s=>s.id==='chemistry-qa').chemicalRecords?.Unknown?.status==='unresolved');
  await page.waitForFunction(()=>syntheses.find(s=>s.id==='chemistry-qa').chemicalRecords?.Nva?.status==='verified');
  assert(await page.evaluate(()=>resolveAminoAcid('Unknown',syntheses.find(s=>s.id==='chemistry-qa')).status==='unresolved'));
  assert(queries.every(url=>!url.includes('/name/Unknown/')&&!url.includes('/name/Nva/')));
  await page.locator('[data-detail-tab="synthesis"]').click();assert(!(await page.locator('#workspace-content').innerText()).includes('NaN'));
  await page.locator('.chemical-mw-input[data-token="Unknown"]').fill('555.12');await page.locator('.chemical-mw-input[data-token="Unknown"]').dispatchEvent('change');
  assert(await page.evaluate(()=>syntheses.find(s=>s.id==='chemistry-qa').chemicalRecords.Unknown.status==='manual'));
  for(const width of [1440,390]){
   await page.setViewportSize({width,height:900});await page.locator('#btn-edit-params').click();
   assert(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1));
   await page.locator('#edit-activator-1').scrollIntoViewIfNeeded();
   fs.mkdirSync(path.join(root,'test-results'),{recursive:true});await page.screenshot({path:path.join(root,'test-results',`chemistry-form-${width}.png`)});
   await page.locator('#edit-modal-cancel').click();
  }
  assert.deepEqual(errors,[]);console.log('PASS: protected AA doses, independent reagents, bench/print parity, per-step selection, solution volumes, persistence, legacy compatibility, strict PubChem identity, unverified cache rejection, desktop/mobile forms. PubChem responses mocked.');
 }finally{if(browser)await browser.close();await new Promise(r=>server.close(r));}
})().catch(e=>{console.error(e);process.exitCode=1;});
