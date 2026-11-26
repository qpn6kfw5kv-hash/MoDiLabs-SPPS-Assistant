const {chromium}=require('C:/Users/franc/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright');
const fs=require('node:fs'),path=require('node:path'),http=require('node:http'),assert=require('node:assert/strict');
const root=path.resolve(__dirname,'..');
const server=http.createServer((req,res)=>{const file=path.resolve(root,'.'+new URL(req.url,'http://localhost').pathname.replace(/\/$/,'/index.html'));if(!file.startsWith(root+path.sep))return res.writeHead(403).end();fs.readFile(file,(error,data)=>{res.writeHead(error?404:200,{'Content-Type':{'.html':'text/html','.js':'text/javascript','.mjs':'text/javascript','.css':'text/css','.png':'image/png','.wasm':'application/wasm'}[path.extname(file)]||'text/plain'});res.end(error?'':data);});});
function pdf(){
 const stream='BT /F1 20 Tf 60 730 Td (HPLC chromatogram - preview test) Tj ET\n60 150 m 60 650 l S\n60 150 m 550 150 l S\n60 150 m 210 150 l 250 600 l 280 150 l 340 150 l 365 420 l 390 150 l 550 150 l S\n';
 const objects=['<< /Type /Catalog /Pages 2 0 R >>','<< /Type /Pages /Kids [3 0 R] /Count 1 >>','<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>','<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',`<< /Length ${stream.length} >>\nstream\n${stream}endstream`];
 let result='%PDF-1.4\n',offsets=[0];objects.forEach((object,i)=>{offsets.push(result.length);result+=`${i+1} 0 obj\n${object}\nendobj\n`;});const xref=result.length;result+=`xref\n0 6\n0000000000 65535 f \n${offsets.slice(1).map(n=>String(n).padStart(10,'0')+' 00000 n \n').join('')}trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF`;
 return 'data:application/pdf;base64,'+Buffer.from(result).toString('base64');
}
(async()=>{let browser;await new Promise(r=>server.listen(0,'127.0.0.1',r));try{
 browser=await chromium.launch({channel:'msedge',headless:true});const page=await browser.newPage({viewport:{width:1440,height:1000}}),errors=[];page.on('pageerror',e=>errors.push(e.message));await page.route('https://**/*',r=>r.abort());await page.addInitScript(()=>sessionStorage.setItem('spps-intro-played-this-session','1'));await page.goto('http://127.0.0.1:'+server.address().port);
 const image=await page.evaluate(()=>{const c=document.createElement('canvas');c.width=1800;c.height=1000;const x=c.getContext('2d');x.fillStyle='white';x.fillRect(0,0,1800,1000);x.fillStyle='black';x.font='36px sans-serif';x.fillText('HPLC - Test di anteprima',80,80);x.strokeStyle='#234ba0';x.lineWidth=3;x.beginPath();for(let t=0;t<1600;t++){const y=850-650*Math.exp(-Math.pow((t-500)/30,2))-350*Math.exp(-Math.pow((t-1000)/55,2));x.lineTo(90+t,y);}x.stroke();return c.toDataURL();});
 fs.mkdirSync(path.join(root,'test-results'),{recursive:true});
 for(const width of [1440,390]){
  await page.setViewportSize({width,height:1000});await page.evaluate(dataUrl=>LabUI.previewFile({name:'Cromatogramma HPLC - campione 01.png',dataUrl}),image);
  const d=page.locator('.file-preview-dialog'),bounds=await d.boundingBox();assert(bounds.width>width*.9);assert(bounds.height>940);assert(await d.locator('img').evaluate(img=>img.complete&&img.naturalWidth===1800));
  await d.locator('[data-preview-action="in"]').click();assert(await d.locator('.file-viewer-viewport').evaluate(el=>el.scrollWidth>el.clientWidth));await d.locator('[data-preview-action="fit"]').click();assert.equal(await d.locator('output').innerText(),'100%');
  const viewport=d.locator('.file-viewer-viewport');await viewport.dispatchEvent('wheel',{deltaY:-400,clientX:100,clientY:200});
  await page.waitForFunction(()=>document.querySelector('.file-viewer-zoom').textContent!=='100%');
  await viewport.evaluate(el=>{el.scrollLeft=150;el.scrollTop=150;});const before=await viewport.evaluate(el=>el.scrollLeft);
  const box=await viewport.boundingBox();await page.mouse.move(box.x+150,box.y+150);await page.mouse.down();await page.mouse.move(box.x+90,box.y+90,{steps:5});await page.mouse.up();assert(await viewport.evaluate(el=>el.scrollLeft)>before);
  await d.locator('[data-preview-action="fit"]').click();
  const event=page.waitForEvent('download');await d.locator('a[download]').click();assert((await event).suggestedFilename().endsWith('.png'));
  await page.screenshot({path:path.join(root,'test-results',`attachment-viewer-${width}.png`)});await page.keyboard.press('Escape');await d.waitFor({state:'detached'});
 }
 await page.setViewportSize({width:1440,height:1000});await page.evaluate(dataUrl=>LabUI.previewFile({name:'Cromatogramma.pdf',dataUrl}),pdf());
 const canvas=page.locator('.file-pdf-pages[data-rendered="true"] canvas');await canvas.waitFor();assert((await canvas.boundingBox()).height>800);
 assert(await canvas.evaluate(c=>{const pixels=c.getContext('2d').getImageData(0,0,c.width,c.height).data;let dark=0;for(let i=0;i<pixels.length;i+=4)if(pixels[i]<100&&pixels[i+1]<100&&pixels[i+2]<100)dark++;return dark>1000;}));
 await page.locator('[data-preview-action="in"]').click();assert.equal(await page.locator('.file-viewer-zoom').innerText(),'140%');await page.locator('[data-preview-action="fit"]').click();
 await page.screenshot({path:path.join(root,'test-results','attachment-viewer-pdf.png')});await page.locator('[data-preview-action="close"]').click();await page.locator('.file-preview-dialog').waitFor({state:'detached'});
 assert.deepEqual(errors,[]);
 console.log('PASS: near-fullscreen image/PDF viewer, mobile/desktop sizing, image zoom/fit, download, Escape, PDF blob lifecycle.');
}finally{if(browser)await browser.close();await new Promise(r=>server.close(r));}})().catch(e=>{console.error(e);process.exitCode=1;});
