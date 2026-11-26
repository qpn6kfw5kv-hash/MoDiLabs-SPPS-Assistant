const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'C:/Users/franc/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright');
const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
const server = http.createServer((req, res) => {
  const file = path.resolve(root, '.' + new URL(req.url, 'http://localhost').pathname.replace(/\/$/, '/index.html'));
  if (!file.startsWith(root + path.sep)) return res.writeHead(403).end();
  fs.readFile(file, (err, data) => {
    res.writeHead(err ? 404 : 200, { 'Content-Type': { '.js': 'text/javascript', '.html': 'text/html', '.css': 'text/css', '.png': 'image/png' }[path.extname(file)] || 'text/plain' });
    res.end(err ? '' : data);
  });
});
// Only the external SDK is substituted. The shipped UI and services run unchanged.
const sdk = `export function createClient(){
 let listener, user=JSON.parse(sessionStorage.getItem('test-user')||'null');
 const rows=[];
 const emit=event=>listener(event,user?{user}:null);
 return {auth:{
 onAuthStateChange(fn){listener=fn;return {data:{subscription:{unsubscribe(){}}}}},
 async getUser(){return {data:{user},error:null}},
 async signInWithOtp(){return {error:null}},
 async verifyOtp({email,token}){if(token!=='123456')return {error:{code:'otp_expired'}};
 user={id:'user-a',email};sessionStorage.setItem('test-user',JSON.stringify(user));emit('SIGNED_IN');return {data:{user}}},
 async signOut(){if(window.failLogout)return {error:{message:'network'}};user=null;sessionStorage.removeItem('test-user');emit('SIGNED_OUT');return {error:null}}
 },from(table){let column,value;const query={select(){return query},eq(c,v){column=c;value=v;return query},is(){return query},order(){return query},in(){return query},
 then(resolve,reject){return Promise.resolve({data:table==='profiles'?[{id:'user-a',display_name:'Researcher'}]:column==='user_id'?rows.map(w=>({role:'owner',workspaces:w})):rows.some(w=>w.id===value)?[{user_id:'user-a',role:'owner'}]:[],error:null}).then(resolve,reject)}};return query},
 async rpc(name,args){const id='ws-'+rows.length;rows.push({id,name:args.p_name,owner_id:user.id,version:1});return {data:id,error:null}}
 }};`;
(async () => {
  let browser;
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  try {
    browser = await chromium.launch({channel:'msedge',headless:true});
    const page = await browser.newPage({viewport:{width:1280,height:900},reducedMotion:'reduce'});
    const errors=[];page.on('pageerror',e=>errors.push(e.message));
    await page.route('https://**/*', route => route.request().url().includes('@supabase/supabase-js@') ? route.fulfill({contentType:'text/javascript',body:sdk}) : route.abort());
    await page.addInitScript(()=>{sessionStorage.setItem('spps-intro-played-this-session','1');localStorage.setItem('spps-lab-assistant-data','[]');});
    const url='http://127.0.0.1:'+server.address().port;
    await page.goto(url);await page.locator('.lab-shell').waitFor();
    await page.evaluate(()=>SPPSCloud.open());await page.getByText('Cloud non attivo.',{exact:false}).waitFor();
    await page.getByRole('button',{name:'Modalità locale',exact:true}).click();await page.locator('.lab-shell').waitFor();
    await page.route('**/cloud-config.js*',r=>r.fulfill({contentType:'text/javascript',body:"window.SPPS_CONFIG={cloudEnabled:true,supabaseUrl:'https://example.supabase.co',supabasePublicKey:'sb_publishable_test'}"}));
    await page.reload();await page.locator('#cloud-email').fill('researcher@example.com');
    await page.getByRole('button',{name:'Invia codice di accesso'}).click();
    await page.locator('#cloud-code').fill('000000');await page.getByRole('button',{name:'Accedi',exact:true}).click();
    await page.getByRole('alert').filter({hasText:'scaduto'}).waitFor();
    await page.locator('#cloud-code').fill('123456');await page.getByRole('button',{name:'Accedi',exact:true}).click();
    await page.locator('#cloud-workspace-name').fill('<img src=x onerror=alert(1)>');
    await page.getByRole('button',{name:'Crea workspace',exact:true}).click();
    await page.locator('.cloud-members').waitFor();assert.equal(await page.locator('.cloud-main h1 img').count(),0);
    assert((await page.locator('.cloud-members').innerText()).includes('Proprietario'));
    await page.evaluate(()=>render());assert.equal(await page.locator('.lab-shell').count(),0);
    await page.keyboard.press('Control+k');assert.equal(await page.locator('#command-input').count(),0);
    assert.equal(await page.evaluate(()=>localStorage.getItem('spps-lab-assistant-data')),'[]');
    await page.setViewportSize({width:390,height:844});assert(await page.evaluate(()=>document.documentElement.scrollWidth<=innerWidth+1));
    await page.evaluate(()=>window.failLogout=true);await page.getByRole('button',{name:'Esci dall’account'}).click();
    await page.getByRole('alert').filter({hasText:'Operazione non riuscita'}).waitFor();
    await page.evaluate(()=>window.failLogout=false);await page.getByRole('button',{name:'Esci dall’account'}).click();
    await page.locator('#cloud-email').waitFor();assert.equal(await page.locator('.cloud-members').count(),0);
    await page.reload();await page.locator('#cloud-email').waitFor();
    await page.getByRole('button',{name:'Modalità locale',exact:true}).click();await page.locator('.lab-shell').waitFor();
    assert.deepEqual(errors,[]);
    console.log('PASS: local/cloud separation, OTP UI, escaped workspace names, roles, render guard, mobile overflow, logout failure/retry, local data preservation. SDK mocked; live Auth not tested.');
  } finally {if(browser)await browser.close();await new Promise(resolve=>server.close(resolve));}
})().catch(error=>{console.error(error);process.exitCode=1;});
