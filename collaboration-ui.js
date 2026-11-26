/* Phase 1 is deliberately a separate account/workspace screen, not a cloud
 * facade over local scientific data. No legacy collections are uploaded. */
window.SPPSCloud = (() => {
  let mode = 'local', state = 'loading', user = null, workspaces = [], members = [], activeId = null;
  let email = '', epoch = 0, requestId = null, requestName = '', errorMessage = '';
  let busy = false, unsubscribe;
  const escape = value => escapeHtml(String(value ?? ''));
  const roles = { owner: 'Proprietario', admin: 'Amministratore', editor: 'Editor', viewer: 'Sola lettura' };
  const input = (id, label, type, autocomplete) => `<label for="${id}">${label}</label><input class="form-input" id="${id}" type="${type}" autocomplete="${autocomplete}" required>`;
  function message(error) {
    if (error?.message === 'cloud_public_key_required') return 'Usa esclusivamente una chiave pubblica Supabase. Le chiavi segrete non sono ammesse.';
    if (/cloud_configuration/.test(error?.message || '')) return 'Configurazione cloud incompleta. Verifica URL e chiave pubblica.';
    if (error?.status === 429 || error?.code === 'over_email_send_rate_limit') return 'Troppe richieste. Attendi prima di richiedere un altro codice.';
    if (error?.code === 'otp_expired') return 'Il codice non è valido o è scaduto. Richiedine uno nuovo.';
    if (error?.code === '42501') return 'Non hai il permesso di eseguire questa operazione.';
    if (error?.code === 'PGRST202' || error?.code === '42P01') return 'Il database cloud non è ancora configurato. Applica la migrazione della fase 1.';
    if (!navigator.onLine) return 'Sei offline. I dati locali restano disponibili.';
    return 'Operazione non riuscita. Controlla la connessione e riprova.';
  }
  function show() {
    if (mode !== 'cloud') return;
    document.body.classList.remove('lab-mode');
    document.body.dataset.page = 'cloud';
    const workspace = workspaces.find(w => w.id === activeId);
    let content = '';
    if (state === 'loading') content = '<p role="status">Connessione al workspace…</p>';
    else if (state === 'unconfigured') content = '<h1>Workspace cloud</h1><p>Cloud non attivo. Configura il progetto Supabase per accedere con il tuo account.</p><p class="muted">Nessun dato locale è stato trasferito.</p>';
    else if (state === 'error') content = '<h1>Connessione non disponibile</h1><button class="ui-button primary" data-cloud="retry">Riprova</button>';
    else if (!user) {
      content = `<div class="eyebrow">ACCOUNT MODILABS</div><h1>Accedi al laboratorio</h1><form id="cloud-login-form">${input('cloud-email','Email','email','email')}<button class="ui-button primary" type="submit" ${busy?'disabled':''}>Invia codice di accesso</button></form>${state === 'code-sent' ? `<form id="cloud-code-form"><p class="muted">Codice inviato a ${escape(email)}</p>${input('cloud-code','Codice ricevuto via email','text','one-time-code')}<button class="ui-button primary" type="submit" ${busy?'disabled':''}>Accedi</button></form>`:''}`;
    } else {
      content = `<div class="eyebrow">WORKSPACE CLOUD / FASE 1</div><h1>${workspace ? escape(workspace.name) : 'I tuoi workspace'}</h1><p class="muted">${escape(user.email || '')}</p>
        ${workspaces.length ? `<label for="cloud-workspace">Workspace</label><select class="form-select" id="cloud-workspace">${workspaces.map(w=>`<option value="${escape(w.id)}" ${w.id===activeId?'selected':''}>${escape(w.name)} · ${roles[w.role]}</option>`).join('')}</select>` : '<p>Non appartieni ancora a un workspace.</p>'}
        ${workspace ? `<section class="cloud-members"><h2>Membri</h2><ul>${members.map(m=>`<li><span>${escape(m.display_name)}${m.user_id===user.id?' (tu)':''}</span><span class="badge">${roles[m.role]}</span></li>`).join('') || '<li>Nessun membro caricato</li>'}</ul></section>` : ''}
        <section class="cloud-create"><h2>Crea workspace</h2><form id="cloud-create-form">${input('cloud-workspace-name','Nome del laboratorio','text','organization')}<button class="ui-button primary" type="submit" ${busy?'disabled':''}>Crea workspace</button></form></section>
        <p class="cloud-phase-notice">Account e ruoli sono attivi. Condivisione delle sintesi, inviti e migrazione dei dati saranno abilitati nelle fasi successive. Le schede locali non appartengono a questo workspace.</p>`;
    }
    document.getElementById('app').innerHTML = `<div class="cloud-shell"><header><button class="cloud-logo" data-cloud="local" aria-label="MoDiLabs, modalità locale"><img src="logo_transparent.png" alt="MoDiLabs"></button><div><button class="ui-button" data-cloud="local">Modalità locale</button>${user?'<button class="ui-button" data-cloud="logout">Esci dall’account</button>':''}</div></header><main class="cloud-main" aria-busy="${busy}">${content}<div id="cloud-error" role="alert">${escape(errorMessage)}</div></main></div>`;
    const emailInput = document.getElementById('cloud-email');
    if (emailInput) emailInput.value = email;
    const nameInput = document.getElementById('cloud-workspace-name');
    if (nameInput) { nameInput.maxLength = 120; nameInput.value = requestName; }
    const codeInput = document.getElementById('cloud-code');
    if (codeInput) { codeInput.inputMode = 'numeric'; codeInput.pattern = '[0-9]{6,10}'; codeInput.maxLength = 10; }
    document.getElementById('cloud-login-form')?.addEventListener('submit', async event => {
      event.preventDefault(); email = emailInput.value.trim();
      await perform(async token => { await SPPSAuth.sendCode(email); if(token===epoch){state='code-sent';show();document.getElementById('cloud-code')?.focus();} });
    });
    document.getElementById('cloud-code-form')?.addEventListener('submit', async event => {
      event.preventDefault(); const code = codeInput.value.trim();
      await perform(async token => { await SPPSAuth.verifyCode(email,code); if(token===epoch)await refresh(); });
    });
    document.getElementById('cloud-create-form')?.addEventListener('submit', async event => {
      event.preventDefault(); const name = nameInput.value.trim(); if (!name) return;
      if (name !== requestName || !requestId) { requestId=crypto.randomUUID(); requestName=name; }
      await perform(async token => {
        const id = await SPPSWorkspaces.create(name,requestId);
        if(token!==epoch)return;
        SPPSWorkspaces.remember(user.id,id);requestId=null;requestName='';await refresh();
      });
    });
    document.getElementById('cloud-workspace')?.addEventListener('change', async event => {
      const id=event.target.value;
      if(!workspaces.some(w=>w.id===id))return;
      activeId=id;members=[];SPPSWorkspaces.remember(user.id,id);
      await refresh();
    });
  }
  async function perform(task) {
    if(busy)return;busy=true;errorMessage='';const token=epoch;
    document.querySelectorAll('.cloud-main button').forEach(button=>button.disabled=true);
    try { await task(token); }
    catch(error) { if(token===epoch) {errorMessage=message(error);show();} }
    finally { if(token===epoch) {busy=false;document.querySelectorAll('.cloud-main button').forEach(button=>button.disabled=false);} }
  }
  async function refresh() {
    const token=++epoch;busy=false;user=SPPSAuth.current();workspaces=[];members=[];activeId=null;
    errorMessage='';state=user?'loading':'signed-out';show();
    if(!user || mode!=='cloud')return;
    const userId=user.id;
    try {
      const rows=await SPPSWorkspaces.list(userId);if(token!==epoch)return;
      workspaces=rows;
      const saved=SPPSWorkspaces.last(userId);
      activeId=rows.find(w=>w.id===saved)?.id||rows[0]?.id||null;
      if(activeId){const loaded=await SPPSWorkspaces.members(activeId);if(token!==epoch)return;members=loaded;SPPSWorkspaces.remember(userId,activeId);}
      state='ready';show();
    } catch(error) {if(token===epoch){workspaces=[];members=[];activeId=null;state='error';errorMessage=message(error);show();}}
  }
  async function open() {
    mode='cloud';epoch++;busy=false;errorMessage='';state=SPPSBackend.configured()?'loading':'unconfigured';show();
    if(!SPPSBackend.configured())return;
    if(!unsubscribe)unsubscribe=SPPSAuth.subscribe(()=>{
      // Clear synchronously, fetch outside the SDK's auth callback.
      epoch++;workspaces=[];members=[];activeId=null;user=SPPSAuth.current();busy=false;
      state=user?'loading':'signed-out';show();
      if(mode==='cloud')setTimeout(()=>refresh(),0);
    });
    const token=epoch;
    try {await SPPSAuth.start();if(mode==='cloud' && token===epoch)await refresh();}
    catch(error){if(mode==='cloud' && token===epoch){state='error';errorMessage=message(error);show();}}
  }
  function local() {
    mode='local';epoch++;busy=false;workspaces=[];members=[];activeId=null;errorMessage='';
    document.body.dataset.page='home';currentTab='home';navigate('dashboard');
    if(SPPSBackend.configured())initInventory().then(()=>render()).catch(()=>LabUI.toast('Connessione locale non disponibile.'));
  }
  function protectsRender() { if(mode==='cloud'){show();return true;}return false; }
  document.addEventListener('click',event=>{
    const action=event.target.closest('[data-cloud]')?.dataset.cloud;
    if(action==='open'||action==='retry')open();
    else if(action==='local')local();
    else if(action==='logout')perform(async()=>{await SPPSAuth.signOut();});
  });
  document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='visible' && mode==='cloud' && user)refresh();});
  return Object.freeze({ open, local, protectsRender, isActive:()=>mode==='cloud' });
})();
