/* Presentation layer. Chemistry and persistence remain owned by app.js. */
const LabUI = (() => {
  const $ = (s, root = document) => root.querySelector(s);
  const esc = value => escapeHtml(String(value ?? ''));
  const read = (key, fallback) => { try { return JSON.parse(localStorage.getItem(key)) ?? fallback; } catch { return fallback; } };
  const prefs = read('spps-ui-preferences-v1', { collapsed: false });
  let section = 'overview', detailId = null, order = 'nc', lab = false;
  let commandItems = [], commandIndex = 0, modalObserver, headerObserver;
  const icon = name => {
    const node = window.LAB_ICONS[name] || window.LAB_ICONS.FlaskConical;
    return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round" class="ui-icon" aria-hidden="true">${node[2].map(([tag, attrs]) => `<${tag} ${Object.entries(attrs).filter(([k]) => k !== 'key').map(([k,v]) => `${k}="${esc(v)}"`).join(' ')}></${tag}>`).join('')}</svg>`;
  };
  const button = (action, label, symbol, cls = '') => `<button type="button" class="ui-button ${cls}" data-ui-action="${action}" title="${esc(label)}">${icon(symbol)}<span>${esc(label)}</span></button>`;
  function destination(where) {
    lab = false;
    if (['home','files','folders'].includes(where)) { currentTab = where; navigate('dashboard'); }
    else navigate(where);
  }
  function shell(content) {
    const active = currentView === 'dashboard' ? currentTab : currentView;
    const title = {home:'Panoramica laboratorio',files:'Sintesi',folders:'Progetti',inventory:'Inventario',research:'Letteratura',settings:'Impostazioni',folder:'Progetto',detail:'Scheda di sintesi'}[active] || 'Laboratorio';
    const nav = (id,label,name) => `<button class="nav-item ${active === id || (id==='files' && active==='detail') ? 'selected' : ''}" data-ui-action="nav:${id}" ${active===id?'aria-current="page"':''} title="${label}">${icon(name)}<span>${label}</span></button>`;
    return `<div class="lab-shell ${prefs.collapsed?'sidebar-collapsed':''}"><aside class="lab-sidebar" aria-label="Navigazione principale"><button class="lab-brand" data-ui-action="nav:home" aria-label="MoDiLabs, home"><img src="logo_transparent.png" alt="MoDiLabs"><span>SPPS Lab Assistant</span></button><nav>${nav('home','Home','House')}<div class="nav-group-label">Workspace</div>${nav('files','Sintesi','FlaskConical')}${nav('folders','Progetti','Folder')}${nav('inventory','Inventario','Package')}<div class="nav-group-label">Ricerca e laboratorio</div>${nav('research','Letteratura','BookOpen')}<button class="nav-item" data-ui-action="lab" title="Modalità banco">${icon('Beaker')}<span>Modalità banco</span></button></nav><div class="sidebar-bottom">${nav('settings','Impostazioni e backup','Settings')}${button('collapse','Comprimi navigazione',prefs.collapsed?'PanelLeftOpen':'PanelLeftClose','collapse-control')}</div></aside><div class="lab-main"><header class="lab-topbar"><span class="topbar-location">${esc(title)}</span><div class="topbar-actions">${button('search','Cerca','Search','search-trigger')}${button('theme','Cambia tema','Sun','icon-only')}${button('copilot','Lab Copilot','Sparkles','icon-only')}${button('new','Nuova sintesi','Plus','primary')}</div></header><main id="workspace-content">${content}</main></div></div>`;
  }
  function connectionRows() {
    return `<div class="connection-list"><div>${icon('Package')}<div><strong>Inventario Excel</strong><small>${esc(inventoryFileName || 'Non collegato')}</small></div>${button('connect-inventory',inventoryFileName?'Sincronizza':'Collega',inventoryFileName?'RefreshCw':'Link')}</div><div>${icon('Folder')}<div><strong>Area condivisa</strong><small>${esc(workspaceName || 'Non collegata')}</small></div>${button('connect-workspace',workspaceName?'Riconnetti':'Collega','Link')}</div><span id="workspace-sync-status" class="notes-saved" role="status">Dati sincronizzati</span></div>`;
  }
  function dashboard() {
    let content;
    if (currentTab==='folders') content=renderFoldersTab();
    else if (currentTab==='files') content=renderAllFilesTab();
    else {
      const recent=[...syntheses].sort((a,b)=>String(b.createdAt||b.dateStarted).localeCompare(String(a.createdAt||a.dateStarted)));
      const active=recent.filter(s=>s.status==='in-progress');
      content=`<div class="home-heading"><div class="eyebrow">MODILABS / WORKSPACE</div><h1>Il tuo laboratorio, oggi.</h1><p>${active.length} sintesi in corso · ${folders.length} progetti · ${syntheses.filter(s=>s.status==='analyzed').length} sintesi analizzate</p></div><div class="home-columns"><section class="home-primary"><div class="section-heading-row"><h2>In corso</h2>${button('nav:files','Tutte le sintesi','ArrowRight')}</div><div class="active-syntheses">${active.slice(0,3).map(s=>{const t=tokenizeSequence(s.sequence),done=countCompletedSequenceResidues(s,t),f=folders.find(f=>f.id===s.folderId);return `<article class="active-synthesis"><div class="active-synthesis-heading"><button data-ui-action="open:${esc(s.id)}"><h3>${esc(s.name)}</h3></button>${renderStatusBadge(s.status)}</div><div class="sequence-preview">${esc(s.sequence)}</div><div class="active-meta">${esc(s.resinType)} · ${s.scale} mmol · ${t.length} residui</div><div class="progress-line"><progress max="${t.length||1}" value="${done}" aria-label="Avanzamento ${esc(s.name)}"></progress><span>${done}/${t.length}</span></div><footer><span>${esc(f?.name || 'Senza progetto')}</span>${button('continue:'+s.id,'Continua','ArrowRight')}</footer></article>`;}).join('') || `<div class="quiet-empty">${icon('FlaskConical')}<h3>Nessuna sintesi in corso</h3>${button('new','Nuova sintesi','Plus','primary')}</div>`}</div><div class="section-heading-row"><h2>Sintesi recenti</h2><span>${syntheses.length} totali</span></div>${recent.length?renderTable(recent.slice(0,6),true):'<p class="muted">Le tue sintesi saranno raccolte qui.</p>'}</section><aside class="home-secondary"><h2>Connessioni</h2>${connectionRows()}<h2>Da verificare</h2><div class="attention-list">${!inventoryFileName?button('connect-inventory','Collega l’inventario','CircleAlert'):''}${!workspaceName?button('connect-workspace','Collega l’area condivisa','Folder'):''}${recent.filter(s=>s.status==='cleaved'&&!s.msObserved).slice(0,3).map(s=>button('analysis:'+s.id,'MS mancante · '+s.name,'Activity')).join('')}${inventoryFileName&&workspaceName&&!recent.some(s=>s.status==='cleaved'&&!s.msObserved)?'<p class="muted">Nessuna verifica pendente.</p>':''}</div><div class="section-heading-row"><h2>Progetti</h2>${button('new-project','Nuovo progetto','Plus','icon-only')}</div>${folders.slice(0,6).map(f=>`<button class="project-shortcut" data-ui-action="folder:${esc(f.id)}">${icon('Folder')}<span>${esc(f.name)}</span><small>${syntheses.filter(s=>s.folderId===f.id).length}</small></button>`).join('') || '<p class="muted">Nessun progetto creato.</p>'}</aside></div>`;
    }
    return content+renderNewFolderModal()+renderNewSynthesisModal(null);
  }
  function inventoryRows(query='') {
    return inventoryData.filter(row=>row.join(' ').toLowerCase().includes(query.toLowerCase())).map(row=>{const index=inventoryData.indexOf(row);return `<tr tabindex="0" data-inventory-index="${index}">${Array.from({length:7},(_,i)=>`<td>${esc(row[i] || 'Non disponibile')}</td>`).join('')}</tr>`;}).join('');
  }
  function auxiliary() {
    if(currentView==='inventory') return `<div class="page-heading"><div><h1>Inventario</h1><p>${esc(inventoryFileName || 'Nessun file Excel collegato')} · ${inventoryData.length} righe</p></div>${button('connect-inventory',inventoryFileName?'Sincronizza':'Collega Excel','RefreshCw','primary')}</div><label class="inventory-search-label" for="inventory-search">Composto, CAS, codice o locazione</label><input id="inventory-search" class="form-input" type="search" placeholder="Cerca nell’inventario"><div class="inventory-table-wrap"><table class="calc-table"><thead><tr>${['Composto','Quantità (conf.)','Marca','Codice','Locazione','CAS','Residuo'].map(t=>`<th>${t}</th>`).join('')}</tr></thead><tbody id="global-inventory-rows">${inventoryRows()}</tbody></table></div>${!inventoryData.length?'<div class="quiet-empty"><h2>Inventario non collegato</h2><p>Collega il file Excel per consultare composti e locazioni.</p></div>':''}`;
    if(currentView==='research') return `<div class="page-heading"><div><h1>Letteratura</h1><p>PubMed · Europe PMC · OpenAlex</p></div></div>${renderLiteratureTipsSection(researchContext(),tokenizeSequence(researchContext().sequence),analyzeCyclizations(researchContext().sequence),suggestCleavageCocktail(tokenizeSequence(researchContext().sequence)),getAvailableActivators(researchContext()))}`;
    return `<div class="page-heading"><h1>Impostazioni</h1></div><section class="settings-section"><h2>Aspetto</h2>${button('theme','Cambia tema','Sun')}<p class="muted">Tema attuale: <span id="theme-description">${getCurrentTheme()==='dark'?'Scuro':'Chiaro'}</span></p></section><section class="settings-section"><h2>Lab Copilot</h2><label for="settings-ai">Endpoint del proxy AI</label><div class="settings-input-row"><input type="url" id="settings-ai" class="form-input" value="${esc(loadAiChatEndpoint())}" placeholder="http://localhost:8787/chat">${button('save-ai','Salva','Save','primary')}</div><p class="muted">Senza endpoint sono disponibili solo le risposte locali. Le chiavi API restano nel proxy.</p></section><section class="settings-section"><h2>Connessioni</h2>${connectionRows()}</section><section class="settings-section"><h2>Backup dei dati</h2><div class="button-row">${button('export','Esporta JSON','Download')}${button('import','Importa JSON','Upload')}</div></section>`;
  }
  function researchContext() { return syntheses.find(s=>s.id===currentSynthesisId) || syntheses[0] || {id:'global-research',name:'SPPS',sequence:'',resinType:'',activators:[],customMWs:{}}; }
  function toast(message,kind='info') {
    let host=$('#ui-toasts'); if(!host){host=document.createElement('div');host.id='ui-toasts';host.setAttribute('role','status');document.body.append(host);}
    const item=document.createElement('div'); item.className='ui-toast '+kind; item.textContent=message; host.append(item);setTimeout(()=>item.remove(),kind==='error'?9000:4500);
  }
  function dialog(content, cls='') {
    const previous=document.activeElement, d=document.createElement('dialog');d.className='ui-dialog '+cls;d.innerHTML=content;document.body.append(d);d.addEventListener('close',()=>{d.remove();if(previous?.isConnected)previous.focus();});d.addEventListener('click',e=>{if(e.target===d)d.close();});d.showModal();return d;
  }
  function confirm(message) {return new Promise(resolve=>{const d=dialog(`<h2>Conferma operazione</h2><p>${esc(message)}</p><div class="dialog-actions"><button class="ui-button" data-cancel>Annulla</button><button class="ui-button danger" data-confirm>Conferma</button></div>`);let ok=false;$('[data-cancel]',d).onclick=()=>d.close();$('[data-confirm]',d).onclick=()=>{ok=true;d.close();};d.addEventListener('close',()=>resolve(ok));$('[data-cancel]',d).focus();});}
  function previewFile(file) {
    if(!/^data:(image\/|application\/pdf)/.test(file.dataUrl))return;
    const d=dialog(`<header class="section-heading-row"><h2>${esc(file.name)}</h2><button class="ui-button icon-only" aria-label="Chiudi">${icon('X')}</button></header>${file.dataUrl.startsWith('data:image/')?`<img class="file-preview-image" src="${esc(file.dataUrl)}" alt="${esc(file.name)}">`:`<p>Documento PDF</p>`}<a class="ui-button" href="${esc(file.dataUrl)}" download="${esc(file.name)}">${icon('Download')} Scarica allegato</a>`,'file-preview-dialog');$('header button',d).onclick=()=>d.close();
  }
  function palette() {
    if($('.command-palette')) return;
    const d=dialog(`<div class="command-search">${icon('Search')}<input autofocus id="command-input" role="combobox" aria-label="Cerca o esegui un comando" aria-expanded="true" aria-controls="command-results" autocomplete="off" placeholder="Cerca sintesi, sequenze, progetti, CAS…"><kbd>Esc</kbd></div><div id="command-results" role="listbox"></div><footer>↑ ↓ Seleziona <span>Invio Apri</span></footer>`,'command-palette');
    const input=$('input',d),list=$('#command-results',d);
    const update=()=>{
      const q=input.value.trim().toLowerCase();
      const recent=read('spps-ui-recent-v1',[]);
      commandItems=[...syntheses.map(s=>({group:recent.includes(s.id)?'Recenti':'Sintesi',label:s.name,meta:s.sequence,action:'open:'+s.id})),...folders.map(f=>({group:'Progetti',label:f.name,meta:'Progetto',action:'folder:'+f.id})),...inventoryData.map((r,i)=>({group:'Inventario',label:String(r[0]||''),meta:String(r[5]||'')+' '+String(r[4]||''),action:'inventory-item:'+i})),... [['new','Nuova sintesi'],['new-project','Nuovo progetto'],['nav:inventory','Apri inventario'],['nav:settings','Impostazioni e backup'],['nav:research','Apri letteratura'],['lab','Modalità banco'],['connect-inventory','Collega o sincronizza inventario'],['export','Esporta backup JSON'],['import','Importa JSON'],['theme','Cambia tema'],['copilot','Apri Lab Copilot'],['nav:home','Torna alla home']].map(([action,label])=>({group:'Azioni',action,label,meta:''}))].filter(x=>(x.label+' '+x.meta).toLowerCase().includes(q)).sort((a,b)=>a.group.localeCompare(b.group)).slice(0,60);
      commandIndex=0; list.innerHTML=commandItems.map((item,i)=>`${i===0||item.group!==commandItems[i-1].group?`<div class="command-group">${item.group}</div>`:''}<button role="option" id="command-option-${i}" aria-selected="${i===0}" data-command-index="${i}"><span>${esc(item.label)}</span><small>${esc(item.meta)}</small></button>`).join('')||'<p class="quiet-empty">Nessun risultato</p>';highlight();
    };
    const highlight=()=>{list.querySelectorAll('[role=option]').forEach((el,i)=>el.setAttribute('aria-selected',i===commandIndex));input.setAttribute('aria-activedescendant','command-option-'+commandIndex);$('#command-option-'+commandIndex,d)?.scrollIntoView({block:'nearest'});};
    const execute=i=>{const item=commandItems[i];if(item){d.close();action(item.action);}};
    input.oninput=update;input.onkeydown=e=>{if(e.key==='ArrowDown'||e.key==='ArrowUp'){e.preventDefault();commandIndex=(commandIndex+(e.key==='ArrowDown'?1:-1)+commandItems.length)%Math.max(1,commandItems.length);highlight();}if(e.key==='Enter'){e.preventDefault();execute(commandIndex);}};
    list.onclick=e=>{const el=e.target.closest('[data-command-index]');if(el)execute(+el.dataset.commandIndex);};update();input.focus();
  }
  function nextIndex(s) {const tokens=tokenizeSequence(s.sequence),p=ensureSequenceProgress(s,tokens);for(let i=tokens.length-1;i>=0;i--)if(!p.completed[i])return i;return -1;}
  function step(s,index) {
    const tokens=tokenizeSequence(s.sequence), token=tokens[index], reverse=tokens.length-1-index;
    const aa=Object.assign({},AMINO_ACIDS[token]||{code3:token,fmocName:token,fmocMW:Number(s.customMWs?.[token])},getCyclizationProtectedOverrides(s.sequence)[index]||{});
    const preloaded=isPreloadedWangStep(s,token,index,tokens.length), available=getAvailableActivators(s);
    const act=ACTIVATORS.find(a=>a.name===(s.aaActivators?.[reverse]||available[0]?.name))||available[0]||ACTIVATORS[0];
    const inv=findInInventory(aa.fmocName),mass=preloaded?0:calculateAAMass(token,s.equivalents,s.scale,aa.fmocMW);
    return {tokens,token,index,reverse,aa,preloaded,act,inv,mass,actMass:calculateActivatorMass(act,s.scale),done:!!ensureSequenceProgress(s,tokens).completed[index]};
  }
  function stepContent(s,index) {
    const x=step(s,index),location=x.inv ? formatInventoryLocation(x.inv) : '';
    return `<div class="eyebrow">POSIZIONE N→C ${index+1} · STEP C→N ${x.reverse+1}</div><h2 class="residue-title">${esc(x.aa.code3)}</h2><p class="residue-full-name">${esc(x.aa.fmocName)}</p><div class="step-state">${x.preloaded?'Precaricato sulla resina':x.done?'Completato':'Da completare'}</div><dl class="step-values"><div><dt>Massa AA</dt><dd>${Number.isFinite(x.aa.fmocMW)?x.mass.toFixed(2)+' <small>mg</small>':'Non disponibile'}</dd></div><div><dt>Equivalenti AA</dt><dd>${x.preloaded?'—':s.equivalents}</dd></div><div><dt>Protezione</dt><dd>${esc(getAminoAcidProtectionLabel(x.aa,x.aa.fmocName))}</dd></div><div><dt>Scala</dt><dd>${s.scale} <small>mmol</small></dd></div></dl>${!x.preloaded?`<div class="step-reagents"><h3>${esc(x.act.name)}</h3><div>${x.actMass.activatorMass.toFixed(2)} mg · ${x.act.defaultEq} eq</div>${x.act.base?`<p>${esc(x.act.base)} · ${x.actMass.baseMass.toFixed(2)} mg · ${x.act.baseEq} eq</p>`:''}${x.act.coReagent?`<p>${esc(x.act.coReagent)} · ${x.actMass.coReagentMass.toFixed(2)} mg</p>`:''}</div>`:''}<div class="step-inventory"><h3>${icon('MapPin')} Inventario</h3><p>${!inventoryFileName?'Inventario non collegato':!x.inv?'Composto non trovato':esc(x.inv.name || x.aa.fmocName)}</p>${x.inv?`<strong>${esc(location || 'Locazione non disponibile')}</strong>`:''}</div><button class="ui-button primary complete-residue" data-complete-index="${index}">${icon(x.done?'ArrowLeft':'Check')}${x.done?'Segna da completare':'Segna come completato'}</button>`;
  }
  function inspectResidue(s,index) {
    document.querySelector('.residue-inspector')?.close();
    const d=dialog(`<header><h2>Dettaglio residuo</h2><button class="ui-button icon-only" aria-label="Chiudi">${icon('X')}</button></header>${stepContent(s,index)}`,'residue-inspector');$('header button',d).onclick=()=>d.close();$('[data-complete-index]',d).onclick=()=>{toggleSequenceResidue(s,index);d.close();toast('Avanzamento aggiornato');};
  }
  function inspectInventory(index) {const row=inventoryData[index];if(!row)return;const d=dialog(`<header><h2>Composto in inventario</h2><button class="ui-button icon-only" aria-label="Chiudi">${icon('X')}</button></header><h3>${esc(row[0])}</h3><dl class="inventory-facts">${['Composto','Quantità (conf.)','Marca','Codice','Locazione','CAS','Residuo'].map((label,i)=>`<div><dt>${label}</dt><dd>${esc(row[i]||'Non disponibile')}</dd></div>`).join('')}</dl>`,'residue-inspector');$('header button',d).onclick=()=>d.close();}
  function labContent(s) {const index=nextIndex(s),tokens=tokenizeSequence(s.sequence);return `<div class="lab-mode-heading"><div><div class="eyebrow">MODALITÀ BANCO / C→N</div><h1>${esc(s.name)}</h1><p>${countCompletedSequenceResidues(s,tokens)} / ${tokens.length} residui completati</p></div>${button('exit-lab','Esci dalla modalità banco','LogOut')}</div><div class="lab-step" aria-live="polite">${index<0?`<div class="quiet-empty">${icon('CircleCheck')}<h2>Sequenza completata</h2>${button('exit-lab','Apri scheda','ArrowRight')}</div>`:stepContent(s,index)}</div>`;}
  function showSection(value) {
    section=value;document.querySelectorAll('[data-detail-section]').forEach(el=>el.hidden=!el.dataset.detailSection.split(' ').includes(value));
    document.querySelectorAll('[data-detail-tab]').forEach(el=>{el.classList.toggle('active',el.dataset.detailTab===value);el.setAttribute('aria-selected',el.dataset.detailTab===value);});
    if(value==='synthesis'||value==='reagents')document.querySelector(`.calc-inventory-tab[data-panel-tab="${value==='reagents'?'inventory':'calculations'}"]`)?.click();
  }
  function updateProgress(s) {
    const next=nextIndex(s);document.querySelectorAll('[data-sequence-index]').forEach(el=>el.classList.toggle('current',+el.dataset.sequenceIndex===next));
    document.querySelectorAll('[data-step-index]').forEach(el=>el.classList.toggle('current-step',+el.dataset.stepIndex===next));
    const tokens=tokenizeSequence(s.sequence),count=countCompletedSequenceResidues(s,tokens);
    if($('#header-progress'))$('#header-progress').textContent=count+' / '+tokens.length+' residui';
    if(lab && $('#lab-mode-view')){$('#lab-mode-view').innerHTML=labContent(s);bindLab(s);}
  }
  function bindLab(s) {$('[data-complete-index]',$('#lab-mode-view'))?.addEventListener('click',e=>{toggleSequenceResidue(s,+e.currentTarget.dataset.completeIndex);$('#lab-mode-view .complete-residue')?.focus();});}
  function mountDetail(s) {
    if(detailId!==s.id){detailId=s.id;section='overview';lab=false;}
    const recent=read('spps-ui-recent-v1',[]).filter(id=>id!==s.id);try{localStorage.setItem('spps-ui-recent-v1',JSON.stringify([s.id,...recent].slice(0,8)));}catch{}
    const header=$('.detail-header'),nav=document.createElement('nav');nav.className='detail-tabs';nav.setAttribute('role','tablist');nav.setAttribute('aria-label','Sezioni sintesi');
    $('.topbar-location').textContent=s.name;
    headerObserver=new ResizeObserver(()=>document.documentElement.style.setProperty('--detail-header-height',header.offsetHeight+'px'));headerObserver.observe(header);
    nav.innerHTML=[['overview','Panoramica'],['synthesis','Sintesi'],['reagents','Reagenti'],['analysis','Analisi'],['literature','Letteratura'],['notes','Note e file']].map(([id,label])=>`<button role="tab" data-detail-tab="${id}">${label}</button>`).join('');header.after(nav);nav.onclick=e=>{const t=e.target.closest('[data-detail-tab]');if(t)showSection(t.dataset.detailTab);};
    nav.onkeydown=e=>{if(!['ArrowLeft','ArrowRight','Home','End'].includes(e.key))return;e.preventDefault();const tabs=[...nav.children],idx=tabs.indexOf(document.activeElement),next=e.key==='Home'?0:e.key==='End'?tabs.length-1:(idx+(e.key==='ArrowRight'?1:-1)+tabs.length)%tabs.length;tabs[next].focus();tabs[next].click();};
    const children=[...$('#workspace-content').children];children.forEach(el=>{if(el.matches('.info-grid'))el.dataset.detailSection='overview';else if(el.matches('.sequence-progress-card'))el.dataset.detailSection='overview synthesis';else if(el.matches('.calc-inventory-card'))el.dataset.detailSection='synthesis reagents';else if(el.matches('.literature-section'))el.dataset.detailSection='literature';else if(el.querySelector('#ms-observed,#hplc-purity'))el.dataset.detailSection='analysis';else if(el.querySelector('#notes-textarea'))el.dataset.detailSection='notes';else if(el.matches('.section-card'))el.dataset.detailSection='synthesis';});
    const resin=RESINS.find(r=>r.name===s.resinType),mw=calculatePeptideMW(tokenizeSequence(s.sequence),resin?.type||'amide',s.customMWs,s.sequence);
    $('.detail-header-left').insertAdjacentHTML('beforeend',`<div class="header-progress-line"><span id="header-progress"></span><span>${esc(s.resinType)} · ${s.scale} mmol</span></div>`);
    $('.info-grid').insertAdjacentHTML('afterbegin',`<div class="info-item"><div class="info-item-label">PM atteso · calcolato</div><div class="info-item-value mono">${mw.toFixed(2)} Da</div></div>`);
    $('.info-grid').insertAdjacentHTML('beforeend',`<div class="info-item"><div class="info-item-label">MS osservato · registrato</div><div class="info-item-value mono">${esc(s.msObserved || 'Non registrato')}</div></div><div class="info-item"><div class="info-item-label">Purezza HPLC · registrata</div><div class="info-item-value mono">${s.hplcPurity?esc(s.hplcPurity)+'%':'Non registrata'}</div></div>`);
    if($('.preloaded-warning'))$('.preloaded-warning').dataset.detailSection='overview synthesis';
    const sequence=$('.sequence-progress-header');sequence?.insertAdjacentHTML('beforeend',`<div class="sequence-direction" role="group" aria-label="Direzione sequenza"><button data-order="nc" aria-pressed="${order==='nc'}">N→C</button><button data-order="cn" aria-pressed="${order==='cn'}">C→N</button></div>`);
    const setOrder=()=>{
      const strip=$('.sequence-progress-strip');
      if(strip){[...strip.children].sort((a,b)=>(Number(a.dataset.sequenceIndex)-Number(b.dataset.sequenceIndex))*(order==='cn'?-1:1)).forEach(el=>strip.append(el));strip.scrollLeft=0;}
      document.querySelectorAll('[data-order]').forEach(el=>el.setAttribute('aria-pressed',el.dataset.order===order));
    };
    document.querySelectorAll('[data-order]').forEach(el=>el.onclick=()=>{order=el.dataset.order;setOrder();});setOrder();
    const overrides=getCyclizationProtectedOverrides(s.sequence), cycles=analyzeCyclizations(s.sequence).cyclizations,tokens=tokenizeSequence(s.sequence);
    document.querySelectorAll('[data-sequence-index]').forEach(el=>{const i=+el.dataset.sequenceIndex,aa=overrides[i]||AMINO_ACIDS[tokens[i]];el.classList.toggle('nonstandard',!AMINO_ACIDS[tokens[i]]);const protection=getAminoAcidProtectionLabel(aa,tokens[i]);el.insertAdjacentHTML('beforeend',`<span class="rail-protection">${esc(protection==='Nessuna'?'':protection)}</span>`);if(cycles.some(c=>i>=c.segmentStartTokenIndex&&i<=c.segmentEndTokenIndex))el.classList.add('cyclized');});
    const more=document.createElement('details');more.className='ui-menu';more.innerHTML=`<summary title="Altre azioni" aria-label="Altre azioni">${icon('MoreHorizontal')}</summary><div></div>`;$('#btn-delete').before(more);$('div',more).append($('#btn-delete'));$('#btn-delete').innerHTML=icon('Trash2')+' Elimina sintesi';$('#btn-delete').setAttribute('aria-label','Elimina sintesi');
    $('.detail-header-actions').insertAdjacentHTML('afterbegin',button('lab','Modalità banco','Beaker','primary'));
    $('#btn-edit-sequence').hidden=true;$('#btn-edit-params').innerHTML=icon('Pencil')+' Modifica';$('#btn-print-lab').innerHTML=icon('Printer')+' Stampa';
    const notes=$('#notes-textarea').parentElement;notes.insertAdjacentHTML('beforeend',`<div class="notes-files"><h3>Allegati analitici</h3><p>${(attachments[s.id]?.ms?.length||0)} spettri MS · ${(attachments[s.id]?.hplc?.length||0)} cromatogrammi HPLC</p>${button('detail:analysis','Apri analisi e allegati','Activity')}</div>`);
    $('.chatbot-title').textContent='Lab Copilot';$('.chatbot-subtitle').textContent=`${s.name} · ${tokens.length} residui · ${s.scale} mmol`;
    $('#chatbot-config-panel')?.remove();$('#chatbot-config-toggle')?.remove();$('.chatbot-message.bot > div:last-child').textContent=loadAiChatEndpoint()?'Quale aspetto della sintesi vuoi approfondire?':'Assistente locale attivo. Per domande libere collega il proxy AI nelle Impostazioni.';
    $('#chatbot-window').setAttribute('aria-label','Lab Copilot');$('#chatbot-input').setAttribute('aria-label','Domanda sulla sintesi');
    const messages=$('#chatbot-messages');(synthesisChatHistory[s.id]||[]).forEach(item=>{const node=document.createElement('div');node.className='chatbot-message '+(item.role==='user'?'user':'bot');node.textContent=item.content||'';messages.append(node);});
    const labView=document.createElement('section');labView.id='lab-mode-view';labView.hidden=!lab;labView.innerHTML=labContent(s);header.before(labView);bindLab(s);
    showSection(section);updateProgress(s);document.body.classList.toggle('lab-mode',lab);
  }
  function action(value) {
    const split=value.indexOf(':'),type=split<0?value:value.slice(0,split),id=split<0?'':value.slice(split+1);
    if(type==='nav')destination(id);
    else if(type==='open'||type==='continue'||type==='analysis'){navigate('detail',id);if(type==='continue')action('lab');if(type==='analysis')showSection('analysis');}
    else if(type==='folder')navigate('folder',id);
    else if(type==='detail')showSection(id);
    else if(type==='new'){if(currentView==='folder')$('#btn-new-synthesis').click();else{currentTab='files';navigate('dashboard');$('#btn-new-synthesis-global').click();}}
    else if(type==='new-project'){currentTab='folders';navigate('dashboard');$('#btn-new-folder').click();}
    else if(type==='search')palette();
    else if(type==='theme'){applyTheme(getCurrentTheme()==='dark'?'light':'dark');if($('#theme-description'))$('#theme-description').textContent=getCurrentTheme()==='dark'?'Scuro':'Chiaro';}
    else if(type==='collapse'){prefs.collapsed=!prefs.collapsed;try{localStorage.setItem('spps-ui-preferences-v1',JSON.stringify(prefs));}catch{}$('.lab-shell').classList.toggle('sidebar-collapsed',prefs.collapsed);}
    else if(type==='export'){exportData();toast('Backup JSON esportato');}
    else if(type==='import')importData();
    else if(type==='connect-inventory')connectInventory();
    else if(type==='connect-workspace')connectWorkspace();
    else if(type==='save-ai'){const value=$('#settings-ai').value.trim();if(value&&!/^https?:\/\//i.test(value)){toast('Inserisci un URL HTTP o HTTPS valido','error');return;}saveAiChatEndpoint(value);toast('Configurazione salvata');}
    else if(type==='inventory-item')inspectInventory(+id);
    else if(type==='copilot'){if(currentView!=='detail'){const s=syntheses.find(s=>s.status==='in-progress')||syntheses[0];if(!s){toast('Crea una sintesi per aprire Lab Copilot');return;}navigate('detail',s.id);}$('#chatbot-bubble').click();}
    else if(type==='lab'){if(currentView!=='detail'){const s=syntheses.find(s=>s.status==='in-progress')||syntheses[0];if(!s){toast('Nessuna sintesi disponibile');return;}navigate('detail',s.id);}lab=true;document.body.classList.add('lab-mode');$('#lab-mode-view').hidden=false;updateProgress(syntheses.find(s=>s.id===currentSynthesisId));window.scrollTo(0,0);}
    else if(type==='exit-lab'){lab=false;document.body.classList.remove('lab-mode');$('#lab-mode-view').hidden=true;showSection(section);}
  }
  function decorate(root=document) {
    root.querySelectorAll('.form-group').forEach(group=>{const label=$('label',group),input=$('input[id],select[id],textarea[id]',group);if(label&&input&&!label.htmlFor)label.htmlFor=input.id;});
    root.querySelectorAll('.synthesis-table tbody tr,.folder-card').forEach(el=>{el.tabIndex=0;el.onkeydown=e=>{if(e.target===el&&(e.key==='Enter'||e.key===' ')){e.preventDefault();el.click();}};});
    root.querySelectorAll('.synthesis-table thead th').forEach((th,index)=>{if(th.querySelector('button'))return;const label=th.textContent;th.innerHTML=`<button class="table-sort">${esc(label)}${icon('ArrowUpDown')}</button>`;$('button',th).onclick=()=>{const body=th.closest('table').tBodies[0],ascending=th.getAttribute('aria-sort')!=='ascending';th.parentElement.querySelectorAll('th').forEach(h=>h.removeAttribute('aria-sort'));th.setAttribute('aria-sort',ascending?'ascending':'descending');[...body.rows].sort((a,b)=>(label==='Data'?String(syntheses.find(s=>s.id===a.dataset.id)?.dateStarted||''):a.cells[index].textContent).localeCompare(label==='Data'?String(syntheses.find(s=>s.id===b.dataset.id)?.dateStarted||''):b.cells[index].textContent,'it',{numeric:true})*(ascending?1:-1)).forEach(row=>body.append(row));};});
    const walker=document.createTreeWalker(root===document?document.body:root,NodeFilter.SHOW_TEXT);const nodes=[];while(walker.nextNode())nodes.push(walker.currentNode);nodes.forEach(node=>{if(node.parentElement.closest('script,style,textarea,input,.sequence-code,.detail-sequence'))return;node.textContent=node.textContent.replace(/[\p{Extended_Pictographic}\uFE0F]/gu,'').replace(/Cartelle Progetto/g,'Progetti').replace(/Gestione Cartelle/g,'Progetti').replace(/Nuova Cartella/g,'Nuovo progetto').replace(/Archivio Globale/g,'Tutte le sintesi');});
    [['#btn-new-folder','Plus'],['#btn-new-synthesis-global','Plus'],['#btn-new-synthesis','Plus'],['#modal-save','Save'],['#edit-modal-save','Save'],['#btn-literature-refresh','RefreshCw'],['#btn-literature-search','Search']].forEach(([selector,name])=>{const el=$(selector,root);if(el&&!el.querySelector('svg')){el.querySelector('span')?.remove();el.insertAdjacentHTML('afterbegin',icon(name));}});
    root.querySelectorAll('.modal-close').forEach(el=>el.setAttribute('aria-label','Chiudi finestra'));
    root.querySelectorAll('.breadcrumb a').forEach(el=>{el.tabIndex=0;el.setAttribute('role','button');el.onkeydown=e=>{if(e.key==='Enter')el.click();};});
  }
  function mount() {
    headerObserver?.disconnect();
    document.body.dataset.page=currentView==='dashboard'?currentTab:currentView;document.body.classList.toggle('lab-mode',currentView==='detail'&&lab);
    if(currentView==='detail')mountDetail(syntheses.find(s=>s.id===currentSynthesisId));
    if(currentView==='research')bindLiteratureTipEvents(researchContext());
    if(currentView==='research'){
      const select=document.createElement('select');select.className='form-select research-context';select.setAttribute('aria-label','Contesto della ricerca');select.innerHTML=syntheses.map(s=>`<option value="${esc(s.id)}" ${s.id===researchContext().id?'selected':''}>${esc(s.name)} · ${esc(s.resinType)}</option>`).join('')||'<option>SPPS generale</option>';$('.literature-section').before(select);select.onchange=()=>{currentSynthesisId=select.value;render();};
    }
    $('#inventory-search')?.addEventListener('input',e=>{$('#global-inventory-rows').innerHTML=inventoryRows(e.target.value);});
    $('#global-inventory-rows')?.addEventListener('click',e=>{const row=e.target.closest('[data-inventory-index]');if(row)inspectInventory(+row.dataset.inventoryIndex);});
    $('#global-inventory-rows')?.addEventListener('keydown',e=>{if(e.key==='Enter')e.target.click();});
    decorate();
    ['form','edit'].forEach(prefix=>{
      [1,2,3].forEach(i=>$('#'+prefix+'-activator-'+i)?.setAttribute('aria-label','Attivatore '+i));
      const body=$('#'+prefix+'-name')?.closest('.modal-body');if(!body||body.querySelector('fieldset'))return;
      const nodes=[...body.children],groups=[['Identità',['name','folder-id','sequence','date']],['Parametri di sintesi',['resin','loading','scale','eq','preloaded-residue']],['Chimica',['activator-1','activator-2','activator-3']]];
      groups.forEach(([label,ids])=>{const fieldset=document.createElement('fieldset');fieldset.className='synthesis-form-group';fieldset.innerHTML=`<legend>${label}</legend>`;nodes.filter(node=>ids.some(id=>node.id===prefix+'-'+id||node.querySelector('#'+prefix+'-'+id))).forEach(node=>fieldset.append(node));body.append(fieldset);});
      const preview=$('#'+prefix+'-preview');if(preview)body.append(preview);
    });
    modalObserver?.disconnect();modalObserver=new MutationObserver(records=>{records.forEach(({target,oldValue})=>{if(!target.matches('.modal-overlay'))return;if(target.classList.contains('active')&&!String(oldValue).includes('active')){target._previousFocus=document.activeElement;target.setAttribute('role','dialog');target.setAttribute('aria-modal','true');target.setAttribute('aria-label',target.querySelector('h2')?.textContent||'Modifica');requestAnimationFrame(()=>target.querySelector('input:not([type=hidden]),textarea,select,button')?.focus());}else if(!target.classList.contains('active')&&String(oldValue).includes('active'))target._previousFocus?.focus();});});document.querySelectorAll('.modal-overlay').forEach(el=>modalObserver.observe(el,{attributes:true,attributeFilter:['class'],attributeOldValue:true}));
    ['form','edit'].forEach(prefix=>{const input=$('#'+prefix+'-sequence');if(!input)return;const preview=document.createElement('div');preview.className='live-sequence-preview';input.after(preview);const refresh=()=>{const tokens=tokenizeSequence(input.value),cycle=analyzeCyclizations(input.value),resin=RESINS.find(r=>r.name===$('#'+prefix+'-resin')?.value);preview.innerHTML=`<div class="sequence-preview">${esc(input.value)}</div><small>${tokens.length} residui · C-terminale ${resin?.type==='acid'?'acido':'ammidico'} · ${tokens.filter(t=>!AMINO_ACIDS[t]).length} non standard · ${cycle.cyclizations.length} ciclizzazioni</small>`;};input.addEventListener('input',refresh);$('#'+prefix+'-resin')?.addEventListener('change',refresh);refresh();});
  }
  document.addEventListener('click',e=>{const el=e.target.closest('[data-ui-action]');if(el)action(el.dataset.uiAction);const file=e.target.closest('[data-attachment-open]');if(file){const item=attachments[currentSynthesisId]?.[file.dataset.attachmentType]?.[+file.dataset.attachmentOpen];if(item)previewFile(item);}});
  document.addEventListener('keydown',e=>{
    if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='k'){e.preventDefault();palette();return;}
    if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='j'){e.preventDefault();action('copilot');return;}
    const overlay=$('.modal-overlay.active');if(overlay){if(e.key==='Escape'){overlay.querySelector('[id$="-cancel"],[id$="-close"]')?.click();}if(e.key==='Tab'){const focusable=[...overlay.querySelectorAll('button,input,select,textarea,[tabindex="0"]')].filter(el=>!el.disabled&&el.offsetParent!==null);const first=focusable[0],last=focusable.at(-1);if(e.shiftKey&&document.activeElement===first){e.preventDefault();last?.focus();}else if(!e.shiftKey&&document.activeElement===last){e.preventDefault();first?.focus();}}}else if(e.key==='Escape'&&$('#synthesis-chatbot.active'))$('#chatbot-close').click();
  });
  return {icon,shell,dashboard,auxiliary,mount,toast,confirm,inspectResidue,updateProgress,decorate,showSection,previewFile};
})();
