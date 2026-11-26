/* Deprotected neutral peptide graph; input sequence remains N to C. */
const PeptideStructure = (() => {
  const sidechains={A:'C',V:'C(C)C',L:'CC(C)C',I:'[C@@H](C)CC',S:'CO',T:'[C@H](O)C',C:'CS',M:'CCSC',D:'CC(=O)O',E:'CCC(=O)O',N:'CC(=O)N',Q:'CCC(=O)N',K:'CCCCN',R:'CCCNC(=N)N',H:'Cc1cnc[nH]1',F:'Cc1ccccc1',Y:'Cc1ccc(O)cc1',W:'Cc1c[nH]c2ccccc12',Nle:'CCCC'};
  const names={Ala:'A',Val:'V',Leu:'L',Ile:'I',Ser:'S',Thr:'T',Cys:'C',Met:'M',Asp:'D',Glu:'E',Asn:'N',Gln:'Q',Lys:'K',Arg:'R',His:'H',Phe:'F',Tyr:'Y',Trp:'W',Gly:'G',Pro:'P',Aib:'Aib',Nle:'Nle'};
  function identify(token){
    const record=resolveAminoAcid(token);
    const code=record.code||names[record.code3];
    if(code)return {code,d:false};
    const match=String(token).match(/^(?:Fmoc-)?D-(Ala|Val|Leu|Ile|Ser|Thr|Cys|Met|Asp|Glu|Asn|Gln|Lys|Arg|His|Phe|Tyr|Trp|Pro|Nle)(?:\((Boc|Trt|tBu|OtBu|Pbf|Alloc|OAll)\))?(?:-OH)?$/i);
    if(!match)return null;
    const name=Object.keys(names).find(n=>n.toLowerCase()===match[1].toLowerCase());
    const expected={C:'Trt',D:'OtBu',E:'OtBu',N:'Trt',Q:'Trt',K:'Boc',R:'Pbf',H:'Trt',W:'Boc',S:'tBu',T:'tBu',Y:'tBu'};
    const resolved=names[name],protection=match[2]?.toLowerCase();
    if(protection && ![expected[resolved],resolved==='K'?'Alloc':resolved==='D'?'OAll':null].filter(Boolean).some(p=>p.toLowerCase()===protection))return null;
    return {code:resolved,d:true};
  }
  function build(s){
    const tokens=tokenizeSequence(s.sequence),info=analyzeCyclizations(s.sequence);
    if(!tokens.length)throw new Error('Inserisci una sequenza.');
    if(tokens.length>100)throw new Error('Anteprima disponibile fino a 100 residui.');
    const residues=tokens.map(identify),missing=tokens.filter((_,i)=>!residues[i]);
    if(missing.length)throw new Error('Struttura non definita per: '+[...new Set(missing)].join(', ')+'. Il solo peso molecolare non definisce la struttura.');
    if(!info.isValid||info.warnings.length)throw new Error('Ciclizzazione non definita: '+info.warnings.join(' '));
    const bridges=new Map();
    info.cyclizations.forEach((cycle,i)=>{
      const a=cycle.segmentStartTokenIndex,b=cycle.segmentEndTokenIndex;
      if(bridges.has(a)||bridges.has(b))throw new Error('Un gruppo laterale non può partecipare a due ponti.');
      const codes=[residues[a]?.code,residues[b]?.code];
      const valid=cycle.type==='disulfide'?codes.every(c=>c==='C'):cycle.type==='amide'&&codes.includes('K')&&codes.some(c=>c==='D'||c==='E');
      if(!valid)throw new Error('Estremi di ciclizzazione non supportati: '+cycle.startToken+' / '+cycle.endToken+'.');
      if(i>89)throw new Error('Numero di ponti non supportato.');
      [a,b].forEach(index=>bridges.set(index,{ring:'%'+(10+i),type:cycle.type}));
    });
    const resin=RESINS.find(r=>r.name===s.resinType);
    if(!resin)throw new Error('Seleziona una resina con C-terminale definito.');
    const smiles=residues.map(({code,d},i)=>{
      const stereo=d?'@':'@@',bridge=bridges.get(i);
      if(code==='G')return 'NCC(=O)';
      if(code==='Aib')return 'NC(C)(C)C(=O)';
      if(code==='P')return `N1[C${stereo}H](CCC1)C(=O)`;
      let side=sidechains[code];
      if(d && code==='T')side='[C@@H](O)C';
      if(d && code==='I')side='[C@H](C)CC';
      if(bridge){
        if(code==='C')side='CS'+bridge.ring;
        if(code==='K')side='CCCCN'+bridge.ring;
        if(code==='D')side='CC'+bridge.ring+'(=O)';
        if(code==='E')side='CCC'+bridge.ring+'(=O)';
      }
      return `N[C${stereo}H](${side})C(=O)`;
    }).join('')+(resin.type==='acid'?'O':'N');
    return {smiles,count:tokens.length,bridges:info.cyclizations.length,terminus:resin.type};
  }
  let worker,serial=0;
  const pending=new Map(),cache=new Map();
  function request(model){
    return new Promise((resolve,reject)=>{
      if(!worker){
        worker=new Worker(new URL('peptide-structure-worker.js',document.baseURI));
        worker.onmessage=({data})=>{const job=pending.get(data.id);if(!job)return;clearTimeout(job.timer);pending.delete(data.id);data.error?job.reject(new Error(data.error)):job.resolve(data.result);};
        worker.onerror=()=>{for(const job of pending.values()){clearTimeout(job.timer);job.reject(new Error('Motore 2D non disponibile. Apri il sito tramite GitHub Pages o un server locale.'));}pending.clear();worker?.terminate();worker=null;cache.clear();};
      }
      const id=++serial,timer=setTimeout(()=>{worker?.terminate();worker=null;for(const job of pending.values()){clearTimeout(job.timer);job.reject(new Error('Generazione 2D non completata. Riprova con una sequenza più breve.'));}pending.clear();cache.clear();},45000);
      pending.set(id,{resolve,reject,timer});worker.postMessage({id,...model});
    });
  }
  async function depiction(s){
    const model=build(s),key=model.smiles;
    if(!cache.has(key)){
      if(cache.size>=12)cache.delete(cache.keys().next().value);
      cache.set(key,request(model).catch(e=>{cache.delete(key);throw e;}));
    }
    return {...await cache.get(key),...model};
  }
  const esc=s=>escapeHtml(String(s));
  function panel(){
    // Lucide 0.468.0, ISC: same icon-node format as the application's icon set.
    Object.assign(window.LAB_ICONS,{
      Minus:['svg',{},[['path',{d:'M5 12h14'}]]],
      Maximize:['svg',{},['M8 3H5a2 2 0 0 0-2 2v3','M21 8V5a2 2 0 0 0-2-2h-3','M3 16v3a2 2 0 0 0 2 2h3','M16 21h3a2 2 0 0 0 2-2v-3'].map(d=>['path',{d}])],
      Scan:['svg',{},['M3 7V5a2 2 0 0 1 2-2h2','M17 3h2a2 2 0 0 1 2 2v2','M21 17v2a2 2 0 0 1-2 2h-2','M7 21H5a2 2 0 0 1-2-2v-2'].map(d=>['path',{d}])]
    });
    const button=(action,label,icon)=>`<button type="button" class="btn btn-secondary btn-sm structure-tool" data-structure-action="${action}" title="${label}" aria-label="${label}" disabled>${LabUI.icon(icon)}</button>`;
    return `<section class="section-card peptide-structure-panel" data-detail-section="overview synthesis"><div class="section-card-header structure-header"><div><div class="section-card-title">Struttura del peptide</div><div class="structure-caption">Composto finale deprotetto · forma neutra</div></div><div class="structure-tools">${button('out','Riduci','Minus')}${button('in','Ingrandisci','Plus')}${button('fit','Adatta','Scan')}${button('full','Schermo intero','Maximize')}${button('svg','Scarica SVG','Download')}${button('mol','Scarica MOL','FileText')}</div></div><div class="structure-status" role="status">Generazione della struttura 2D…</div><div class="structure-viewport" tabindex="0" aria-label="Struttura chimica bidimensionale" hidden><div class="structure-drawing"></div></div><div class="structure-metadata"></div></section>`;
  }
  function download(value,type,name){const url=URL.createObjectURL(new Blob([value],{type})),a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
  async function mount(s){
    const host=document.querySelector('.peptide-structure-panel');if(!host)return;
    const status=host.querySelector('.structure-status');
    try{
      const data=await depiction(s);if(!host.isConnected)return;
      const viewport=host.querySelector('.structure-viewport'),drawing=host.querySelector('.structure-drawing');
      drawing.innerHTML=data.svg;viewport.hidden=false;status.hidden=true;
      host.querySelector('.structure-metadata').textContent=`${data.formula} · ${data.count} residui · ${data.bridges} ${data.bridges===1?'ponte':'ponti'} · C-terminale ${data.terminus==='acid'?'–COOH':'–CONH₂'}`;
      let zoom=1;
      const resize=()=>{drawing.style.width=Math.round(viewport.clientWidth*zoom)+'px';drawing.style.height=Math.round(viewport.clientHeight*zoom)+'px';};
      const observer=new ResizeObserver(resize);observer.observe(viewport);
      const removed=new MutationObserver(()=>{if(!host.isConnected){observer.disconnect();removed.disconnect();}});removed.observe(document.body,{childList:true,subtree:true});
      host.querySelectorAll('[data-structure-action]').forEach(button=>{
        button.disabled=false;button.onclick=async()=>{
          const action=button.dataset.structureAction;
          if(action==='in')zoom=Math.min(6,zoom*1.4);
          if(action==='out')zoom=Math.max(.5,zoom/1.4);
          if(action==='fit'){zoom=1;viewport.scrollTo(0,0);}
          if(action==='full'){try{if(document.fullscreenElement===host)await document.exitFullscreen();else await host.requestFullscreen();}catch{status.hidden=false;status.textContent='Schermo intero non disponibile in questo browser.';}}
          const filename=(s.name||'peptide').replace(/[^a-z0-9_-]/gi,'_');
          if(action==='svg')download(data.svg,'image/svg+xml',filename+'.svg');
          if(action==='mol')download(data.molblock,'chemical/x-mdl-molfile',filename+'.mol');
          resize();
        };
      });resize();
    }catch(e){if(host.isConnected){status.textContent=e.message;host.dataset.unavailable='true';}}
  }
  async function print(s){
    try{const data=await depiction(s);return `<section class="print-peptide-structure"><h4>STRUTTURA DEL COMPOSTO FINALE · FORMA NEUTRA</h4>${data.svg}<div>${esc(data.formula)} · C-terminale ${data.terminus==='acid'?'COOH':'CONH₂'}</div></section>`;}
    catch(e){return `<section class="print-peptide-structure"><h4>STRUTTURA DEL COMPOSTO FINALE</h4><p>${esc(e.message)}</p></section>`;}
  }
  return {build,identify,depiction,panel,mount,print};
})();
