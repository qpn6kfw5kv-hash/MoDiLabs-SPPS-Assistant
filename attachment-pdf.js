/* Local PDF.js rendering; scripts and attachments never leave the browser. */
const AttachmentPDF=(()=>{
  let library;
  function bindPanZoom(dialog,viewport,changeZoom){
    const abort=new AbortController(),options={signal:abort.signal};
    let drag=null,delta=0,frame=0;
    viewport.classList.add('file-viewer-grab');
    viewport.addEventListener('dragstart',e=>e.preventDefault(),options);
    viewport.addEventListener('pointerdown',e=>{
      if(e.button!==0||e.pointerType==='touch')return;
      e.preventDefault();drag={id:e.pointerId,x:e.clientX,y:e.clientY,left:viewport.scrollLeft,top:viewport.scrollTop};
      viewport.setPointerCapture(e.pointerId);viewport.classList.add('is-grabbing');
    },options);
    viewport.addEventListener('pointermove',e=>{if(drag&&e.pointerId===drag.id){viewport.scrollLeft=drag.left-e.clientX+drag.x;viewport.scrollTop=drag.top-e.clientY+drag.y;}},options);
    const end=()=>{drag=null;viewport.classList.remove('is-grabbing');};
    viewport.addEventListener('pointerup',end,options);viewport.addEventListener('pointercancel',end,options);viewport.addEventListener('lostpointercapture',end,options);
    viewport.addEventListener('wheel',e=>{
      e.preventDefault();delta+=e.deltaY*(e.deltaMode===1?16:e.deltaMode===2?viewport.clientHeight:1);
      const rect=viewport.getBoundingClientRect(),x=e.clientX-rect.left,y=e.clientY-rect.top;
      cancelAnimationFrame(frame);frame=requestAnimationFrame(async()=>{
        const factor=Math.exp(Math.max(-1,Math.min(1,-delta*.002))),left=viewport.scrollLeft,top=viewport.scrollTop,width=viewport.scrollWidth,height=viewport.scrollHeight;delta=0;
        await changeZoom(factor);if(!dialog.isConnected)return;
        viewport.scrollLeft=(left+x)*viewport.scrollWidth/width-x;viewport.scrollTop=(top+y)*viewport.scrollHeight/height-y;
      });
    },{...options,passive:false});
    dialog.addEventListener('close',()=>{abort.abort();cancelAnimationFrame(frame);end();},{once:true});
  }
  async function mount(dialog,bytes,fallbackUrl){
    const viewport=dialog.querySelector('.file-viewer-viewport'),pages=dialog.querySelector('.file-pdf-pages');
    let closed=false,task,observer,frame;
    dialog.addEventListener('close',()=>{closed=true;observer?.disconnect();cancelAnimationFrame(frame);task?.destroy().catch(()=>{});});
    try{
      if(!library)library=import(new URL('vendor/pdfjs/build/pdf.mjs',document.baseURI).href).catch(error=>{library=null;throw error;});
      const pdfjs=await library;if(closed)return;
      const base=new URL('vendor/pdfjs/',document.baseURI);
      pdfjs.GlobalWorkerOptions.workerSrc=new URL('build/pdf.worker.mjs',base).href;
      task=pdfjs.getDocument({data:bytes,cMapUrl:new URL('cmaps/',base).href,cMapPacked:true,standardFontDataUrl:new URL('standard_fonts/',base).href,wasmUrl:new URL('wasm/',base).href,isEvalSupported:false,enableXfa:false});
      const pdf=await task.promise;if(closed)return;
      let pageNumber=1,zoom=1,queue=Promise.resolve(),revision=0;
      const controls=Object.fromEntries([...dialog.querySelectorAll('[data-preview-action]')].map(b=>[b.dataset.previewAction,b]));
      function render(){
        const current=++revision;
        controls.previous.disabled=pageNumber===1;controls.next.disabled=pageNumber===pdf.numPages;
        dialog.querySelector('.file-viewer-page').textContent=pageNumber+' / '+pdf.numPages;
        dialog.querySelector('.file-viewer-zoom').textContent=Math.round(zoom*100)+'%';
        queue=queue.catch(()=>{}).then(async()=>{
          if(closed||current!==revision)return;
          const page=await pdf.getPage(pageNumber),natural=page.getViewport({scale:1});
          const fit=Math.min((viewport.clientWidth-24)/natural.width,(viewport.clientHeight-24)/natural.height),scale=Math.max(.05,fit)*zoom;
          const view=page.getViewport({scale}),ratio=Math.min(window.devicePixelRatio||1,2,Math.sqrt(16000000/(view.width*view.height)));
          const canvas=document.createElement('canvas');canvas.setAttribute('aria-label','Pagina '+pageNumber);canvas.style.width=Math.round(view.width)+'px';canvas.style.height=Math.round(view.height)+'px';canvas.width=Math.max(1,Math.round(view.width*ratio));canvas.height=Math.max(1,Math.round(view.height*ratio));
          await page.render({canvasContext:canvas.getContext('2d'),viewport:view,transform:[ratio,0,0,ratio,0,0]}).promise;
          if(!closed&&current===revision){pages.replaceChildren(canvas);pages.dataset.rendered='true';}
        }).catch(error=>{if(!closed){pages.textContent='Impossibile visualizzare la pagina PDF.';console.warn(error);}});
        return queue;
      }
      bindPanZoom(dialog,viewport,factor=>{zoom=Math.max(1,Math.min(8,zoom*factor));return render();});
      for(const action of ['in','out','fit','previous','next'])controls[action].onclick=()=>{
        if(action==='in')zoom=Math.min(8,zoom*1.4);
        if(action==='out')zoom=Math.max(1,zoom/1.4);
        if(action==='fit')zoom=1;
        if(action==='previous')pageNumber=Math.max(1,pageNumber-1);
        if(action==='next')pageNumber=Math.min(pdf.numPages,pageNumber+1);
        viewport.scrollTo(0,0);render();
      };
      observer=new ResizeObserver(()=>{cancelAnimationFrame(frame);frame=requestAnimationFrame(render);});observer.observe(viewport);render();
    }catch(error){
      if(closed)return;
      if(location.protocol==='file:'){
        const iframe=document.createElement('iframe');iframe.className='file-preview-pdf';iframe.title=dialog.querySelector('h2').textContent;iframe.src=fallbackUrl+'#view=FitH';viewport.replaceChildren(iframe);
        dialog.querySelectorAll('[data-preview-action]:not([data-preview-action="close"])').forEach(b=>b.disabled=true);
      }else{pages.textContent=error.name==='PasswordException'?'PDF protetto da password.':'Impossibile aprire il PDF. Verifica il file e la presenza della cartella vendor/pdfjs.';console.warn(error);}
    }
  }
  return {mount,bindPanZoom};
})();
