/* ===== v2: carga real de documentos, evaluación de 5 preguntas (80 %), firma condicionada y descarga por idioma ===== */
const PASS=80;
const XLSXJS='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',MAMMOTH='https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js',PDFJS='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',PDFW='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
const loadLib=u=>new Promise((ok,no)=>{const s=document.createElement('script');s.src=u;s.onload=ok;s.onerror=()=>no(new Error('lib'));document.head.appendChild(s)});
const docHash=d=>{let h=0;const q=(d.secs||[]).map(z=>(z.h||'')+'|'+(z.t||'')).join('¶');for(let i=0;i<q.length;i++)h=(h*31+q.charCodeAt(i))|0;return h};
const pause=ms=>new Promise(r=>setTimeout(r,ms));

/* ---------- lectura de archivos ---------- */
function htmlToMd(html){
  const doc=new DOMParser().parseFromString(html,'text/html'),out=[];
  const inl=n=>{let s='';n.childNodes.forEach(c=>{if(c.nodeType==3)s+=c.nodeValue;else if(c.nodeType==1){const g=c.tagName.toLowerCase();if(g=='strong'||g=='b'){const v=inl(c).trim();s+=v?'**'+v+'**':''}else if(g=='br')s+=' ';else s+=inl(c)}});return s.replace(/\s+/g,' ')};
  const walk=n=>{n.childNodes.forEach(c=>{if(c.nodeType!=1)return;const g=c.tagName.toLowerCase();
    if(/^h[1-6]$/.test(g)){const v=inl(c).replace(/\*\*/g,'').trim();if(v)out.push('','# '+v)}
    else if(g=='p'){const v=inl(c).trim();if(v)out.push(v)}
    else if(g=='ul'||g=='ol'){c.querySelectorAll(':scope > li').forEach(li=>{const v=inl(li).trim();if(v)out.push('- '+v)})}
    else if(g=='table'){out.push('');c.querySelectorAll('tr').forEach(tr=>{out.push('| '+[...tr.children].map(td=>inl(td).replace(/\|/g,'/').trim()).join(' | ')+' |')});out.push('')}
    else walk(c)})};
  walk(doc.body);return out.join('\n')}
const isHead=l=>{const z=(l||'').trim();if(!z||z.length>90||z.includes('|'))return false;if(/^#{1,6}\s+/.test(z))return true;if(/[.:;,]$/.test(z))return false;if(/^\d+(\.\d+)*[.)]?\s+[A-ZÁÉÍÓÚÑ]/.test(z)&&z.split(' ').length<=12)return true;const L=z.replace(/[^A-Za-zÁÉÍÓÚÑáéíóúñ]/g,'');return L.length>=5&&z===z.toUpperCase()&&z.split(' ').length<=12};
function rewrap(t){const out=[];t.replace(/\r/g,'').split('\n').forEach(raw=>{const l=raw.replace(/\s+/g,' ').trim();if(!l){out.push('');return}const p=out.length?out[out.length-1]:'';
  if(p&&!/[.:;!?)»"]$/.test(p)&&!isHead(p)&&!isHead(l)&&!/^([-•*·]|\d+[.)])\s/.test(l)&&!p.includes('|')&&!l.includes('|')&&!/^#/.test(l))out[out.length-1]=p+' '+l;else out.push(l.replace(/^[•·]\s*/,'- '))});return out.join('\n')}
async function pdfToText(f){await loadLib(PDFJS);pdfjsLib.GlobalWorkerOptions.workerSrc=PDFW;const pdf=await pdfjsLib.getDocument({data:new Uint8Array(await f.arrayBuffer())}).promise,pages=[];
  for(let p=1;p<=pdf.numPages;p++){const c=await(await pdf.getPage(p)).getTextContent(),rows=[];let y=null,cur='';c.items.forEach(it=>{const yy=Math.round(it.transform[5]);if(y!==null&&Math.abs(yy-y)>3){rows.push(cur);cur=''}y=yy;cur+=it.str+(it.hasEOL?'':'')});rows.push(cur);pages.push(rows.join('\n'))}
  const t=pages.join('\n');if(t.replace(/\s/g,'').length<40)throw new Error('scan');return t}
async function fileToMd(f){const n=f.name.toLowerCase();
  if(/\.docx$/.test(n)){await loadLib(MAMMOTH);const r=await mammoth.convertToHtml({arrayBuffer:await f.arrayBuffer()});return htmlToMd(r.value)}
  if(/\.pdf$/.test(n))return rewrap(await pdfToText(f));
  if(/\.xlsx$/.test(n)){await loadLib(XLSXJS);const wb=XLSX.read(await f.arrayBuffer(),{type:'array'});return wb.SheetNames.map(sn=>{const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,blankrows:false,defval:''}).map(r=>r.map(c=>String(c).replace(/\s+/g,' ').replace(/\|/g,'/').trim())).filter(r=>r.some(Boolean)).slice(0,600);const w=Math.max(0,...rows.map(r=>{let k=r.length;while(k&&!r[k-1])k--;return k}));return'# '+sn.trim()+'\n'+rows.map(r=>'| '+r.slice(0,w).join(' | ')+' |').join('\n')}).join('\n\n')}
  throw new Error('fmt')}
function splitSections(md,title){
  const lines=md.replace(/\r/g,'').split('\n'),secs=[];let cur={h:title,b:[],first:1};
  const push=()=>{const t=cur.b.join('\n').replace(/\n{2,}/g,'\n').trim();if(t.length>=20||(!cur.first&&cur.h&&t))secs.push({h:cur.h.slice(0,120),t})};
  lines.forEach(l=>{if(isHead(l)){const h=l.trim().replace(/^#{1,6}\s+/,'').replace(/\*\*/g,'');if(!cur.b.join('').trim()&&!cur.first){cur.h=secs.length?(cur.h+' — '+h).slice(0,120):h}else{push();cur={h,b:[]}}}else cur.b.push(l)});
  push();
  if(secs.length<2){const all=lines.map(z=>z.trim()).filter(Boolean),out=[];let c=[],n=0;all.forEach(z=>{c.push(z);n+=z.length;if(n>1800){out.push(c);c=[];n=0}});if(c.length)out.push(c);return out.map((c,k)=>({h:out.length>1?title+' · '+(k+1):title,t:c.join('\n')}))}
  return secs}

/* ---------- creación del documento ---------- */
async function createDocV2(ti,law,area,secs){const id=1000+Math.floor(Math.random()*9e5),ids=secs.map((z,k)=>'s'+id+'_'+k);
  ids.forEach((sid,k)=>{SS[sid]={id:sid,doc:id,h:secs[k].h,t:secs[k].t}});MM['d'+id]={id,title:ti,law:law||'',cat:'doc',area:area||'all',nw:1,order:ids};build();
  if(dbh){await put('docs','d'+id,MM['d'+id]);await Promise.all(ids.map(s=>put('secs',s,SS[s])))}return id}

/* ---------- texto plano del documento (uso interno) ---------- */
const plainSecs=d=>(d.secs||[]).map(z=>({h:z.h,t:(z.t||'').split('\n').filter(l=>l.trim()&&l.trim()[0]!='|').map(l=>l.replace(/\*\*/g,'').replace(/^\s*[-*•]\s+/,'').trim()).join('\n')}));
const docPlain=d=>plainSecs(d).map(z=>z.h+'\n'+z.t).join('\n\n');

/* ---------- carga de documentos (administrador) ---------- */
V.sub=function(){return`<div class="card"><p class="mut">${x('Sube un documento en Word, PDF o Excel. El archivo original queda guardado y es el que se descarga. Luego escribes las preguntas de evaluación: el documento no es visible para los trabajadores hasta que las tenga.','Upload a Word, PDF or Excel document. The original file is stored and is what gets downloaded. Then you write the test questions: the document is not visible to workers until it has them.')}</p><label>${x('Tipo','Type')}</label><select id="s1"><option value="d">${x('Documento o protocolo','Document or protocol')}</option><option value="d">${x('Documento legal / contrato','Legal document / contract')}</option></select><label>${t('tt')}</label><input id="s2"><label>${t('lw')}</label><select id="s3">${['DS 44','Ley 16.744','DS 594','Código del Trabajo','Ley 21.643','Contrato'].map(l=>`<option value="${l}">${tx(l)}</option>`).join('')}</select><label>${x('Dirigido a','Audience')}</label><select id="s4"><option value="all">${x('Todas las áreas','All areas')}</option>${AO('')}</select><label>${x('Archivo (.docx, .pdf, .xlsx)','File (.docx, .pdf, .xlsx)')}</label><input type="file" id="s5" accept=".docx,.pdf,.xlsx" onchange="if(!$('#s2').value&&this.files[0])$('#s2').value=this.files[0].name.replace(/\\.[^.]+$/,'')"><label>${x('…o pega el texto del documento','…or paste the document text')}</label><textarea id="s6" rows="7"></textarea><p></p><button id="pubb" onclick="pub()">${x('Publicar','Publish')}</button></div>`};
async function pub(){const ty=$('#s1').value,v=$('#s2').value.trim(),f=$('#s5').files[0],pasted=$('#s6').value.trim(),law=$('#s3').value,area=$('#s4').value;
  if(!f){toast(x('Sube el archivo original (Word, PDF o Excel)','Upload the original file (Word, PDF or Excel)'));return}
  if(!/\.(docx|pdf|xlsx)$/i.test(f.name)){toast(x('Formato no permitido. Solo Word (.docx), PDF o Excel (.xlsx).','Format not allowed. Only Word (.docx), PDF or Excel (.xlsx).'));return}
  const title=v||(f?f.name.replace(/\.[^.]+$/,''):'');if(!title){toast(x('Escribe un título','Enter a title'));return}
  M(`<h3>${x('Procesando documento…','Processing document…')}</h3><p class="mut" id="pst">${x('Leyendo el archivo…','Reading the file…')}</p><div class="bar"><i style="width:30%"></i></div>`);
  try{let md=f?await fileToMd(f):rewrap(pasted);if(f&&pasted)md+='\n\n'+rewrap(pasted);
    const secs=splitSections(md,title);if(!secs.length||secs.map(z=>z.t).join('').length<80)throw new Error('empty');
    $('#pst').textContent=x('Publicando el documento…','Publishing the document…');const id=await createDocV2(title,law,area,secs);
    $('#pst').textContent=x('Guardando el archivo original…','Saving the original file…');await upFile(id,'es',f,'original');
    M(`<h3>✓ ${x('Documento guardado','Document saved')}</h3><p><b>${esc(title)}</b></p><p>${x('Ahora escribe las preguntas de evaluación. <b>El documento no será visible para los trabajadores hasta que tenga preguntas.</b> Al guardarlas, se traducen una sola vez a los idiomas del documento y quedan guardadas en Supabase.','Now write the test questions. <b>The document will not be visible to workers until it has questions.</b> When saved, they are translated once into the document languages and stored in Supabase.')}</p><p class="mut">${x('Las traducciones del archivo se suben con el botón 📁 en Documentos.','File translations are uploaded with the 📁 button in Documents.')}</p><div style="display:flex;gap:8px;flex-wrap:wrap"><button onclick="edQz(${id})">✎ ${x('Escribir preguntas','Write questions')}</button><button class="g" onclick="fmOpen(${id})">📁 ${x('Archivos por idioma','Files by language')}</button><button class="g" onclick="X()">${t('cl')}</button></div>`)}
  catch(e){const m=e&&e.message=='scan'?x('El PDF no tiene texto (parece escaneado). Sube un Word o pega el texto.','The PDF has no text (looks scanned). Upload a Word file or paste the text.'):e&&e.message=='fmt'?x('Formato no permitido. Solo Word (.docx), PDF o Excel (.xlsx).','Format not allowed. Only Word (.docx), PDF or Excel (.xlsx).'):e&&e.message=='lib'?x('No se pudo cargar el lector de archivos. Revisa tu conexión o pega el texto.','Could not load the file reader. Check your connection or paste the text.'):x('No se pudo leer contenido suficiente del archivo.','Not enough content could be read from the file.');
    M(`<h3>${x('No se pudo publicar','Could not publish')}</h3><p>${m}</p><button class="g" onclick="X()">${t('cl')}</button>`)}}
const allQ=()=>[...DOCS.map(d=>({id:d.id,law:d.law,t:{es:d.title,_doc:1}}))];

/* ---------- evaluación aprobada con 80 %: firma habilitada ---------- */
function qfail(id,sc,ok){const i=items().find(z=>z.id==id),n=QS(id).length;M(`<h3>${t('sc')}: ${sc}%</h3><p>${ok}/${n} ${x('respuestas correctas.','correct answers.')}</p><p><b>${x('Necesitas al menos 80 % para aprobar y habilitar la firma.','You need at least 80% to pass and unlock signing.')}</b><br>${x('Debes repetir la evaluación.','You must retake the test.')}</p><div style="display:flex;gap:8px;flex-wrap:wrap"><button onclick="quiz(${id})">${t('rt')}</button>${i&&i.secs?`<button class="g" onclick="reader(items().find(z=>z.id==${id}),0)">${x('Volver a leer','Read again')}</button>`:''}<button class="g" onclick="X()">${t('cl')}</button></div>`)}
function nextR(id){setProg(id,1);hasQ(id)?quiz(id):sig(items().find(z=>z.id==id),null)}

/* ---------- descargas: solo el archivo guardado en Supabase, en su formato original (Word, PDF o Excel) ---------- */
const FEXT={pdf:'PDF',docx:'Word',xlsx:'Excel'},FICO={pdf:'📕',docx:'📘',xlsx:'📗'};
const fext=n=>(String(n||'').toLowerCase().match(/\.([a-z0-9]+)$/)||[])[1]||'';
async function filesOf(id){try{const r=await fetch('/api/files?doc='+id);if(!r.ok)return[];return(await r.json()).files||[]}catch(e){return[]}}
const dlBtns=id=>`<span class="dlx" data-doc="${id}" style="display:inline-flex;gap:8px;flex-wrap:wrap"><span class="mut">${x('Buscando archivos…','Looking for files…')}</span></span>`;
async function fillDl(id){const box=document.querySelector(`.dlx[data-doc="${id}"]`);if(!box)return;const fs=(await filesOf(id)).filter(f=>role=='p'||f.status=='published');
  if(!box.isConnected)return;if(!fs.length){box.innerHTML=`<span class="mut">${x('Este documento aún no tiene archivos para descargar.','This document has no files to download yet.')}</span>`;return}
  fs.sort((a,b)=>(a.lang==lang?-1:b.lang==lang?1:0)||(a.kind=='original'?-1:1));
  box.innerHTML=fs.map(f=>{const e=fext(f.file_name);return`<span style="display:inline-flex;gap:4px;align-items:center"><a class="btnl" style="${f.lang==lang?'':'background:transparent;color:var(--blue);border:1px solid var(--blue)'}" href="${esc(f.url||'#')}" download="${esc(f.file_name)}">${FICO[e]||'⬇'} ${LN[f.lang]||f.lang.toUpperCase()} · ${FEXT[e]||e.toUpperCase()}</a>${role=='p'&&f.kind!='original'?(f.status=='published'?`<span class="tag ok">${x('publicado','published')}</span>`:`<button class="g" onclick="pubFile(${id},'${f.lang}')">${x('Publicar','Publish')}</button>`):''}</span>`}).join('')}
async function pubFile(id,l){try{const r=await fetch('/api/files/status',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({doc:id,lang:l,status:'published'})});if(!r.ok)throw new Error((await r.json().catch(()=>({}))).error||'HTTP '+r.status);toast('✓ '+x('Publicado para los trabajadores','Published for workers'))}catch(e){toast(e.message)}fillDl(id)}
async function upFile(id,l,file,kind){return upDirect(id,l,file,kind||'translation','review','')}
function doneModal(id){const i=items().find(z=>z.id==id)||{t:{es:''}};M(`<h3>✓ ${t('sd')}</h3><p><b>${esc(T4(i.t))}</b></p><p class="mut">${x('Puedes descargar el documento en su formato original, en los idiomas disponibles.','You can download the document in its original format, in the available languages.')}</p>${dlBtns(id)}<p></p><button onclick="X()">${t('cl')}</button>`);fillDl(id)}

/* ===== v3: avances en 0, nuevos trabajadores, avances guardados por usuario ===== */
const ZERO={ps:Date.now(),now:Date.now(),i:{},c:{}};
function localRec(m){try{const nk=s=>s.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'');let key=null;for(let j=0;j<localStorage.length;j++){const k=localStorage.key(j);if(/^pia_/.test(k)&&!/_(ps|av)$/.test(k)&&!/^pia_(lang|tr_.*)$/.test(k)&&nk(k.slice(4))==m.u){key=k;break}}
  if(!key)return null;const o=JSON.parse(localStorage.getItem(key)||'null');if(!o||!o.p)return null;const i={};for(const id in o.p){const p=o.p[id];i[id]=[p.d||0,p.t0||0,p.t1||0,p.sc??null,o.s&&o.s[id]&&o.s[id].s?1:0]}
  const c={};CAMP.forEach(q=>{const z=(o.c||{})[q.id];if(z&&z.r)c[q.id]=Math.round((z.r.filter(Boolean).length/q.r.length+z.k/q.k.length+(z.d?1:0))/3*100)});
  return{u:m.u,ps:+localStorage.getItem(key+'_ps')||ZERO.ps,now:Date.now(),i,c}}catch(e){return null}}
[['Ricardo Paredes','Capataz / jefe de cuadrilla','4417'],['René Triviño','Albañil','2593'],['German','Operador de maquinaria','7126'],['Daniela','Administrativo/a de obra','3860'],['Marcela Jaramillo','Recursos humanos y remuneraciones','5274'],['Macarena Mancilla','Prevencionista de riesgos','6039'],['Rodrigo Gonzalez','Jefe de obra / ingeniero residente','1948'],['Paulina Beltran','Contabilidad y finanzas','8352'],['Anai','Bodeguero','2761']].forEach(([n,c,k])=>{const u=n.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'');if(!TSEED.some(m=>m.u==u))TSEED.push({n,a:cargoArea(c),cargo:c,l:'es',p:0,u,k})});
TSEED.forEach(m=>{m.p=0});mergeTeam();
{const aul=applyUserLang;applyUserLang=function(){aul();try{const m=meRec();if(m&&!prof.cargo){prof.cargo=m.cargo||'';prof.area=m.a;area=m.a}}catch(e){}}}
function askReset(){M(`<h3>${x('Reiniciar todos los avances','Reset all progress')}</h3><p>${x('Se borrarán las firmas, evaluaciones y campañas de todos los trabajadores y el panel quedará en 0 %. No se pueden recuperar.','Signatures, tests and campaigns of all workers will be erased and the dashboard will go back to 0%. This cannot be undone.')}</p><div style="display:flex;gap:8px"><button style="background:var(--red)" onclick="doReset()">${x('Sí, dejar todo en 0','Yes, reset to 0')}</button><button class="g" onclick="X()">${x('Cancelar','Cancel')}</button></div>`)}
async function doReset(){try{Object.keys(localStorage).forEach(k=>{if(/_ps$/.test(k)&&/^pia_/.test(k))localStorage.removeItem(k);else if(/^pia_/.test(k)&&!/_av$/.test(k)&&!/^pia_(lang|tr_.*)$/.test(k)){try{const o=JSON.parse(localStorage.getItem(k));if(o&&o.p)localStorage.removeItem(k)}catch(e){}}})}catch(e){}
  if(dbh){const ks=new Set([...Object.keys(PR),...team.map(m=>m.u)]);try{await fetch('/api/data/reset-progress',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({keys:[...ks]})})}catch(e){}}
  Object.keys(PR).forEach(k=>delete PR[k]);X();toast('✓ '+x('Avances en 0','Progress reset'))}
{const e1=V.emp;V.emp=function(){return e1.call(this)+`<div class="card" style="margin-top:16px"><b>${x('Reiniciar avances','Reset progress')}</b><p class="mut">${x('Deja en 0 las firmas y avances de todos los trabajadores. Cada trabajador guarda sus avances con su propio usuario.','Sets every worker\'s signatures and progress back to 0. Each worker\'s progress is saved under their own username.')}</p><button class="g" onclick="askReset()">${x('Dejar todo en 0','Reset everything to 0')}</button></div>`}}

{const l0=loadLS;loadLS=function(){try{['','_ps','_av'].forEach(sf=>{const o='pia_'+user.toLowerCase()+sf,n='pia_'+ukey()+sf;if(o!=n&&localStorage.getItem(o)!=null&&localStorage.getItem(n)==null)localStorage.setItem(n,localStorage.getItem(o))})}catch(e){}l0()}}


/* ---------- archivos por idioma (administrador): original + traducciones en el mismo formato ---------- */
async function fmOpen(id){if(role!='p')return;const d=dm(id);if(!d)return X();M(`<h3>📁 ${x('Archivos del documento','Document files')}</h3><p class="mut">${x('Cargando…','Loading…')}</p>`);
  const fs=await filesOf(id),by={};fs.forEach(f=>by[f.lang]=f);const o=fs.find(f=>f.kind=='original'),oe=o?fext(o.file_name):'';
  const st=f=>!f?`<span class="tag">${x('sin archivo','no file')}</span>`:f.status=='published'?`<span class="tag ok">✓ ${x('publicado','published')}</span>`:`<span class="tag go">${x('en revisión','in review')}</span>`;
  const row=(l)=>{const f=by[l];return`<div class="row"><div style="min-width:0"><b>${LN[l]||l}</b> ${l=='es'?`<span class="mut">(${x('original','original')})</span>`:''}<br>${st(f)} ${f?`<a href="${esc(f.url||'#')}" download="${esc(f.file_name)}">${esc(f.file_name)}</a>`:''}</div><span style="white-space:nowrap;display:inline-flex;gap:6px;align-items:center">${f&&l!='es'&&f.status!='published'?`<button class="g" onclick="pubFile(${id},'${l}').then(()=>fmOpen(${id}))">${x('Publicar','Publish')}</button>`:''}${(l=='es'||o)?`<label class="btnl" style="cursor:pointer;margin:0">⬆ ${f?x('Reemplazar','Replace'):x('Subir','Upload')}<input type="file" accept="${l=='es'?'.docx,.pdf,.xlsx':'.'+oe}" style="display:none" onchange="fmUp(${id},'${l}',this)"></label>`:''}</span></div>`};
  M(`<h3>📁 ${x('Archivos del documento','Document files')}</h3><p><b>${esc(docTitle(d))}</b></p><p class="mut">${o?x('Formato del original: ','Original format: ')+(FEXT[oe]||oe)+'. '+x('Cada traducción debe subirse en ese mismo formato y queda disponible de inmediato para los trabajadores.','Each translation must be uploaded in that same format and is available to workers right away.'):x('Primero sube el archivo original en español (Word, PDF o Excel).','First upload the original Spanish file (Word, PDF or Excel).')}</p>${['es',...tLangs(id)].map(row).join('')}<p></p><button onclick="X()">${t('cl')}</button>`)}
async function fmUp(id,l,inp){const f=inp.files&&inp.files[0];if(!f)return;M(`<h3>${x('Subiendo archivo…','Uploading file…')}</h3><p class="mut">${esc(f.name)}</p><div class="bar"><i style="width:60%"></i></div>`);
  try{await upFile2(id,l,f,l=='es'?'original':'translation');toast('✓ '+x('Archivo guardado en Supabase','File saved in Supabase'))}catch(e){toast(e.message)}fmOpen(id)}

/* ---------- idiomas disponibles según los archivos guardados (no según la traducción de texto) ---------- */
const FLG={};let flgReady=0;
async function loadFL(){try{const r=await fetch('/api/files?lite=1');if(!r.ok)return;const j=await r.json();Object.keys(FLG).forEach(k=>delete FLG[k]);(j.files||[]).forEach(f=>{(FLG[f.doc_id]||(FLG[f.doc_id]={}))[f.lang]=f.status});flgReady=1;if($('#app').style.display=='grid'&&!$('#m').classList.contains('on'))render()}catch(e){}}
const fileLangs=id=>{const o=FLG[id]||{};return['es','en','ht','pt','de'].filter(l=>o[l]&&(role=='p'||o[l]=='published'))};
const flTags=id=>{const o=FLG[id]||{};return['es','en','ht','pt','de'].filter(l=>o[l]&&(role=='p'||o[l]=='published')).map(l=>`<span class="tag ${o[l]=='published'?(l==lang?'ok':''):'go'}" title="${o[l]=='published'?x('Disponible para descargar','Available to download'):x('En revisión','In review')}">${l.toUpperCase()}${role=='p'?(o[l]=='published'?' ✓':' …'):''}</span>`).join(' ')||`<span class="mut">${x('sin archivos','no files')}</span>`};
loadFL();
{const p0=pubFile;pubFile=async function(id,l){await p0(id,l);loadFL()}}
{const u0=upFile;upFile=async function(id,l,f,k){const r=await u0(id,l,f,k);loadFL();return r}}


/* ---------- evaluación: escrita por el prevencionista, traducida una vez con la API y guardada en Supabase ---------- */
const QL=['en','ht','pt','de'];
const qzLangs=id=>tLangs(id);
const qzState=(id,l)=>{const es=QES(id),t=QZT[id+'_'+l];if(!es.length)return'none';if(!t)return'missing';return t.src==qkey(es)&&t.qs.length==es.length?'ok':'old'};
{const e0=edQz;edQz=function(id){e0(id);const box=$('#mc');if(!box)return;const es=QES(id);
  box.insertAdjacentHTML('beforeend',`<div class="card" style="margin-top:12px;padding:12px 16px"><b>🌐 ${x('Traducciones de las preguntas','Question translations')}</b><p class="mut" style="margin:4px 0 8px">${es.length?x('Se traducen automáticamente con IA y quedan guardadas en Supabase. Si cambias una pregunta, se vuelve a traducir al guardar.','Translated automatically with AI and stored in Supabase. If you change a question, it is translated again on save.'):x('Escribe al menos una pregunta. Mientras no haya preguntas, el documento no es visible para los trabajadores.','Write at least one question. Until there are questions, the document is hidden from workers.')}</p>${es.length?qzLangs(id).map(l=>{const s=qzState(id,l);return`<span class="tag ${s=='ok'?'ok':s=='old'?'go':''}">${l.toUpperCase()} ${s=='ok'?'✓':s=='old'?x('desactualizada','outdated'):x('falta','missing')}</span>`}).join(' ')+`<p style="margin:10px 0 0"><button onclick="trAllQ(${id})">🌐 ${x('Traducir a todos los idiomas','Translate into all languages')}</button></p>`:''}</div>`)}}
{const s0=qzSave;qzSave=function(id,qs){const prev=JSON.parse(JSON.stringify(QES(id)));s0(id,qs);trQuiz(id,0,prev)}}
/* PRESENTACIÓN: la traducción automática de preguntas está SIMULADA (no hay API de traducción conectada).
   Las preguntas que no cambiaron conservan su traducción real; las nuevas o editadas se guardan como copia
   en español con engine='simulado' hasta conectar la traducción real. */
async function trQuiz(id,force,prev){const es=QES(id);if(!es.length)return;const src=qkey(es),langs=qzLangs(id).filter(l=>qzState(id,l)!='ok');
  if(!langs.length){if(force)toast('✓ '+x('Las preguntas ya están traducidas','Questions are already translated'));return}
  const same=(a,b)=>JSON.stringify([a.q,a.o,a.c||0,a.w||''])==JSON.stringify([b.q,b.o,b.c||0,b.w||'']);let done=0;
  for(const l of langs){toast(x('Traduciendo preguntas','Translating questions')+' · '+(LN[l]||l)+'…');await pause(700+Math.random()*500);
    const old=QZT[id+'_'+l],oldOk=old&&prev&&old.src==qkey(prev)&&old.qs.length==prev.length;let sim=0;
    const qs=es.map(q=>{const j=oldOk?prev.findIndex(z=>same(z,q)):-1;if(j>=0)return old.qs[j];sim=1;return JSON.parse(JSON.stringify(q))});
    const rec={id,lang:l,qs,src,engine:sim?'simulado':'manual'};QZT[id+'_'+l]=rec;await put('qzt',id+'_'+l,rec);done++}
  toast('✓ '+x('Preguntas traducidas y guardadas','Questions translated and saved')+' ('+done+')');if($('#mc')&&$('#m').classList.contains('on')&&$('#mc').innerText.includes(L5(NS.ev)))edQz(id)}


/* ---------- previsualizador de archivos (reemplaza la vista de texto): PDF, Word y Excel en el idioma elegido ---------- */
const JSZIPJS='https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',DOCXJS='https://cdn.jsdelivr.net/npm/docx-preview@0.4.1/dist/docx-preview.min.js';
const libs={};const need=u=>libs[u]||(libs[u]=loadLib(u));
let PV=null;
async function viewer(id,ro,vl){const d=dm(id);if(!d)return;const tok=PV={id,ro};
  M(`<div class="pvw"><h3 style="margin:0 0 6px">${esc(docTitle(d,vl||lang))}</h3><p class="mut">${x('Cargando documento…','Loading document…')}</p></div>`);
  const fs=(await filesOf(id)).filter(f=>role=='p'||f.status=='published');if(PV!==tok)return;
  if(!fs.length){M(`<div class="pvw"><h3>${esc(docTitle(d))}</h3><p>${x('Este documento todavía no tiene archivos publicados.','This document has no published files yet.')}</p><div class="sticky"><button class="g" onclick="X()">${t('cl')}</button></div></div>`);return}
  const by={};fs.forEach(f=>by[f.lang]=f);const L=by[vl]?vl:by[lang]?lang:(by.es?'es':fs[0].lang),f=by[L],e=fext(f.file_name);
  const langs=['es','en','ht','pt','de'].filter(l=>by[l]);
  M(`<div class="pvw"><h3 style="margin:0 0 6px">${esc(docTitle(d,L))}</h3><span class="tag">${esc(tx(d.law||''))}</span> <span class="tag">${FICO[e]||''} ${FEXT[e]||e.toUpperCase()}</span>
    <div class="lang" style="margin:10px 0">${langs.map(l=>`<button class="g ${l==L?'on':''}" onclick="viewer(${id},${ro?1:0},'${l}')">${LN[l]}</button>`).join('')}</div>
    ${L!=lang&&lang!='es'&&!by[lang]?`<p class="mut" style="margin:0 0 8px">${x('Este documento aún no está disponible en tu idioma.','This document is not yet available in your language.')}</p>`:''}
    <div id="pv" class="pv"><p class="mut" style="padding:16px">${x('Cargando documento…','Loading document…')}</p></div><div id="pvend" style="height:1px"></div>
    <div class="sticky" id="pvbar"><a class="btnl" href="${esc(f.url||'#')}" download="${esc(f.file_name)}" style="background:transparent;color:var(--blue);border:1px solid var(--blue)">⬇ ${FEXT[e]||e.toUpperCase()} · ${LN[L]}</a><button class="g" onclick="X()">${t('cl')}</button>${ro?'':`<button id="pvnx" disabled onclick="nextR(${id})">${hasQ(id)?t('test'):t('sign')} →</button>`}</div>
    ${ro?'':`<p class="mut" id="pvhint" style="margin:8px 0 0">${x('Llega al final del documento para continuar.','Reach the end of the document to continue.')}</p>`}</div>`);
  const box=$('#pv');try{
    if(e=='pdf')await pvPdf(box,f.url,tok);else if(e=='xlsx')await pvXlsx(box,f.url);else if(e=='docx')await pvDocx(box,f.url);else box.innerHTML=`<p class="mut">${x('Vista previa no disponible; descarga el archivo.','Preview not available; download the file.')}</p>`}
  catch(err){if(PV===tok)box.innerHTML=`<p class="mut" style="padding:16px">${x('No se pudo mostrar la vista previa. Puedes descargar el archivo.','Could not show the preview. You can download the file.')}</p>`}
  if(PV!==tok||ro)return;const nb=$('#pvnx'),end=$('#pvend');if(!nb||!end)return;
  window.__pvio&&window.__pvio.disconnect();window.__pvio=new IntersectionObserver(en=>{if(en[0].isIntersecting){nb.disabled=false;const h=$('#pvhint');if(h)h.remove();window.__pvio.disconnect()}},{root:$('#mc')});window.__pvio.observe(end)}
async function pvPdf(box,url,tok){await need(PDFJS);pdfjsLib.GlobalWorkerOptions.workerSrc=PDFW;const pdf=await pdfjsLib.getDocument({url}).promise;if(PV!==tok)return;
  const p1=await pdf.getPage(1),v1=p1.getViewport({scale:1}),W=box.clientWidth||700,ratio=v1.height/v1.width;box.innerHTML='';
  const io=new IntersectionObserver(async en=>{for(const it of en){if(!it.isIntersecting||it.target.dataset.done)continue;it.target.dataset.done=1;const n=+it.target.dataset.p,pg=await pdf.getPage(n),vp0=pg.getViewport({scale:1}),sc=(box.clientWidth||W)/vp0.width*Math.min(2,window.devicePixelRatio||1),vp=pg.getViewport({scale:sc}),c=document.createElement('canvas');c.width=vp.width;c.height=vp.height;it.target.innerHTML='';it.target.appendChild(c);await pg.render({canvasContext:c.getContext('2d'),viewport:vp}).promise}},{root:$('#mc'),rootMargin:'800px 0px'});
  for(let n=1;n<=pdf.numPages;n++){const pd=document.createElement('div');pd.className='pvpage';pd.dataset.p=n;pd.style.aspectRatio=(1/ratio).toFixed(4);pd.innerHTML=`<span class="mut">${n}/${pdf.numPages}</span>`;box.appendChild(pd);io.observe(pd)}}
async function pvXlsx(box,url){await need(XLSXJS);const wb=XLSX.read(await (await fetch(url)).arrayBuffer(),{type:'array'});
  const show=k=>{box.querySelectorAll('.pvtab').forEach((b,j)=>b.classList.toggle('on',j==k));box.querySelector('.pvsheet').innerHTML=XLSX.utils.sheet_to_html(wb.Sheets[wb.SheetNames[k]],{editable:false})};
  box.innerHTML=`<div class="lang" style="margin:0 0 8px">${wb.SheetNames.map((n,k)=>`<button class="g pvtab" data-k="${k}">${esc(n)}</button>`).join('')}</div><div class="pvsheet pvx"></div><p class="mut" style="margin-top:8px">${x('Vista simplificada: colores y formato completos en el archivo descargado.','Simplified view: full colours and formatting in the downloaded file.')}</p>`;
  box.querySelectorAll('.pvtab').forEach(b=>b.onclick=()=>show(+b.dataset.k));show(0)}
async function pvDocx(box,url){await need(JSZIPJS);await need(DOCXJS);const blob=await (await fetch(url)).blob();box.innerHTML='';box.classList.add('pvd');await docx.renderAsync(blob,box,null,{inWrapper:true,breakPages:true,ignoreLastRenderedPageBreak:true})}
reader=function(i,ro,vl){return viewer(i&&i.id!=null?i.id:i,ro,vl)};

/* ---------- firma con mouse (PC) o dedo (celular), después de aprobar la evaluación ---------- */
const SLOC=()=>({es:'es-CL',en:'en-US',ht:'fr-HT',pt:'pt-BR',de:'de-DE'})[lang]||'es-CL';
function sigPad(c,onInk){
  const g=c.getContext('2d');let dpr=1,down=false,last=null,mid=null,ink=0;
  const fit=()=>{dpr=Math.max(1,window.devicePixelRatio||1);const w=c.clientWidth||560,h=c.clientHeight||200;c.width=Math.round(w*dpr);c.height=Math.round(h*dpr);g.setTransform(dpr,0,0,dpr,0,0);g.lineCap='round';g.lineJoin='round';g.strokeStyle=g.fillStyle='#0E2038'};
  fit();const onR=()=>{if(!c.isConnected){removeEventListener('resize',onR);return}if(!ink)fit()};addEventListener('resize',onR);
  const pt=e=>{const r=c.getBoundingClientRect();return{x:e.clientX-r.left,y:e.clientY-r.top,p:e.pointerType=='mouse'||!e.pressure?.5:e.pressure}};
  c.onpointerdown=e=>{e.preventDefault();try{c.setPointerCapture(e.pointerId)}catch(_){}down=true;last=mid=pt(e);g.beginPath();g.arc(last.x,last.y,1.4,0,7);g.fill();ink+=2;onInk(ink)};
  c.onpointermove=e=>{if(!down)return;e.preventDefault();const evs=e.getCoalescedEvents&&e.getCoalescedEvents().length?e.getCoalescedEvents():[e];
    for(const ev of evs){const p=pt(ev),m={x:(last.x+p.x)/2,y:(last.y+p.y)/2};g.beginPath();g.lineWidth=1.8+p.p*1.6;g.moveTo(mid.x,mid.y);g.quadraticCurveTo(last.x,last.y,m.x,m.y);g.stroke();ink+=Math.hypot(p.x-last.x,p.y-last.y);last=p;mid=m}onInk(ink)};
  const up=()=>{if(!down)return;down=false;g.beginPath();g.moveTo(mid.x,mid.y);g.lineTo(last.x,last.y);g.stroke()};
  c.onpointerup=c.onpointercancel=c.onlostpointercapture=up;
  return{
    clear(){g.save();g.setTransform(1,0,0,1,0,0);g.clearRect(0,0,c.width,c.height);g.restore();ink=0;onInk(0)},
    png(){/* recorta al trazo y reduce a 600 px de ancho como máximo */
      const W=c.width,H=c.height,d=g.getImageData(0,0,W,H).data;let x0=W,y0=H,x1=-1,y1=-1;
      for(let y=0;y<H;y+=2)for(let x=0;x<W;x+=2)if(d[(y*W+x)*4+3]>16){if(x<x0)x0=x;if(x>x1)x1=x;if(y<y0)y0=y;if(y>y1)y1=y}
      if(x1<0)return'';const pad=8*dpr;x0=Math.max(0,x0-pad);y0=Math.max(0,y0-pad);x1=Math.min(W,x1+pad);y1=Math.min(H,y1+pad);
      const w=x1-x0,h=y1-y0,s=Math.min(1,600/w),o=document.createElement('canvas');o.width=Math.round(w*s);o.height=Math.round(h*s);
      o.getContext('2d').drawImage(c,x0,y0,w,h,0,0,o.width,o.height);return o.toDataURL('image/png')}
  }
}
function sig(i,sc){
  if(!i)return;if(hasQ(i.id)&&!(sc>=PASS)){quiz(i.id);return}
  M(`<h3>${esc(T4(i.t))}</h3>${sc!=null?`<p><span class="tag ok">${t('ok')} ${t('sc')}: ${sc}%</span></p>`:''}
  <p style="margin:8px 0">${x('Declaro haber leído y comprendido este documento.','I declare that I have read and understood this document.')}</p>
  <div class="sgw"><canvas id="cv"></canvas><span class="sgl" id="sgl">${t('sg')}</span></div>
  <p class="mut" style="font-size:.85em;margin:6px 0 12px">${esc(user||'')} · ${new Date().toLocaleDateString(SLOC(),{day:'numeric',month:'long',year:'numeric'})}</p>
  <div style="display:flex;gap:8px;flex-wrap:wrap"><button class="g" id="cc">${t('clr')}</button><button id="cs" disabled>${t('cf')}</button></div>`);
  const cs=$('#cs'),lb=$('#sgl');
  const pad=sigPad($('#cv'),ink=>{cs.disabled=ink<60;lb.style.opacity=ink?0:1});
  $('#cc').onclick=()=>pad.clear();
  cs.onclick=async()=>{const png=pad.png();if(!png)return;cs.disabled=true;
    st[i.id]={s:1,sc,sg:png,at:Date.now(),lg:lang};setProg(i.id,3,sc);
    clearTimeout(snapT);const sn=snap();PR[ukey()]=sn;await put('prog',ukey(),sn);
    doneModal(i.id);toast('✓ '+t('sd'));if(typeof render=='function')render()};
}
function sigCard(id){const s=st[id];if(!s||!s.s)return'';const d=s.at?new Date(s.at).toLocaleString(SLOC(),{day:'numeric',month:'long',year:'numeric',hour:'2-digit',minute:'2-digit'}):'';
  return`<div class="sgc">${s.sg?`<img src="${s.sg}" alt="${esc(t('sd'))}">`:''}<div class="mut">${esc(user||'')}${d?' · '+d:''}${s.sc!=null?' · '+t('sc')+': '+s.sc+'%':''}</div></div>`}
function doneModal(id){const i=items().find(z=>z.id==id)||{t:{es:''}};M(`<h3>✓ ${t('sd')}</h3><p><b>${esc(T4(i.t))}</b></p>${sigCard(id)}<p class="mut">${x('Puedes descargar el documento en su formato original, en los idiomas disponibles.','You can download the document in its original format, in the available languages.')}</p>${dlBtns(id)}<p></p><button onclick="X()">${t('cl')}</button>`);fillDl(id)}
{const _open0=open_;open_=function(id){if(st[id]&&st[id].s){const i=items().find(z=>z.id==id);if(!i)return;
  M(`<h3>${esc(T4(i.t))}</h3><p><span class="tag ok">✓ ${t('sd')}</span></p>${sigCard(id)}<div style="margin-top:12px">${dlBtns(id)}</div><p></p><div style="display:flex;gap:8px;flex-wrap:wrap"><button onclick="viewer(${id},1)">${t('open')}</button><button class="g" onclick="X()">${t('cl')}</button></div>`);fillDl(id);return}
  return _open0(id)}}


/* ===== v4: subida y gestión de documentos según el modelo de archivos (sin texto interno) ===== */
const TL=['en','ht','pt','de'],LAWS=['DS 44','Ley 16.744','DS 594','Código del Trabajo','Ley 21.643','Contrato'];
function tLangs(id){const d=dm(id)||{},f=Object.keys(FLG[id]||{}).filter(l=>l!='es'),s=new Set([...(Array.isArray(d.ln)&&d.ln.length?d.ln:(f.length?f:TL)),...f]);return TL.filter(l=>s.has(l))}
const gOf=d=>GR.includes(d.g)?d.g:'Otros documentos';
const tTitle=(d,l)=>!!(TR[l]&&TR[l][d.title]&&TR[l][d.title]!=d.title);
const lawDL=()=>`<datalist id="lawl">${LAWS.map(l=>`<option value="${esc(l)}">`).join('')}</datalist>`;
const areaTxt=a=>a&&a!='all'&&H[a]?P(H[a].n):x('Todas las áreas','All areas');
const UL=a=>a.map(l=>l.toUpperCase()).join(', ');
async function upFile2(id,l,file,kind){const j=await upDirect(id,l,file,kind||'translation','published','manual');loadFL();return j}
async function putOk(c,k,d){const r=await fetch('/api/data/'+encodeURIComponent(c)+'/'+encodeURIComponent(k),{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify(d)});if(!r.ok){const j=await r.json().catch(()=>({}));throw new Error(j.error||'HTTP '+r.status)}}
async function saveTitles(title,map){for(const l in map){const v=map[l];if(!v)continue;(TR[l]||(TR[l]={}))[title]=v;await putOk('ui',l+'_t',{l,k:[title],v:[v]})}}

/* estado de cada documento: qué le falta para quedar completo */
function dStat(d){const o=FLG[d.id]||{},L=tLangs(d.id),es=QES(d.id),r=[];
  if(!o.es)r.push(x('Falta el archivo original','Original file missing'));
  const mf=L.filter(l=>!o[l]);if(mf.length)r.push(x('Falta el archivo en','File missing in')+' '+UL(mf));
  const rv=['es',...L].filter(l=>o[l]&&o[l]!='published');if(rv.length)r.push(x('Archivo en revisión','File in review')+': '+UL(rv));
  if(es.length){const mq=L.filter(l=>qzState(d.id,l)!='ok');if(mq.length)r.push(x('Preguntas sin traducir','Untranslated questions')+': '+UL(mq))}
    return r}

/* ---------- Subir material ---------- */
V.sub=function(){return`<div class="card upw">
<p class="mut" style="margin-top:0">${x('Sube el documento en español y elige los idiomas de tus trabajadores. El archivo se muestra y se descarga tal como lo subes. Las traducciones del archivo puedes subirlas ahora o después.','Upload the document in Spanish and choose your workers\' languages. The file is shown and downloaded exactly as you upload it. You can upload the file translations now or later.')}</p>
<h4>1. ${x('Archivo original en español','Original file in Spanish')}</h4>
<input type="file" id="u_f" accept=".docx,.pdf,.xlsx" onchange="upPick(1)">
<p class="mut" id="u_fi" style="margin:6px 0 0">${x('Word (.docx), PDF o Excel (.xlsx), hasta 50 MB.','Word (.docx), PDF or Excel (.xlsx), up to 50 MB.')}</p>
<h4>2. ${x('Datos del documento','Document details')}</h4>
<label>${t('tt')}</label><input id="u_t">
<div class="grid g2" style="gap:0 14px"><div><label>${t('lw')}</label><input id="u_law" list="lawl" value="DS 44">${lawDL()}</div>
<div><label>${x('Grupo','Group')}</label><select id="u_g">${GR.map(g=>`<option value="${esc(g)}">${esc(tx(g))}</option>`).join('')}</select></div></div>
<label>${x('Dirigido a','Audience')}</label><select id="u_a"><option value="all">${x('Todas las áreas','All areas')}</option>${AO('')}</select>
<h4>3. ${x('Idiomas de tus trabajadores','Your workers\' languages')}</h4>
<p class="mut" style="margin:0 0 8px">${x('Marca los idiomas que necesitas. Para cada uno puedes escribir el título traducido y subir el archivo traducido, en el mismo formato que el original.','Tick the languages you need. For each one you can enter the translated title and upload the translated file, in the same format as the original.')}</p>
${TL.map(l=>`<div class="upl" id="u_r_${l}"><label class="upc"><input type="checkbox" id="u_c_${l}" checked onchange="upPick()"> <b>${LN[l]}</b></label>
<div class="upd"><span class="upf"><label class="btnl g">⬆ ${x('Archivo traducido','Translated file')}<input type="file" id="u_tf_${l}" style="display:none" onchange="upPick()"></label> <span class="mut" id="u_tn_${l}">${x('opcional','optional')}</span></span></div></div>`).join('')}
<h4>4. ${x('Preguntas de evaluación','Test questions')}</h4>
<p class="mut" style="margin:0 0 12px">${x('Al guardar, escribes las preguntas. El documento no se muestra a los trabajadores hasta que tenga preguntas.','After saving, you write the questions. The document is not shown to workers until it has questions.')}</p>
<button id="pubb" onclick="pub()">${x('Guardar documento','Save document')}</button></div>`};
function upPick(first){const f=$('#u_f')&&$('#u_f').files[0],e=f?fext(f.name):'';
  if(f){if(first&&!$('#u_t').value.trim())$('#u_t').value=f.name.replace(/\.[^.]+$/,'').replace(/[_]+/g,' ').trim();
    $('#u_fi').innerHTML=FEXT[e]?`${FICO[e]||''} <b>${esc(f.name)}</b> · ${FEXT[e]} · ${(f.size/1048576).toFixed(1)} MB`:`<span style="color:var(--red)">${x('Formato no permitido. Solo Word (.docx), PDF o Excel (.xlsx).','Format not allowed. Only Word (.docx), PDF or Excel (.xlsx).')}</span>`}
  TL.forEach(l=>{const on=$('#u_c_'+l).checked,tf=$('#u_tf_'+l),tn=$('#u_tn_'+l),g=tf.files[0];$('#u_r_'+l).classList.toggle('off',!on);if(e)tf.accept='.'+e;
    tn.innerHTML=!g?x('opcional','optional'):(e&&fext(g.name)!=e)?`<span style="color:var(--red)">${esc(g.name)} · ${x('debe ser','must be')} ${FEXT[e]}</span>`:`✓ ${esc(g.name)}`})}
async function pub(){const f=$('#u_f').files[0];if(!f){toast(x('Elige el archivo original en español','Choose the original Spanish file'));return}
  const e=fext(f.name);if(!FEXT[e]){toast(x('Formato no permitido. Solo Word (.docx), PDF o Excel (.xlsx).','Format not allowed. Only Word (.docx), PDF or Excel (.xlsx).'));return}
  const title=$('#u_t').value.trim();if(!title){toast(x('Escribe un título','Enter a title'));return}
  if(DOCS.some(d=>d.title.trim().toLowerCase()==title.toLowerCase())){toast(x('Ya existe un documento con ese título','A document with that title already exists'));return}
  const law=$('#u_law').value.trim(),g=$('#u_g').value,area=$('#u_a').value,ln=TL.filter(l=>$('#u_c_'+l).checked);
  const trs=ln.map(l=>({l,f:$('#u_tf_'+l).files[0],tt:(($('#u_tt_'+l)||{}).value||'').trim()}));
  const bad=trs.find(z=>z.f&&fext(z.f.name)!=e);if(bad){toast(x('El archivo en','The file in')+' '+LN[bad.l]+' '+x('debe ser','must be')+' '+FEXT[e]);return}
  M(`<h3>${x('Guardando documento…','Saving document…')}</h3><p class="mut" id="pst"></p><div class="bar"><i id="pbar" style="width:5%"></i></div>`);
  const steps=2+trs.filter(z=>z.f).length+1;let n=0;const step=s=>{$('#pst').textContent=s;$('#pbar').style.width=Math.round(++n/steps*100)+'%'};
  let id=null;
  try{id=Math.max(1000,...DOCS.map(d=>+d.id||0))+1;const m={id,title,law,cat:'doc',area,g,ln,nw:1,order:[]};
    step(x('Creando el documento…','Creating the document…'));await putOk('docs','d'+id,m);MM['d'+id]=m;build();
    step(x('Subiendo el archivo original…','Uploading the original file…'));await upFile2(id,'es',f,'original');
    for(const z of trs)if(z.f){step(x('Subiendo el archivo en','Uploading the file in')+' '+LN[z.l]+'…');await upFile2(id,z.l,z.f,'translation')}
    step(x('Guardando los títulos…','Saving the titles…'));const map={};trs.forEach(z=>{if(z.tt)map[z.l]=z.tt});await saveTitles(title,map);
    await loadFL();upDone(id)}
  catch(err){if(id&&MM['d'+id]&&!(FLG[id]||{}).es){try{await fetch('/api/docs/'+id,{method:'DELETE'})}catch(_){}delete MM['d'+id];build()}
    M(`<h3>${x('No se pudo guardar','Could not save')}</h3><p>${esc(err.message||String(err))}</p><p class="mut">${x('Revisa que el servidor esté encendido e inténtalo de nuevo.','Check that the server is running and try again.')}</p><button class="g" onclick="X()">${t('cl')}</button>`)}}
function upDone(id){const d=dm(id),r=dStat(d);M(`<h3>✓ ${x('Documento guardado','Document saved')}</h3><p><b>${esc(docTitle(d))}</b></p>
  <div class="card" style="padding:12px 16px;margin:10px 0"><b>${x('Pendiente','Pending')}</b><ul style="margin:6px 0 0;padding-left:18px">${[x('Escribir las preguntas de evaluación (sin ellas, los trabajadores no ven el documento)','Write the test questions (without them, workers cannot see the document)'),...r].map(z=>`<li>${esc(z)}</li>`).join('')}</ul></div>
  <div style="display:flex;gap:8px;flex-wrap:wrap"><button onclick="edQz(${id})">❓ ${x('Escribir preguntas','Write questions')}</button><button class="g" onclick="fmOpen(${id})">📁 ${x('Archivos por idioma','Files by language')}</button><button class="g" onclick="X();tab='dx';render()">${t('cl')}</button></div>`)}

/* ---------- lista de documentos ---------- */
function newDoc(){tab='sub';render()}
V.dx=function(){const row=d=>{const r=dStat(d),q=QES(d.id).length,L=tLangs(d.id);return`<div class="row"><div style="min-width:0;flex:1"><b>${esc(docTitle(d))}</b><br>
<span class="tag">${esc(tx(d.law||'—'))}</span> <span class="mut" style="font-size:.85em">${esc(areaTxt(d.area))}</span>
<div style="margin-top:6px;display:flex;gap:4px;flex-wrap:wrap;align-items:center">${flTags(d.id)}${L.filter(l=>!(FLG[d.id]||{})[l]).map(l=>`<span class="tag" style="opacity:.45" title="${x('Falta el archivo','File missing')}">${l.toUpperCase()}</span>`).join('')}</div>
<div style="margin-top:6px;font-size:.85em">${!q?`<span style="color:var(--red)">● ${x('Sin preguntas: oculto para trabajadores','No questions: hidden from workers')}</span>`:r.length?`<span style="color:var(--amber,#d9a02b)">● ${x('Visible para trabajadores','Visible to workers')}</span> <span class="mut">· ${esc(r.join(' · '))}</span>`:`<span style="color:var(--green,#2fa36b)">● ${x('Completo: visible para trabajadores en','Complete: visible to workers in')} ${L.length+1} ${x('idiomas','languages')}</span>`}</div></div>
<span class="dxb"><button class="g" onclick="viewer(${d.id},1)">${t('open')}</button><button class="g" onclick="fmOpen(${d.id})" title="${x('Archivos por idioma','Files by language')}">📁</button><button class="g" onclick="edQz(${d.id})" title="${x('Preguntas de evaluación','Test questions')}">❓ ${q}</button><button onclick="edDoc(${d.id})" title="${x('Datos del documento','Document details')}">✎</button><button class="g" onclick="askDelDoc(${d.id})" title="${x('Eliminar','Delete')}">✕</button></span></div>`};
  const by=GR.map(g=>[g,DOCS.filter(d=>gOf(d)==g)]).filter(z=>z[1].length);
  return`<div style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap"><p class="mut" style="margin:0;max-width:640px">${x('Un documento se muestra a los trabajadores cuando tiene preguntas. Aquí ves qué le falta a cada uno: archivos por idioma, preguntas y títulos traducidos.','A document is shown to workers once it has questions. Here you can see what each one is missing: files by language, questions and translated titles.')}</p><button onclick="tab='sub';render()">⬆ ${x('Subir documento','Upload document')}</button></div>
  ${by.length?by.map(([g,a])=>`<h4 class="gh">${esc(tx(g))} · ${a.length}</h4><div class="card" style="padding:4px 16px">${a.map(row).join('')}</div>`).join(''):`<div class="card" style="margin-top:16px"><p class="mut" style="margin:0">${x('Aún no hay documentos. Súbelos en "Subir material".','There are no documents yet. Upload them in "Upload material".')}</p></div>`}`};

/* ---------- datos del documento ---------- */
function edDoc(id){const d=dm(id);if(!d)return X();const L=tLangs(id);
  M(`<h3>✎ ${x('Datos del documento','Document details')}</h3>
  <label>${t('tt')} (${LN.es})</label><input id="e1" value="${esc(d.title)}">
  <div class="grid g2" style="gap:0 14px"><div><label>${t('lw')}</label><input id="e2" list="lawl" value="${esc(d.law||'')}">${lawDL()}</div>
  <div><label>${x('Grupo','Group')}</label><select id="e3">${GR.map(g=>`<option value="${esc(g)}" ${gOf(d)==g?'selected':''}>${esc(tx(g))}</option>`).join('')}</select></div></div>
  <label>${x('Dirigido a','Audience')}</label><select id="e4"><option value="all">${x('Todas las áreas','All areas')}</option>${AO(d.area)}</select>
  <label>${x('Idiomas de los trabajadores','Workers\' languages')}</label>
  ${TL.map(l=>`<div class="upl"><label class="upc"><input type="checkbox" id="e_c_${l}" ${L.includes(l)?'checked':''}> <b>${LN[l]}</b></label></div>`).join('')}
  <p class="mut" style="font-size:.85em">${x('Quitar un idioma no borra sus archivos. Para reemplazar o quitar archivos usa 📁.','Removing a language does not delete its files. Use 📁 to replace or remove files.')}</p>
  <div style="display:flex;gap:8px;flex-wrap:wrap"><button onclick="saveMeta(${id})">${x('Guardar datos','Save details')}</button><button class="g" onclick="X()">${t('cl')}</button></div>`)}
async function saveMeta(id){const d=dm(id),title=$('#e1').value.trim()||d.title,ln=TL.filter(l=>$('#e_c_'+l).checked),map={};TL.forEach(l=>{const v=(($('#e_t_'+l)||{}).value||'').trim();if(v)map[l]=v});
  const m={...(MM['d'+id]||{}),id,title,law:$('#e2').value.trim(),cat:'doc',area:$('#e4').value,g:$('#e3').value,ln,order:(MM['d'+id]||{}).order||(d.secs||[]).map(z=>z.id)};
  try{await putOk('docs','d'+id,m);MM['d'+id]=m;build();await saveTitles(title,map);X();render();toast('✓ '+x('Datos guardados','Details saved'))}catch(e){toast(x('No se pudo guardar: ','Could not save: ')+e.message)}}

/* ===== v5 (PRESENTACIÓN): traducción automática simulada al subir =====
   Cambia SIM_TR a false para volver a la subida manual (v4).
   Si para un idioma se adjunta una traducción propia, se usa esa. Si no, se sube una copia del original
   marcada con engine='simulado' (se ve como "simulada" en 📁 para reemplazarla después). */
const SIM_TR=true;
if(SIM_TR){
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
V.sub=function(){return`<div class="card upw">
<p class="mut" style="margin-top:0">${x('Sube el documento en español y elige a qué idiomas traducirlo. La plataforma lo traduce con IA manteniendo el formato original y lo publica para tus trabajadores.','Upload the document in Spanish and choose which languages to translate it into. The platform translates it with AI, keeps the original format and publishes it for your workers.')}</p>
<h4>1. ${x('Archivo original en español','Original file in Spanish')}</h4>
<input type="file" id="u_f" accept=".docx,.pdf,.xlsx" onchange="upPick(1)">
<p class="mut" id="u_fi" style="margin:6px 0 0">${x('Word (.docx), PDF o Excel (.xlsx), hasta 50 MB.','Word (.docx), PDF or Excel (.xlsx), up to 50 MB.')}</p>
<h4>2. ${x('Datos del documento','Document details')}</h4>
<label>${t('tt')}</label><input id="u_t">
<div class="grid g2" style="gap:0 14px"><div><label>${t('lw')}</label><input id="u_law" list="lawl" value="DS 44">${lawDL()}</div>
<div><label>${x('Grupo','Group')}</label><select id="u_g">${GR.map(g=>`<option value="${esc(g)}">${esc(tx(g))}</option>`).join('')}</select></div></div>
<label>${x('Dirigido a','Audience')}</label><select id="u_a"><option value="all">${x('Todas las áreas','All areas')}</option>${AO('')}</select>
<h4>3. ${x('Traducción automática','Automatic translation')}</h4>
<p class="mut" style="margin:0 0 10px">${x('Elige los idiomas de tus trabajadores. Cada traducción conserva el formato, las tablas y las imágenes del original.','Choose your workers\' languages. Each translation keeps the format, tables and images of the original.')}</p>
<div class="trc">${TL.map(l=>`<label class="trp"><input type="checkbox" id="u_c_${l}" checked onchange="upPick()"><span>${LN[l]}</span></label>`).join('')}</div>
<details class="upx"><summary>${x('Traducciones propias (opcional)','Your own translations (optional)')}</summary>
<p class="mut" style="margin:6px 0">${x('Si ya tienes una traducción revisada, súbela aquí y se usará en lugar de la automática.','If you already have a reviewed translation, upload it here and it will be used instead of the automatic one.')}</p>
${TL.map(l=>`<div class="upl" id="u_r_${l}"><span class="upc"><b>${LN[l]}</b></span><div class="upd"><span class="upf"><label class="btnl g">⬆ ${x('Archivo traducido','Translated file')}<input type="file" id="u_tf_${l}" style="display:none" onchange="upPick()"></label> <span class="mut" id="u_tn_${l}">${x('opcional','optional')}</span></span></div></div>`).join('')}</details>
<h4>4. ${x('Preguntas de evaluación','Test questions')}</h4>
<p class="mut" style="margin:0 0 12px">${x('Al terminar la traducción, escribes las preguntas. El documento no se muestra a los trabajadores hasta que tenga preguntas.','When the translation is done, you write the questions. The document is not shown to workers until it has questions.')}</p>
<button id="pubb" onclick="pub()">🌐 ${x('Traducir y publicar','Translate and publish')}</button></div>`};
pub=async function(){const f=$('#u_f').files[0];if(!f){toast(x('Elige el archivo original en español','Choose the original Spanish file'));return}
  const e=fext(f.name);if(!FEXT[e]){toast(x('Formato no permitido. Solo Word (.docx), PDF o Excel (.xlsx).','Format not allowed. Only Word (.docx), PDF or Excel (.xlsx).'));return}
  const title=$('#u_t').value.trim();if(!title){toast(x('Escribe un título','Enter a title'));return}
  if(DOCS.some(d=>d.title.trim().toLowerCase()==title.toLowerCase())){toast(x('Ya existe un documento con ese título','A document with that title already exists'));return}
  const law=$('#u_law').value.trim(),g=$('#u_g').value,area=$('#u_a').value,ln=TL.filter(l=>$('#u_c_'+l).checked);
  const trs=ln.map(l=>({l,f:$('#u_tf_'+l).files[0],tt:(($('#u_tt_'+l)||{}).value||'').trim()}));
  const bad=trs.find(z=>z.f&&fext(z.f.name)!=e);if(bad){toast(x('El archivo en','The file in')+' '+LN[bad.l]+' '+x('debe ser','must be')+' '+FEXT[e]);return}
  const li=(k,txt)=>`<div class="trs" id="tr_${k}"><span class="tri">○</span><span style="flex:1">${txt}</span><span class="mut trm"></span></div>`;
  M(`<h3>🌐 ${x('Traduciendo documento','Translating document')}</h3><p><b>${esc(title)}</b> <span class="tag">${FICO[e]||''} ${FEXT[e]}</span></p>
  <div class="bar"><i id="pbar" style="width:3%"></i></div><div style="margin-top:12px">${li('es',x('Subir el archivo original','Upload the original file'))}${ln.map(l=>li(l,x('Traducir a','Translate into')+' '+LN[l])).join('')}${li('pub',x('Publicar para los trabajadores','Publish for workers'))}</div>`);
  const total=2+ln.length;let n=0;const bar=()=>{$('#pbar').style.width=Math.round(++n/total*100)+'%'};
  const st=(k,s,m)=>{const r=$('#tr_'+k);if(!r)return;r.className='trs '+s;r.querySelector('.tri').textContent=s=='on'?'◐':s=='ok'?'✓':s=='bad'?'✕':'○';r.querySelector('.trm').textContent=m||''};
  let id=null;
  try{id=Math.max(1000,...DOCS.map(d=>+d.id||0))+1;const m={id,title,law,cat:'doc',area,g,ln,nw:1,order:[]};
    st('es','on',x('subiendo…','uploading…'));await putOk('docs','d'+id,m);MM['d'+id]=m;build();await upFile2(id,'es',f,'original');st('es','ok');bar();
    const titles={};
    for(const z of trs){st(z.l,'on',x('traduciendo…','translating…'));const t0=Date.now(),dur=1800+Math.random()*1600,
        tick=setInterval(()=>{const p=Math.min(99,Math.round((Date.now()-t0)/dur*100));const r=$('#tr_'+z.l);if(r)r.querySelector('.trm').textContent=p+' %'},120);
      try{const file=z.f||new File([f],f.name.replace(/\.[^.]+$/,'')+'_'+z.l.toUpperCase()+'.'+e,{type:f.type});
        const up=upDirect(id,z.l,file,'translation','published',z.f?'manual':'simulado');
        if(z.tt)titles[z.l]=z.tt;
        await Promise.all([up,sleep(Math.max(0,dur-(Date.now()-t0)))]);clearInterval(tick);st(z.l,'ok',x('listo','done'))}
      catch(err){clearInterval(tick);st(z.l,'bad',err.message)}bar()}
    st('pub','on');await saveTitles(title,titles);await loadFL();await sleep(500);st('pub','ok');bar();await sleep(600);upDone(id)}
  catch(err){if(id&&MM['d'+id]&&!(FLG[id]||{}).es){try{await fetch('/api/docs/'+id,{method:'DELETE'})}catch(_){}delete MM['d'+id];build()}
    M(`<h3>${x('No se pudo guardar','Could not save')}</h3><p>${esc(err.message||String(err))}</p><p class="mut">${x('Revisa que el servidor esté encendido e inténtalo de nuevo.','Check that the server is running and try again.')}</p><button class="g" onclick="X()">${t('cl')}</button>`)}};
upDone=function(id){const d=dm(id),r=dStat(d).filter(z=>!/^(Falta el archivo en|File missing in|Título sin traducir|Untranslated title)/.test(z)),L=tLangs(id);M(`<h3>✓ ${x('Documento traducido y publicado','Document translated and published')}</h3><p><b>${esc(docTitle(d))}</b></p>
  <p>${['es',...L].filter(l=>(FLG[id]||{})[l]).map(l=>`<span class="tag ok">${l.toUpperCase()} ✓</span>`).join(' ')}</p>
  <div class="card" style="padding:12px 16px;margin:10px 0"><b>${x('Siguiente paso','Next step')}</b><ul style="margin:6px 0 0;padding-left:18px">${[x('Escribir las preguntas de evaluación (sin ellas, los trabajadores no ven el documento)','Write the test questions (without them, workers cannot see the document)'),...r].map(z=>`<li>${esc(z)}</li>`).join('')}</ul></div>
  <div style="display:flex;gap:8px;flex-wrap:wrap"><button onclick="edQz(${id})">❓ ${x('Escribir preguntas','Write questions')}</button><button class="g" onclick="viewer(${id},1)">${t('open')}</button><button class="g" onclick="X();tab='dx';render()">${t('cl')}</button></div>`)};
/* en 📁 se marca la copia simulada para reemplazarla después (solo prevencionista) */
{const fo=fmOpen;fmOpen=async function(id){await fo(id);try{const fs=await filesOf(id);fs.filter(f=>f.engine=='simulado').forEach(f=>{const a=[...document.querySelectorAll('#mc a[download]')].find(z=>z.getAttribute('download')==f.file_name);if(a)a.insertAdjacentHTML('afterend',` <span class="tag go" title="${x('Copia del original: reemplázala por la traducción real','Copy of the original: replace it with the real translation')}">${x('simulada','simulated')}</span>`)})}catch(e){}}}
}


/* ===== v6 (PRESENTACIÓN): confirmar antes de borrar una pregunta y traducir preguntas a todos los idiomas (simulado) ===== */
{const dq0=delQ;delQ=function(id,k){const q=QES(id)[k];if(!q)return;const n=QES(id).length;
  M(`<h3>${x('¿Eliminar esta pregunta?','Delete this question?')}</h3><div class="card" style="padding:12px 16px;margin:10px 0"><b>${k+1}. ${esc(q.q)}</b><ul style="margin:6px 0 0;padding-left:18px">${(q.o||[]).map((o,j)=>`<li>${j==(q.c||0)?'✔ ':''}${esc(o)}</li>`).join('')}</ul></div>
  <p class="mut">${x('Se eliminará también en todos los idiomas. Esta acción no se puede deshacer.','It will also be deleted in all languages. This cannot be undone.')}${n<=1?' <b>'+x('Es la última pregunta: el documento dejará de ser visible para los trabajadores.','It is the last question: the document will no longer be visible to workers.')+'</b>':''}</p>
  <div style="display:flex;gap:8px;flex-wrap:wrap"><button style="background:var(--red)" onclick="delQok(${id},${k})">${x('Sí, eliminar','Yes, delete')}</button><button class="g" onclick="edQz(${id})">${x('Cancelar','Cancel')}</button></div>`)};
  window.delQok=function(id,k){dq0(id,k);toast('✓ '+x('Pregunta eliminada','Question deleted'))}}
async function trAllQ(id){const es=QES(id);if(!es.length){toast(x('Escribe al menos una pregunta','Write at least one question'));return}
  const L=tLangs(id),li=l=>`<div class="trs" id="tq_${l}"><span class="tri">○</span><span style="flex:1">${x('Traducir a','Translate into')} ${LN[l]}</span><span class="mut trm"></span></div>`;
  M(`<h3>🌐 ${x('Traduciendo preguntas','Translating questions')}</h3><p class="mut">${esc(docTitle(dm(id)))} · ${es.length} ${x('preguntas','questions')}</p><div class="bar"><i id="tqbar" style="width:3%"></i></div><div style="margin-top:12px">${L.map(li).join('')}</div>`);
  const src=qkey(es);let k=0;
  for(const l of L){const r=$('#tq_'+l);r.className='trs on';r.querySelector('.tri').textContent='◐';const t0=Date.now(),dur=900+Math.random()*900;
    const tk=setInterval(()=>{r.querySelector('.trm').textContent=Math.min(99,Math.round((Date.now()-t0)/dur*100))+' %'},100);
    if(qzState(id,l)!='ok'){const rec={id,lang:l,qs:JSON.parse(JSON.stringify(es)),src,engine:'simulado'};QZT[id+'_'+l]=rec;await put('qzt',id+'_'+l,rec)}
    await pause(Math.max(0,dur-(Date.now()-t0)));clearInterval(tk);r.className='trs ok';r.querySelector('.tri').textContent='✓';r.querySelector('.trm').textContent=x('listo','done');
    $('#tqbar').style.width=Math.round(++k/L.length*100)+'%'}
  await pause(500);toast('✓ '+x('Preguntas traducidas y guardadas','Questions translated and saved')+' ('+L.length+')');edQz(id)}

/* ===== v7: sello del documento firmado, constancia en PDF y vista de firmas del prevencionista =====
   Nada de esto modifica los archivos guardados: el archivo solo se LEE para calcular su huella (SHA-256),
   y la constancia es un PDF nuevo que se genera aparte en el navegador. */
const JSPDF='https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
const hex=b=>[...new Uint8Array(b)].map(z=>z.toString(16).padStart(2,'0')).join('');
async function sealOf(id){/* archivo que el trabajador tiene en su idioma (o el original) y su huella */
  const fs=(await filesOf(id)).filter(f=>f.status=='published'),by={};fs.forEach(f=>by[f.lang]=f);const f=by[lang]||by.es;if(!f||!f.url)return null;
  const buf=await(await fetch(f.url)).arrayBuffer();return{fl:f.lang,fn:f.file_name,h:hex(await crypto.subtle.digest('SHA-256',buf))}}
{const sg0=sig;sig=function(i,sc){sg0(i,sc);if(!i||!$('#cs'))return;
  const seal=sealOf(i.id).catch(()=>null);
  $('#cv').closest('.sgw').insertAdjacentHTML('afterend',`<p class="mut" style="font-size:.78em;margin:6px 0 0">🔒 ${x('Se guardará un registro con tu firma, la fecha, tu puntaje y la huella digital del documento que leíste.','A record will be saved with your signature, the date, your score and the digital fingerprint of the document you read.')}</p>`);
  const cs=$('#cs'),go=cs.onclick;cs.onclick=async()=>{cs.disabled=true;const s=await Promise.race([seal,pause(8000).then(()=>null)]);await go();
    if(s&&st[i.id]&&st[i.id].s){Object.assign(st[i.id],s,{lg:s.fl});saveLS();clearTimeout(snapT);const sn=snap();PR[ukey()]=sn;put('prog',ukey(),sn,1)}}}}
/* constancia */
const whoName=u=>{const m=team.find(z=>z.u==u);return m?m.n:u};
async function constancia(id,u){const me=!u||u==ukey(),rec=me?st[id]:((PR[u]||{}).s||{})[id];if(!rec||!rec.s){toast(x('No hay firma registrada','No signature recorded'));return}
  try{await need(JSPDF)}catch(e){toast(x('No se pudo cargar el generador de PDF. Revisa tu conexión.','Could not load the PDF generator. Check your connection.'));return}
  const d=dm(id)||{title:'—',law:''},name=me?user:whoName(u),L=rec.lg||rec.fl||'es',{jsPDF}=window.jspdf,pdf=new jsPDF({unit:'mm',format:'a4'});
  const W=210,M=18;let y=20;const line=(k,v)=>{pdf.setFont('helvetica','bold');pdf.setFontSize(9);pdf.setTextColor(90);pdf.text(k,M,y);pdf.setFont('helvetica','normal');pdf.setFontSize(11);pdf.setTextColor(20);const t=pdf.splitTextToSize(String(v||'—'),W-M*2-48);pdf.text(t,M+48,y);y+=Math.max(7,t.length*5+2)};
  pdf.setFillColor(14,32,56);pdf.rect(0,0,W,30,'F');pdf.setTextColor(255);pdf.setFont('helvetica','bold');pdf.setFontSize(18);pdf.text('Prevent-IA',M,14);pdf.setFont('helvetica','normal');pdf.setFontSize(10);pdf.text('Constancia de lectura, evaluación y firma',M,22);
  y=44;pdf.setTextColor(20);pdf.setFont('helvetica','bold');pdf.setFontSize(13);const tt=pdf.splitTextToSize(d.title,W-M*2);pdf.text(tt,M,y);y+=tt.length*6+4;
  line('Trabajador',name);line('Usuario',me?ukey():u);line('Normativa',d.law||'—');line('Idioma leído',(LN[L]||L)+(L!='es'&&TR[L]&&TR[L][d.title]?' — '+TR[L][d.title]:''));
  line('Archivo leído',rec.fn||'—');line('Evaluación',rec.sc!=null?rec.sc+' % (aprobada, mínimo 80 %)':'Sin evaluación');line('Fecha y hora',rec.at?new Date(rec.at).toLocaleString('es-CL',{dateStyle:'long',timeStyle:'short'}):'—');
  y+=2;pdf.setFont('helvetica','bold');pdf.setFontSize(9);pdf.setTextColor(90);pdf.text('Declaración',M,y);y+=5;pdf.setFont('helvetica','normal');pdf.setFontSize(10.5);pdf.setTextColor(20);
  const dc=pdf.splitTextToSize('Declaro haber leído y comprendido este documento en el idioma indicado, y haber aprobado la evaluación correspondiente.',W-M*2);pdf.text(dc,M,y);y+=dc.length*5+6;
  pdf.setDrawColor(200);pdf.roundedRect(M,y,W-M*2,46,3,3);if(rec.sg){try{const im=new Image();im.src=rec.sg;await im.decode();const r=Math.min((W-M*2-10)/im.width,36/im.height),w=im.width*r,h=im.height*r;pdf.addImage(rec.sg,'PNG',(W-w)/2,y+5+(36-h)/2,w,h)}catch(e){}}
  pdf.setDrawColor(150);pdf.line(M+30,y+41,W-M-30,y+41);y+=51;pdf.setFontSize(9);pdf.setTextColor(90);pdf.text('Firma de '+name,W/2,y,{align:'center'});y+=12;
  pdf.setFont('helvetica','bold');pdf.text('Huella digital del documento (SHA-256)',M,y);y+=5;pdf.setFont('courier','normal');pdf.setFontSize(8.5);pdf.setTextColor(40);pdf.text(rec.h||'No registrada (firma anterior a esta función)',M,y);y+=5;
  pdf.setFont('helvetica','normal');pdf.setFontSize(8);pdf.setTextColor(110);pdf.text(pdf.splitTextToSize('La huella identifica el archivo exacto que se leyó: si el archivo cambia, la huella deja de coincidir. Registro generado por Prevent-IA el '+new Date().toLocaleString('es-CL')+'.',W-M*2),M,y);
  pdf.save(('Constancia_'+(d.title||'documento').slice(0,40)+'_'+name).normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/[^A-Za-z0-9]+/g,'_')+'.pdf')}
/* botón de constancia para el trabajador */
const cBtn=id=>`<button class="g" onclick="constancia(${id})">📄 ${x('Descargar constancia','Download certificate')}</button>`;
{const dm0=doneModal;doneModal=function(id){dm0(id);const c=[...document.querySelectorAll('#mc button')].find(z=>(z.getAttribute('onclick')||'')=='X()');if(c)c.insertAdjacentHTML('beforebegin',cBtn(id)+' ')}}
{const o0=open_;open_=function(id){o0(id);if(st[id]&&st[id].s){const b=[...document.querySelectorAll('#mc button')].find(z=>/viewer\(/.test(z.getAttribute('onclick')||''));if(b)b.insertAdjacentHTML('afterend',' '+cBtn(id))}}}
/* vista de firmas del prevencionista */
let FF={d:'all'};
function sigRows(){const rows=[];for(const u in PR){const s=(PR[u]||{}).s||{};for(const id in s){const r=s[id];if(r&&r.s&&dm(id))rows.push({u,id:+id,...r})}}return rows.sort((a,b)=>(b.at||0)-(a.at||0))}
V.fir=function(){const all=sigRows(),rows=all.filter(r=>FF.d=='all'||r.id==FF.d),ds=DOCS.filter(d=>all.some(r=>r.id==d.id));
  return`<div style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap"><p class="mut" style="margin:0;max-width:640px">${x('Registro de lectura y firma de cada trabajador: documento, idioma leído, puntaje, fecha y firma. Descarga la constancia de cada uno.','Reading and signature record for each worker: document, language read, score, date and signature. Download each certificate.')}</p>
  <select style="max-width:340px;margin:0" onchange="FF.d=this.value;render()"><option value="all">${x('Todos los documentos','All documents')} (${all.length})</option>${ds.map(d=>`<option value="${d.id}" ${FF.d==d.id?'selected':''}>${esc(docTitle(d))} (${all.filter(r=>r.id==d.id).length})</option>`).join('')}</select></div>
  <div class="card" style="margin-top:14px;padding:4px 16px">${rows.length?rows.map(r=>`<div class="row"><div style="display:flex;gap:12px;align-items:center;min-width:0;flex:1">${r.sg?`<img src="${r.sg}" alt="" class="sgth">`:''}<div style="min-width:0"><b>${esc(whoName(r.u))}</b> <span class="mut" style="font-size:.85em">${esc(r.u)}</span><br><span style="font-size:.9em">${esc(docTitle(dm(r.id)))}</span><br>
  <span class="tag">${(r.lg||'es').toUpperCase()}</span> ${r.sc!=null?`<span class="tag ok">${r.sc} %</span>`:''} <span class="mut" style="font-size:.85em">${r.at?new Date(r.at).toLocaleString('es-CL',{dateStyle:'medium',timeStyle:'short'}):''}</span> ${r.h?`<span class="mut" style="font-size:.8em" title="SHA-256 ${r.h}">🔒 ${r.h.slice(0,10)}…</span>`:''}</div></div>
  <button class="g" onclick="constancia(${r.id},'${esc(r.u)}')">📄 ${x('Constancia','Certificate')}</button></div>`).join(''):`<p class="mut">${x('Aún no hay firmas registradas.','No signatures recorded yet.')}</p>`}</div>`};

/* ===== v8: editor de preguntas por idioma =====
   Español es la versión fuente (agregar, eliminar y reordenar). En los demás idiomas se edita la traducción de cada pregunta. */
const QZL={};
{const eq0=edQz;edQz=function(id,l){const L=['es',...tLangs(id)];l=l||QZL[id]||(L.includes(lang)?lang:'es');if(!L.includes(l))l='es';QZL[id]=l;
  const tabs=`<div class="lang" style="margin:6px 0 10px">${L.map(z=>`<button class="g ${z==l?'on':''}" onclick="edQz(${id},'${z}')">${LN[z]}${z!='es'&&qzState(id,z)!='ok'&&QES(id).length?' ⚠':''}</button>`).join('')}</div>`;
  if(l=='es'){eq0(id);const h=$('#mc h3');if(h)h.insertAdjacentHTML('afterend',tabs);return}
  const es=QES(id),qst=qzState(id,l),tq=QZT[id+'_'+l],qs=qst=='ok'?tq.qs:[];
  M(`<h3>${L5(NS.ev)}</h3>${tabs}<p class="mut">${esc(docTitle(dm(id),l))}</p>
  ${!es.length?`<p>${x('Primero escribe las preguntas en Español.','First write the questions in Spanish.')}</p>`:qst!='ok'?`<div class="card" style="padding:12px 16px"><p style="margin:0 0 10px">${x('Las preguntas en','The questions in')} ${LN[l]} ${x('no están al día con la versión en Español.','are not up to date with the Spanish version.')}</p><button onclick="trAllQ(${id})">🌐 ${x('Traducir a todos los idiomas','Translate into all languages')}</button></div>`:
  qs.map((q,k)=>`<div class="row"><span style="flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${k+1}. ${esc(q.q)}<br><span class="mut" style="font-size:.85em">${esc(es[k]?es[k].q:'')}</span></span><span style="white-space:nowrap"><button class="g" onclick="edQt(${id},'${l}',${k})" title="${x('Editar traducción','Edit translation')}">✎</button></span></div>`).join('')}
  <p class="mut" style="font-size:.85em;margin-top:12px">${x('Aquí editas la traducción de cada pregunta. Para agregar o eliminar preguntas, hazlo en Español: el cambio se aplica a todos los idiomas.','Here you edit the translation of each question. To add or delete questions, do it in Spanish: the change applies to all languages.')}</p>
  <div style="display:flex;gap:8px;flex-wrap:wrap"><button class="g" onclick="edQz(${id},'es')">${LN.es} →</button><button class="g" onclick="X()">${t('cl')}</button></div>`)}}
function edQt(id,l,k){const tq=QZT[id+'_'+l],q=tq&&tq.qs[k],s=QES(id)[k];if(!q||!s)return edQz(id,l);const c=q.c||0,wr=q.o.filter((_,j)=>j!=c),sw=s.o.filter((_,j)=>j!=(s.c||0));
  const f=(lab,idn,v,ref,ta)=>`<label>${lab}</label>${ta?`<textarea id="${idn}" rows="3">${esc(v||'')}</textarea>`:`<input id="${idn}" value="${esc(v||'')}">`}<p class="mut" style="font-size:.8em;margin:2px 0 8px">🇪🇸 ${esc(ref||'—')}</p>`;
  M(`<h3>${L5(NS.qq)} ${k+1} · ${LN[l]}</h3>${f(L5(NS.qq),'t0',q.q,s.q)}${f('✔ '+L5(NS.qc),'t1',q.o[c],s.o[s.c||0])}${f('✖ '+L5(NS.qw)+' 1','t2',wr[0],sw[0])}${f('✖ '+L5(NS.qw)+' 2','t3',wr[1],sw[1])}${f(L5(NS.qe),'t4',q.w,s.w,1)}
  <div style="display:flex;gap:8px;flex-wrap:wrap"><button onclick="saveQt(${id},'${l}',${k})">${x('Guardar','Save')}</button><button class="g" onclick="edQz(${id},'${l}')">${x('Volver','Back')}</button></div>`)}
async function saveQt(id,l,k){const v=n=>$('#t'+n).value.trim();if(!v(0)||!v(1)||!v(2)){toast(x('Completa la pregunta, la respuesta correcta y al menos una incorrecta','Fill in the question, the correct answer and at least one wrong answer'));return}
  const tq=QZT[id+'_'+l],c=(tq.qs[k].c||0),w=[v(2),v(3)].filter(Boolean),o=[];let wi=0;const n=Math.max(tq.qs[k].o.length,1+w.length);for(let j=0;j<n;j++){if(j==c)o.push(v(1));else if(wi<w.length)o.push(w[wi++])}
  const qs=tq.qs.map((q,j)=>j==k?{...q,q:v(0),o,c:Math.min(c,o.length-1),w:v(4)}:q),rec={...tq,id,lang:l,qs,engine:'manual'};QZT[id+'_'+l]=rec;await put('qzt',id+'_'+l,rec);toast('✓ '+x('Traducción guardada','Translation saved'));edQz(id,l)}

/* ===== v10: editor de preguntas con hasta 4 alternativas (1 correcta + 3 incorrectas) ===== */
function edQ(id,k){const q=QES(id)[k]||{q:'',o:['','','',''],c:0,w:''},c=q.c||0,wr=(q.o||[]).filter((_,j)=>j!=c);
  M(`<h3>${L5(NS.qq)} ${k+1}</h3><label>${L5(NS.qq)}</label><input id="q0" value="${esc(q.q)}"><label>✔ ${L5(NS.qc)}</label><input id="q1" value="${esc(q.o[c]||'')}">
  ${[0,1,2].map(j=>`<label>✖ ${L5(NS.qw)} ${j+1}${j?` <span class="mut">(${x('opcional','optional')})</span>`:''}</label><input id="q${j+2}" value="${esc(wr[j]||'')}">`).join('')}
  <label>${L5(NS.qe)}</label><textarea id="q5" rows="3">${esc(q.w||'')}</textarea><p></p><div style="display:flex;gap:8px"><button onclick="saveQ(${id},${k})">${x('Guardar','Save')}</button><button class="g" onclick="edQz(${id},'es')">${x('Volver','Back')}</button></div>`)}
function saveQ(id,k){const v=n=>$('#q'+n).value.trim();if(!v(0)||!v(1)||!v(2)){toast(x('Completa la pregunta, la respuesta correcta y al menos una incorrecta','Fill in the question, the correct answer and at least one wrong answer'));return}
  const qs=[...QES(id)];qs[k]={q:v(0),o:[v(1),v(2),v(3),v(4)].filter(Boolean),c:0,w:v(5)};qzSave(id,qs)}
function edQt(id,l,k){const tq=QZT[id+'_'+l],q=tq&&tq.qs[k],s=QES(id)[k];if(!q||!s)return edQz(id,l);const c=q.c||0,wr=q.o.filter((_,j)=>j!=c),sw=s.o.filter((_,j)=>j!=(s.c||0));
  const f=(lab,idn,v,ref,ta)=>`<label>${lab}</label>${ta?`<textarea id="${idn}" rows="3">${esc(v||'')}</textarea>`:`<input id="${idn}" value="${esc(v||'')}">`}<p class="mut" style="font-size:.8em;margin:2px 0 8px">🇪🇸 ${esc(ref||'—')}</p>`;
  M(`<h3>${L5(NS.qq)} ${k+1} · ${LN[l]}</h3>${f(L5(NS.qq),'t0',q.q,s.q)}${f('✔ '+L5(NS.qc),'t1',q.o[c],s.o[s.c||0])}${sw.map((z,j)=>f('✖ '+L5(NS.qw)+' '+(j+1),'t'+(j+2),wr[j],z)).join('')}${f(L5(NS.qe),'t9',q.w,s.w,1)}
  <div style="display:flex;gap:8px;flex-wrap:wrap"><button onclick="saveQt(${id},'${l}',${k})">${x('Guardar','Save')}</button><button class="g" onclick="edQz(${id},'${l}')">${x('Volver','Back')}</button></div>`)}
async function saveQt(id,l,k){const v=n=>($('#t'+n)||{value:''}).value.trim();if(!v(0)||!v(1)||!v(2)){toast(x('Completa la pregunta, la respuesta correcta y al menos una incorrecta','Fill in the question, the correct answer and at least one wrong answer'));return}
  const tq=QZT[id+'_'+l],w=[v(2),v(3),v(4)].filter(Boolean),c=Math.min(tq.qs[k].c||0,w.length),o=[...w];o.splice(c,0,v(1));
  const qs=tq.qs.map((q,j)=>j==k?{...q,q:v(0),o,c,w:v(9)}:q),rec={...tq,id,lang:l,qs,engine:'manual'};QZT[id+'_'+l]=rec;await put('qzt',id+'_'+l,rec);toast('✓ '+x('Traducción guardada','Translation saved'));edQz(id,l)}


/* ===== subida directa a Supabase (no pasa por el servidor: sin límite de 4,5 MB en Vercel) ===== */
async function upDirect(id,l,file,kind,status,engine){
  const meta={doc:id,lang:l,kind:kind||'translation',name:file.name,size:file.size,status:status||'review',engine:engine||''};
  const post=async(u,b)=>{const r=await fetch(u,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(b)});const j=await r.json().catch(()=>({}));if(!r.ok)throw new Error(j.error||'HTTP '+r.status);return j};
  const sg=await post('/api/files/sign-upload',meta);
  const up=await fetch(sg.url,{method:'PUT',headers:{'Content-Type':sg.mime,'x-upsert':'true'},body:file});
  if(!up.ok){const t=await up.text().catch(()=>'');throw new Error(x('No se pudo subir el archivo a Supabase','Could not upload the file to Supabase')+' ('+up.status+') '+t.slice(0,120))}
  return post('/api/files/commit',meta)}
