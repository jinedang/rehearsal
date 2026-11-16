"use strict";
/* 발표 리허설 룸 — 정적 웹사이트 + Supabase */

/* ───────── 공통 ───────── */
const CONFIG = window.RR_CONFIG || {};
const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const pad = n => String(n).padStart(2,'0');
const fmt = sec => { const neg = sec < 0; sec = Math.abs(Math.floor(sec)); return (neg?'-':'') + pad(Math.floor(sec/60)) + ':' + pad(sec%60); };
const fmtKo = sec => { sec = Math.round(Math.abs(sec)); const m = Math.floor(sec/60), s = sec%60; return m ? (s ? `${m}분 ${s}초` : `${m}분`) : `${s}초`; };
const LS = {
  get(k,d){ try{ const v = localStorage.getItem('rr:'+k); return v == null ? d : JSON.parse(v); }catch(e){ return d; } },
  set(k,v){ try{ localStorage.setItem('rr:'+k, JSON.stringify(v)); }catch(e){} }
};
const nonSpace = t => (t||'').replace(/\s/g,'').length;
const CPM = 320; // 발표 속도: 공백 제외 분당 320자
const estSec = t => Math.round(nonSpace(t) / CPM * 60);
const tms = v => v ? (typeof v === 'number' ? v : Date.parse(v)) : 0;
const dt = v => { const ms = tms(v); if(!ms) return ''; const d = new Date(ms); return `${d.getMonth()+1}/${d.getDate()} ${pad(d.getHours())}:${pad(d.getMinutes())}`; };
const DEFAULT_CATS = ['사업 이해','기술·아키텍처','수행 방안','일정·인력','데이터·보안','기대효과·정량 목표','경쟁사 대비 차별화','기타'];
const me = () => (S.me || '').trim() || '이름 없음';

function toast(msg, ms=3200){ const d = document.createElement('div'); d.textContent = msg; $('#toast').appendChild(d); setTimeout(()=>d.remove(), ms); }
function showNotice(msg){ const n = $('#notice'); n.textContent = msg; n.classList.remove('hidden'); }
function hideNotice(){ $('#notice').classList.add('hidden'); }
function ask(title, msg, yes='확인'){
  return new Promise(res => {
    $('#ovTitle').textContent = title; $('#ovMsg').textContent = msg; $('#ovYes').textContent = yes;
    $('#overlay').classList.remove('hidden'); $('#ovYes').focus();
    const done = v => { $('#overlay').classList.add('hidden'); $('#ovYes').onclick = $('#ovNo').onclick = null; res(v); };
    $('#ovYes').onclick = () => done(true); $('#ovNo').onclick = () => done(false);
  });
}

/* ───────── 상태 ───────── */
const S = {
  me: LS.get('me',''),
  projects:[], projectsLoaded:false, pid:null, project:null,
  slides:[], scripts:{}, qa:[], marks:{}, history:[],
  thumbs:{}, tab:'rehearsal', channel:null, ready:false
};
let sb = null;

async function q(promise){ const {data, error} = await promise; if(error) throw error; return data; }
async function W(fn, okMsg){
  try { const r = await fn(); if(okMsg) toast(okMsg); return r; }
  catch(e){
    console.error(e);
    const m = (e && (e.message || e.error_description)) || String(e);
    if(/Failed to fetch|NetworkError/i.test(m)) toast('서버에 연결하지 못했습니다. 인터넷 연결을 확인하세요.', 5000);
    else if(/row-level security|permission|not allowed/i.test(m)) toast('저장 권한이 없습니다. supabase/schema.sql을 실행했는지 확인하세요.', 6000);
    else if(/exceeded|too large|Payload/i.test(m)) toast('파일이 너무 큽니다. 크기를 줄여 다시 올려 주세요.', 5000);
    else toast('저장하지 못했습니다: ' + m, 5000);
    throw e;
  }
}

/* ───────── 외부 라이브러리 (필요할 때만 로드) ───────── */
const LIBS = {
  zip: ['https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js'],
  pdf: ['https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js','https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js'],
  pptx: ['https://cdn.jsdelivr.net/npm/pptx-preview@1.0.7/dist/pptx-preview.umd.js']
};
const libLoads = {};
function loadScript(src){ return new Promise((res,rej)=>{ const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('라이브러리를 불러오지 못했습니다: ' + src)); document.head.appendChild(s); }); }
function ensureLibs(name){
  if(!libLoads[name]) libLoads[name] = (async () => {
    for(const u of LIBS[name]) await loadScript(u);
    if(name === 'pdf' && window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = LIBS.pdf[1];
  })().catch(e => { delete libLoads[name]; throw e; });
  return libLoads[name];
}

/* ───────── 파일 파싱 ───────── */
const LN = (node, name) => Array.from(node.getElementsByTagNameNS('*', name));
function resolvePath(base, target){
  if(target.startsWith('/')) return target.slice(1);
  const st = [];
  for(const p of (base + target).split('/')){ if(p === '..') st.pop(); else if(p && p !== '.') st.push(p); }
  return st.join('/');
}
function phType(el){
  let n = el;
  while(n && n.localName !== 'sp') n = n.parentNode;
  if(!n || !n.getElementsByTagNameNS) return null;
  const ph = LN(n, 'ph')[0];
  return ph ? (ph.getAttribute('type') || 'body') : null;
}
function paraLines(root, skipTypes){
  return LN(root, 'p').filter(p => p.namespaceURI && p.namespaceURI.includes('drawingml'))
    .filter(p => !skipTypes || !skipTypes.includes(phType(p)))
    .map(p => LN(p, 't').map(t => t.textContent).join('').replace(/\s+/g,' ').trim())
    .filter(Boolean);
}
async function parsePptx(file){
  await ensureLibs('zip');
  const zip = await JSZip.loadAsync(file);
  const xml = async p => { const f = zip.file(p); if(!f) return null; return new DOMParser().parseFromString(await f.async('string'), 'application/xml'); };
  const pres = await xml('ppt/presentation.xml');
  const rels = await xml('ppt/_rels/presentation.xml.rels');
  if(!pres || !rels) throw new Error('PPTX 구조를 읽지 못했습니다');
  const relMap = {}; LN(rels,'Relationship').forEach(r => relMap[r.getAttribute('Id')] = r.getAttribute('Target'));
  const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const ids = LN(pres,'sldId').map(s => s.getAttributeNS(R_NS,'id') || s.getAttribute('r:id'));
  const out = [];
  for(let k = 0; k < ids.length; k++){
    const target = relMap[ids[k]]; if(!target) continue;
    const path = resolvePath('ppt/', target);
    const doc = await xml(path); if(!doc) continue;
    if(doc.documentElement.getAttribute('show') === '0') continue; // 숨긴 슬라이드 제외
    let title = '';
    for(const sp of LN(doc,'sp')){
      const ph = LN(sp,'ph')[0]; const t = ph && ph.getAttribute('type');
      if(t === 'title' || t === 'ctrTitle'){ title = paraLines(sp).join(' '); if(title) break; }
    }
    const lines = paraLines(doc, ['sldNum','dt','ftr']);
    if(!title) title = lines[0] || '';
    const body = lines.filter((l,i) => !(i === 0 && l === title)).join('\n');
    let notes = '';
    const sr = await xml(path.replace(/([^/]+)$/, '_rels/$1.rels'));
    if(sr){
      const nr = LN(sr,'Relationship').find(r => /notesSlide$/.test(r.getAttribute('Type')||''));
      if(nr){
        const nd = await xml(resolvePath(path.replace(/[^/]+$/,''), nr.getAttribute('Target')));
        if(nd){
          const bodies = LN(nd,'sp').filter(sp => { const ph = LN(sp,'ph')[0]; return ph && ph.getAttribute('type') === 'body'; });
          notes = (bodies.length ? bodies.flatMap(sp => paraLines(sp)) : paraLines(nd, ['sldNum','sldImg','hdr','ftr','dt'])).join('\n');
        }
      }
    }
    out.push({ title: title.slice(0,200), body: body.slice(0,5000), notes: notes.slice(0,12000), pptx_index: k });
  }
  return out;
}
async function parseDocx(file){
  await ensureLibs('zip');
  const zip = await JSZip.loadAsync(file);
  const f = zip.file('word/document.xml'); if(!f) throw new Error('DOCX 본문을 찾지 못했습니다');
  const doc = new DOMParser().parseFromString(await f.async('string'), 'application/xml');
  return LN(doc,'p').map(p => LN(p,'t').map(t => t.textContent).join('').trim()).filter(Boolean).join('\n');
}
async function pdfPageTexts(pdf, onProgress){
  const out = [];
  for(let i = 1; i <= pdf.numPages; i++){
    const pg = await pdf.getPage(i); const tc = await pg.getTextContent();
    const lines = []; let line = '', lastY = null;
    for(const it of tc.items){
      if(!('str' in it)) continue;
      const y = it.transform[5];
      if(lastY !== null && Math.abs(y - lastY) > 2){ lines.push(line); line = ''; }
      line += it.str; lastY = y;
      if(it.hasEOL){ lines.push(line); line = ''; lastY = null; }
    }
    if(line) lines.push(line);
    out.push(lines.map(s => s.replace(/\s+/g,' ').trim()).filter(Boolean).join('\n'));
    if(onProgress) onProgress(i, pdf.numPages);
  }
  return out;
}

/* ───────── 엑셀 복사(TSV)·JSON 파서 ───────── */
function parseTSV(text){
  text = text.replace(/\r\n?/g,'\n');
  const rows = []; let row = [], cell = '', inQ = false;
  for(let i = 0; i < text.length; i++){
    const c = text[i];
    if(inQ){ if(c === '"'){ if(text[i+1] === '"'){ cell += '"'; i++; } else inQ = false; } else cell += c; continue; }
    if(c === '"' && cell === ''){ inQ = true; continue; }
    if(c === '\t'){ row.push(cell); cell = ''; continue; }
    if(c === '\n'){ row.push(cell); rows.push(row); row = []; cell = ''; continue; }
    cell += c;
  }
  if(cell !== '' || row.length){ row.push(cell); rows.push(row); }
  return rows.map(r => r.map(c => c.trim())).filter(r => r.some(Boolean));
}
function rowsToQA(rows){
  if(!rows.length) return [];
  const h = rows[0].map(c => c.replace(/\s/g,''));
  const isHeader = h.some(c => /질문|질의/.test(c) && !/답/.test(c));
  let map;
  if(isHeader){
    map = {};
    h.forEach((c,i) => {
      if(map.cat == null && /구분|분류|유형|카테고리/.test(c)) map.cat = i;
      else if(map.q == null && /질문|질의/.test(c) && !/답/.test(c)) map.q = i;
      else if(map.a == null && /답/.test(c)) map.a = i;
      else if(map.page == null && /페이지|쪽|page|장표/i.test(c)) map.page = i;
      else if(map.rfp == null && /RFP|요구사항|ID/i.test(c)) map.rfp = i;
      else if(map.kw == null && /키워드|태그/.test(c)) map.kw = i;
    });
    rows = rows.slice(1);
  } else {
    const n = Math.max(...rows.map(r => r.length));
    map = n >= 4 ? {cat:0,q:1,page:2,a:3,rfp:4,kw:5} : n === 3 ? {cat:0,q:1,a:2} : n === 2 ? {q:0,a:1} : {q:0};
  }
  const g = (r,k) => map[k] != null ? (r[map[k]] || '') : '';
  return rows.map(r => ({cat:g(r,'cat'), q:g(r,'q'), a:g(r,'a'), page:g(r,'page'), rfp:g(r,'rfp'), kw:g(r,'kw')})).filter(x => x.q);
}
function parseJSONArray(t){
  t = String(t || '').replace(/```(json)?/g, '');
  const a = t.indexOf('['); if(a < 0) return [];
  const body = t.slice(a); const b = body.lastIndexOf(']');
  const tryParse = s => { try{ const v = JSON.parse(s); return Array.isArray(v) ? v : null; }catch(e){ return null; } };
  let v = b > 0 ? tryParse(body.slice(0, b + 1)) : null;
  if(!v){ const c = body.lastIndexOf('}'); if(c > 0) v = tryParse(body.slice(0, c + 1) + ']'); }
  return (v || []).filter(x => x && typeof x.q === 'string' && x.q.trim()).map(x => ({cat:String(x.cat||''), q:x.q.trim(), a:String(x.a||'').trim(), page:String(x.page||''), rfp:String(x.rfp||''), kw: Array.isArray(x.kw) ? x.kw.join(', ') : String(x.kw||'')}));
}
function parseQAInput(text){ const t = text.trim(); if(t.startsWith('[') || t.startsWith('```')) { const j = parseJSONArray(t); if(j.length) return j; } return rowsToQA(parseTSV(text)); }

/* ───────── 발표자료 표시 (PPTX 또는 PDF) ───────── */
const Deck = {
  key:null, kind:null, pdf:null, holder:null, ready:null, error:null,
  reset(){ if(this.holder) this.holder.remove(); this.holder = null; this.pdf = null; this.kind = null; this.error = null; this.ready = null; this.key = null; },
  async load(){
    const d = S.project && S.project.deck; const key = d ? d.url : null;
    if(key === this.key) return this.ready;
    this.reset(); this.key = key;
    if(!d){ return null; }
    this.kind = d.kind;
    this.ready = (async () => {
      if(d.kind === 'pdf'){
        await ensureLibs('pdf');
        const doc = await pdfjsLib.getDocument({url: d.url}).promise;
        if(this.key === key) this.pdf = doc;
      } else {
        await ensureLibs('pptx');
        const res = await fetch(d.url); if(!res.ok) throw new Error('발표자료 파일을 받지 못했습니다 (' + res.status + ')');
        const buf = await res.arrayBuffer();
        const holder = document.createElement('div'); holder.className = 'pptx-holder'; holder.setAttribute('aria-hidden','true');
        document.body.appendChild(holder);
        const pv = pptxPreview.init(holder, {width: 960, height: 540, mode: 'list'});
        await pv.preview(buf);
        if(this.key === key) this.holder = holder; else holder.remove();
      }
      return true;
    })().catch(e => { console.warn('발표자료 표시 실패', e); this.error = e; return null; });
    await this.ready;
    S.thumbs = {}; R.lastShown = null;
    if(S.tab === 'rehearsal') renderRehearsal();
    if(S.tab === 'scripts') renderScripts();
    return this.ready;
  },
  pages(){ if(this.pdf) return this.pdf.numPages; if(this.holder) return this.holder.querySelectorAll('.pptx-preview-slide-wrapper').length; return 0; },
  srcEl(no){
    if(!this.holder) return null;
    const s = S.slides[no-1]; const idx = s && s.pptx_index != null ? s.pptx_index : no - 1;
    return this.holder.querySelector('.pptx-preview-slide-wrapper-' + idx);
  },
  has(no){ return this.pdf ? no <= this.pdf.numPages : !!this.srcEl(no); },
  domClone(no, maxW, maxH){
    const src = this.srcEl(no); if(!src) return null;
    const w = src.offsetWidth || 960, h = src.offsetHeight || 540;
    const sc = Math.min(maxW / w, maxH / h);
    const outer = document.createElement('div'); outer.className = 'deck-slide';
    outer.style.width = Math.floor(w * sc) + 'px'; outer.style.height = Math.floor(h * sc) + 'px';
    const c = src.cloneNode(true);
    const a = src.querySelectorAll('canvas'), b = c.querySelectorAll('canvas');
    b.forEach((cv, i) => { try{ cv.getContext('2d').drawImage(a[i], 0, 0); }catch(e){} });
    c.style.width = w + 'px'; c.style.height = h + 'px'; c.style.transform = `scale(${sc})`;
    outer.appendChild(c);
    return outer;
  }
};
async function drawPdfPage(pdf, no, canvas, w, h){
  if(canvas._task){ try{ canvas._task.cancel(); }catch(e){} canvas._task = null; }
  const page = await pdf.getPage(no);
  const v1 = page.getViewport({scale:1});
  const scale = Math.min(w / v1.width, h / v1.height);
  const dpr = Math.min(window.devicePixelRatio || 1, 2);
  const vp = page.getViewport({scale: scale * dpr});
  canvas.width = Math.floor(vp.width); canvas.height = Math.floor(vp.height);
  canvas.style.width = Math.floor(vp.width / dpr) + 'px'; canvas.style.height = Math.floor(vp.height / dpr) + 'px';
  const task = page.render({canvasContext: canvas.getContext('2d'), viewport: vp});
  canvas._task = task;
  try{ await task.promise; }catch(e){}
}
const thumbQueue = []; let thumbBusy = false;
function queueThumb(no){ if(!Deck.pdf || S.thumbs[no] || thumbQueue.includes(no) || no > Deck.pdf.numPages) return; thumbQueue.push(no); runThumbs(); }
async function runThumbs(){
  if(thumbBusy) return; thumbBusy = true;
  while(thumbQueue.length){
    const no = thumbQueue.shift(); const pdf = Deck.pdf; if(!pdf) break;
    try{
      const c = document.createElement('canvas');
      const page = await pdf.getPage(no); const v1 = page.getViewport({scale:1});
      const vp = page.getViewport({scale: 400 / v1.width});
      c.width = vp.width; c.height = vp.height;
      await page.render({canvasContext:c.getContext('2d'), viewport:vp}).promise;
      if(pdf === Deck.pdf){ S.thumbs[no] = c.toDataURL('image/jpeg', .82); const el = document.querySelector(`[data-thumb="${no}"]`); if(el) el.innerHTML = `<img alt="슬라이드 ${no}" src="${S.thumbs[no]}">`; }
    }catch(e){}
  }
  thumbBusy = false;
}
function fillThumb(el){
  const no = +el.dataset.thumb;
  if(Deck.pdf){ if(S.thumbs[no]) el.innerHTML = `<img alt="슬라이드 ${no}" src="${S.thumbs[no]}">`; else queueThumb(no); }
  else if(Deck.holder){ const cs = getComputedStyle(el); const w = el.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight), h = el.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom); const c = Deck.domClone(no, w || 188, h || 100); if(c){ el.replaceChildren(c); } }
}

/* ───────── 데이터 (Supabase) ───────── */
async function init(){
  bindUI(); $('#meName').value = S.me;
  if(!CONFIG.SUPABASE_URL || !CONFIG.SUPABASE_ANON_KEY || !window.supabase){
    showSetupGuide(); return;
  }
  sb = window.supabase.createClient(CONFIG.SUPABASE_URL, CONFIG.SUPABASE_ANON_KEY, {auth: {persistSession: false}});
  route();
  try{ await loadProjects(); }
  catch(e){ console.error(e); S.projectsLoaded = true; showNotice('데이터를 불러오지 못했습니다. config.js의 Supabase 주소·키와 supabase/schema.sql 실행 여부를 확인하세요. (' + (e.message || e) + ')'); renderAll(); return; }
  sb.channel('rr-projects').on('postgres_changes', {event:'*', schema:'public', table:'projects'}, () => debounced('projects', loadProjects)).subscribe();
}
function showSetupGuide(){
  document.querySelector('main').innerHTML = `<div class="card setup-guide">
    <h2>처음 설정이 필요합니다</h2>
    <p>이 사이트는 팀 공용 데이터를 Supabase에 저장합니다. 아래 순서대로 한 번만 설정하면 됩니다.</p>
    <ol class="steps">
      <li>supabase.com에서 무료 프로젝트를 만듭니다.</li>
      <li>프로젝트의 SQL Editor에 저장소의 <code>supabase/schema.sql</code> 내용을 붙여넣고 실행합니다.</li>
      <li>프로젝트 주소(URL)와 publishable(또는 anon public) 키를 <code>config.js</code>에 넣고 저장소에 다시 올립니다.</li>
    </ol>
    <p class="progress-note">${window.supabase ? '' : 'Supabase 라이브러리를 불러오지 못했습니다. 인터넷 연결이나 회사 방화벽을 확인하세요.'}</p>
  </div>`;
  $('#projMeta').textContent = '설정 필요';
  $('#projSel').innerHTML = '<option>설정 필요</option>';
}
const timers = {};
function debounced(key, fn, ms=300){ clearTimeout(timers[key]); timers[key] = setTimeout(() => { fn().catch(e => console.warn(e)); }, ms); }
async function loadProjects(){
  S.projects = await q(sb.from('projects').select('*').order('created_at'));
  S.projectsLoaded = true; hideNotice();
  let want = (S.pid && S.projects.some(p => p.id === S.pid)) ? S.pid : LS.get('pid', null);
  if(!S.projects.some(p => p.id === want)) want = S.projects.length ? S.projects[S.projects.length-1].id : null;
  if(want !== S.pid) await selectProject(want);
  else { S.project = S.projects.find(p => p.id === S.pid) || null; Deck.load(); renderAll(); }
}
async function loadTable(t){
  const pid = S.pid; if(!pid) return;
  if(t === 'slides'){ const r = await q(sb.from('slides').select('*').eq('project_id', pid).order('no')); if(pid === S.pid){ S.slides = r; R.clampIdx(); } }
  else if(t === 'scripts'){ const r = await q(sb.from('scripts').select('*').eq('project_id', pid)); if(pid === S.pid){ const m = {}; r.forEach(x => m[x.no] = x); S.scripts = m; } }
  else if(t === 'qa'){ const r = await q(sb.from('qa').select('*').eq('project_id', pid).order('created_at')); if(pid === S.pid) S.qa = r; }
  onData(t);
}
async function selectProject(pid){
  if(S.channel){ try{ sb.removeChannel(S.channel); }catch(e){} S.channel = null; }
  S.pid = pid; if(pid) LS.set('pid', pid);
  S.project = S.projects.find(p => p.id === pid) || null;
  S.slides = []; S.scripts = {}; S.qa = []; S.thumbs = {};
  S.marks = pid ? LS.get('marks:' + pid, {}) : {};
  S.history = pid ? LS.get('hist:' + pid, []) : [];
  setupMsg.deck = setupMsg.prop = '';
  R.reset(); stopDrill(true); qaPanel = null; aiItems = null; scEditing = null;
  applyRhSettings(); renderAll();
  if(!pid){ Deck.reset(); return; }
  Deck.load();
  await Promise.all(['slides','scripts','qa'].map(t => loadTable(t).catch(e => console.warn(e))));
  const ch = sb.channel('rr-' + pid);
  ['slides','scripts','qa'].forEach(t => ch.on('postgres_changes', {event:'*', schema:'public', table:t, filter:'project_id=eq.' + pid}, () => debounced(t, () => loadTable(t))));
  ch.subscribe(); S.channel = ch;
}
function onData(kind){
  if(S.tab === 'rehearsal') renderRehearsal(); else R.lastShown = null;
  if(S.tab === 'scripts') renderScripts();
  if(S.tab === 'qa') renderQA();
  if(S.tab === 'practice') renderPractice();
  if(S.tab === 'setup') renderSetup();
  renderProjSel();
}
async function uploadFile(pid, file, kind){
  const ext = (file.name.split('.').pop() || 'bin').toLowerCase().replace(/[^a-z0-9]/g,'');
  const path = `${pid}/${kind}-${Date.now()}.${ext}`;
  const types = {pptx:'application/vnd.openxmlformats-officedocument.presentationml.presentation', pdf:'application/pdf', docx:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'};
  const {error} = await sb.storage.from('files').upload(path, file, {contentType: file.type || types[ext] || 'application/octet-stream', upsert: true});
  if(error) throw error;
  return {path, url: sb.storage.from('files').getPublicUrl(path).data.publicUrl};
}
async function removeFiles(paths){ paths = paths.filter(Boolean); if(paths.length){ try{ await sb.storage.from('files').remove(paths); }catch(e){} } }

/* ───────── 라우팅 ───────── */
const TABS = ['rehearsal','scripts','qa','practice','setup'];
function route(){
  let t = (location.hash || '').replace('#','');
  if(!TABS.includes(t)) t = 'rehearsal';
  S.tab = t;
  document.querySelectorAll('.tabs a').forEach(a => { if(a.dataset.tab === t) a.setAttribute('aria-current','page'); else a.removeAttribute('aria-current'); });
  TABS.forEach(x => { const v = $('#v-'+x); if(v) v.classList.toggle('on', x === t); });
  if(t !== 'rehearsal' && R.running) R.pause();
  if(t === 'rehearsal') R.lastShown = null;
  renderAll();
}
window.addEventListener('hashchange', () => { if(sb) route(); });

function renderProjSel(){
  const sel = $('#projSel');
  if(!S.projectsLoaded) return;
  sel.innerHTML = S.projects.length ? S.projects.map(p => `<option value="${esc(p.id)}"${p.id === S.pid ? ' selected':''}>${esc(p.name || '이름 없는 사업')}</option>`).join('') : '<option value="">등록된 사업 없음</option>';
  $('#projMeta').textContent = S.project ? `슬라이드 ${S.slides.length}장 · 스크립트 ${Object.values(S.scripts).filter(x => (x.body||'').trim()).length}장 · 질의 ${S.qa.length}개` : '사업을 선택하세요';
}
function renderAll(){
  if(!sb) return;
  renderProjSel();
  if(S.tab === 'rehearsal'){ renderRehearsal(); renderHistory(); }
  if(S.tab === 'scripts') renderScripts();
  if(S.tab === 'qa') renderQA();
  if(S.tab === 'practice') renderPractice();
  if(S.tab === 'setup') renderSetup();
}
function noProjectHTML(){
  if(!S.projectsLoaded) return '<div class="empty"><b>불러오는 중입니다</b>잠시만 기다려 주세요.</div>';
  return '<div class="empty"><b>아직 등록된 사업이 없습니다</b><a href="#setup">사업 자료 관리</a> 탭에서 사업을 만들고 발표자료와 정성제안서를 올려 주세요.</div>';
}

/* ───────── 1. 셀프 리허설 ───────── */
const R = {
  running:false, base:0, segStart:0, idx:0, per:[], lastShown:null, raf:0, started:false,
  slideCount(){ return Math.max(1, S.slides.length || Deck.pages()); },
  clampIdx(){ const n = this.slideCount(); if(this.idx >= n) this.idx = n - 1; },
  live(){ return this.running ? performance.now() - this.segStart : 0; },
  total(){ return this.base + this.live(); },
  slideElapsed(i = this.idx){ return (this.per[i] || 0) + (i === this.idx ? this.live() : 0); },
  commit(){ if(!this.running) return; const now = performance.now(); const d = now - this.segStart; this.base += d; this.per[this.idx] = (this.per[this.idx] || 0) + d; this.segStart = now; },
  start(){ if(this.running) return; this.running = true; this.started = true; this.segStart = performance.now(); loop(); updateToggle(); },
  pause(){ if(!this.running) return; this.commit(); this.running = false; cancelAnimationFrame(this.raf); tick(); updateToggle(); },
  toggle(){ this.running ? this.pause() : this.start(); },
  go(i){ const n = this.slideCount(); i = Math.max(0, Math.min(n - 1, i)); if(i === this.idx) return; this.commit(); this.idx = i; renderRehearsal(); tick(); },
  reset(){ cancelAnimationFrame(this.raf); this.running = false; this.base = 0; this.per = []; this.idx = 0; this.started = false; this.lastShown = null; }
};
function totalTargetSec(){ return Math.max(0, (parseInt($('#rhMin').value,10) || 0) * 60 + Math.min(59, parseInt($('#rhSec').value,10) || 0)); }
function targets(){
  const n = R.slideCount(), total = totalTargetSec(), mode = $('#rhAlloc').value;
  if(mode === 'none' || !total) return null;
  let w = Array.from({length:n}, () => 1);
  if(mode === 'script'){
    const c = Array.from({length:n}, (_, i) => nonSpace((S.scripts[i+1] || {}).body));
    const sum = c.reduce((a,b) => a+b, 0);
    if(sum > 0){ const floor = sum / n * 0.25; w = c.map(x => Math.max(x, floor)); }
  }
  const ws = w.reduce((a,b) => a+b, 0);
  const t = w.map(x => Math.floor(total * x / ws));
  let rest = total - t.reduce((a,b) => a+b, 0);
  for(let i = 0; rest > 0; i = (i+1) % n, rest--) t[i]++;
  return t;
}
function applyRhSettings(){
  const p = S.project; const saved = S.pid ? LS.get('rh:' + S.pid, null) : null;
  $('#rhMin').value = saved ? saved.min : (p && p.default_min != null ? p.default_min : 20);
  $('#rhSec').value = saved ? saved.sec : (p && p.default_sec != null ? p.default_sec : 0);
  if(saved && saved.alloc) $('#rhAlloc').value = saved.alloc;
  if(saved && saved.warn) $('#rhWarn').value = saved.warn;
  R.lastShown = null; tick();
}
function saveRhSettings(){ if(S.pid) LS.set('rh:' + S.pid, {min: $('#rhMin').value, sec: $('#rhSec').value, alloc: $('#rhAlloc').value, warn: $('#rhWarn').value}); }
function updateToggle(){ const b = $('#rhToggle'); b.textContent = R.running ? '일시정지' : (R.started ? '이어서 시작' : '시작'); b.classList.toggle('running', R.running); }
function loop(){ tick(); if(R.running) R.raf = requestAnimationFrame(() => setTimeout(loop, 100)); }
function tick(){
  const total = totalTargetSec(), elapsed = R.total() / 1000, remain = total - elapsed;
  const warn = parseInt($('#rhWarn').value, 10);
  const state = !R.started ? 'idle' : remain < 0 ? 'over' : remain <= warn ? 'warn' : 'run';
  $('#clock').dataset.state = state;
  $('#ckLabel').textContent = remain < 0 ? '시간 초과' : state === 'warn' ? '마무리할 시간' : '남은 시간';
  $('#ckMain').textContent = remain < 0 ? '+' + fmt(-remain) : fmt(Math.ceil(remain));
  $('#ckTotal').textContent = '목표 ' + fmt(total);
  $('#ckBar').style.width = (total ? Math.min(100, elapsed / total * 100) : 0) + '%';
  $('#ckElapsed').textContent = fmt(elapsed);
  const tg = targets(); const se = R.slideElapsed() / 1000;
  $('#ckSlide').textContent = fmt(se);
  const sb2 = $('#ckSlideBar'), pace = $('#ckPace');
  if(tg){
    const t = tg[R.idx] || 0;
    $('#ckSlideT').textContent = '/ ' + fmt(t);
    sb2.style.width = (t ? Math.min(100, se / t * 100) : 0) + '%'; sb2.classList.toggle('over', se > t);
    const plannedStart = tg.slice(0, R.idx).reduce((a,b) => a+b, 0);
    const delta = (elapsed - se - plannedStart) + Math.max(0, se - t);
    if(!R.started){ pace.className = 'pace'; pace.textContent = `슬라이드당 평균 ${fmtKo(total / R.slideCount())} 목표`; }
    else if(delta > 3){ pace.className = 'pace late'; pace.textContent = `계획보다 ${fmtKo(delta)} 늦음`; }
    else if(delta < -3){ pace.className = 'pace early'; pace.textContent = `계획보다 ${fmtKo(-delta)} 여유`; }
    else { pace.className = 'pace'; pace.textContent = '계획대로 진행 중'; }
  } else {
    $('#ckSlideT').textContent = ''; sb2.style.width = '0%';
    pace.className = 'pace'; pace.textContent = R.started ? '슬라이드별 목표 없이 전체 시간만 재는 중' : '전체 시간만 잽니다';
  }
}
function renderRehearsal(){
  const n = R.slideCount(); R.clampIdx();
  $('#rhPos').textContent = `${R.idx + 1} / ${n}`;
  $('#rhPrev').disabled = R.idx === 0; $('#rhNext').disabled = R.idx >= n - 1;
  const no = R.idx + 1; const slide = S.slides[R.idx];
  const mode = Deck.pdf ? 'pdf' : Deck.holder ? 'pptx' : 'txt';
  const key = [mode, S.pid, no, slide ? slide.title + (slide.body||'').length : ''].join(':');
  if(R.lastShown !== key){
    R.lastShown = key;
    const cv = $('#slideCanvas'), dom = $('#slideDom'), tx = $('#slideText');
    const box = $('#slideBox').getBoundingClientRect();
    const bw = Math.max(200, box.width), bh = Math.max(120, box.height);
    cv.classList.add('hidden'); dom.classList.add('hidden'); tx.classList.add('hidden');
    if(mode === 'pdf' && Deck.has(no)){ cv.classList.remove('hidden'); drawPdfPage(Deck.pdf, no, cv, bw, bh); }
    else if(mode === 'pptx' && Deck.has(no)){ dom.classList.remove('hidden'); dom.replaceChildren(Deck.domClone(no, bw, bh)); }
    else {
      tx.classList.remove('hidden');
      const loading = S.project && S.project.deck && !Deck.error && (!Deck.pdf && !Deck.holder);
      if(!S.project) tx.innerHTML = `<span class="st-no">타이머 전용</span><h3>사업을 선택하면 발표자료가 여기에 나옵니다</h3><pre>사업이 없어도 위에서 시간을 정하고 시작하면 타이머로 쓸 수 있습니다.</pre>`;
      else if(!slide) tx.innerHTML = `<span class="st-no">${loading ? '발표자료를 불러오는 중' : '발표자료 없음'}</span><h3>${loading ? '잠시만 기다려 주세요' : '아직 발표자료가 올라오지 않았습니다'}</h3><pre>타이머는 지금도 쓸 수 있습니다. 발표자료는 사업 자료 관리 탭에서 올립니다.</pre>`;
      else tx.innerHTML = `<span class="st-no">슬라이드 ${no}</span><h3>${esc(slide.title || '제목 없음')}</h3><pre>${esc(slide.body || '')}</pre><span class="st-note">${loading ? '슬라이드 화면을 불러오는 중입니다' : Deck.error ? '슬라이드 화면을 그리지 못해 글자만 보여 줍니다' : ''}</span>`;
    }
  }
  const sc = S.scripts[no]; const txt = (sc && sc.body || '').trim();
  $('#spTitle').textContent = `슬라이드 ${no} 스크립트` + (txt ? ` · ${nonSpace(txt)}자 · 약 ${fmtKo(estSec(txt))}` : '');
  const sp = $('#spText'); sp.textContent = txt || '이 슬라이드에 작성된 스크립트가 없습니다. 스크립트 탭에서 작성할 수 있습니다.'; sp.classList.toggle('none', !txt);
  tick(); updateToggle();
}
function finishRehearsal(){
  if(!R.started){ toast('아직 리허설을 시작하지 않았습니다.'); return; }
  R.pause();
  const n = R.slideCount(), tg = targets(), total = totalTargetSec();
  const per = Array.from({length:n}, (_, i) => Math.round((R.per[i] || 0) / 1000));
  const actual = Math.round(R.base / 1000);
  const sign = d => d > 0 ? '+' + fmt(d) : d < 0 ? '-' + fmt(-d) : '0';
  const rows = per.map((s, i) => {
    const t = tg ? tg[i] : null; const d = t != null ? s - t : null; const title = (S.slides[i] && S.slides[i].title) || '';
    return `<tr><td class="num">${i+1}</td><td>${esc(title)}</td><td class="num">${t != null ? fmt(t) : '-'}</td><td class="num">${fmt(s)}</td><td class="num ${d > 5 ? 'd-late' : d < -5 ? 'd-early' : ''}">${d != null ? sign(d) : '-'}</td></tr>`;
  }).join('');
  const diff = actual - total;
  $('#rhResult').innerHTML = `<div class="card result" style="display:flex;flex-direction:column;gap:12px">
    <div class="row" style="justify-content:space-between"><h3>리허설 결과</h3><span class="pill ${Math.abs(diff) <= 15 ? 'ok' : diff > 0 ? 'bad' : 'warn'}">${diff > 0 ? fmtKo(diff) + ' 초과' : diff < 0 ? fmtKo(-diff) + ' 남김' : '정확히 맞춤'}</span></div>
    <p>목표 <b class="num">${fmt(total)}</b> · 실제 <b class="num">${fmt(actual)}</b> · 머문 슬라이드 ${per.filter(x => x > 0).length}/${n}장</p>
    <div class="tbl-wrap"><table><thead><tr><th>장</th><th>제목</th><th>목표</th><th>실제</th><th>차이</th></tr></thead><tbody>${rows}</tbody></table></div>
    <div class="row"><button id="rhAgain" type="button" class="primary">다시 리허설</button><span class="progress-note">이 브라우저의 내 리허설 기록에 저장했습니다</span></div></div>`;
  $('#rhAgain').onclick = () => { R.reset(); $('#rhResult').innerHTML = ''; renderRehearsal(); window.scrollTo({top:0, behavior:'smooth'}); };
  $('#rhResult').scrollIntoView({behavior:'smooth', block:'start'});
  if(S.pid){ S.history = [{at: Date.now(), target: total, actual, slides: n, per}, ...S.history].slice(0, 10); LS.set('hist:' + S.pid, S.history); renderHistory(); }
}
function renderHistory(){
  const el = $('#rhHistory');
  if(!S.history.length){ el.innerHTML = ''; return; }
  el.innerHTML = `<div class="card" style="display:flex;flex-direction:column;gap:10px"><h3>내 최근 리허설 <small style="color:var(--muted);font-weight:400">이 브라우저에만 저장됩니다</small></h3><div class="tbl-wrap"><table class="tbl"><thead><tr><th>일시</th><th>목표</th><th>실제</th><th>차이</th></tr></thead><tbody>${
    S.history.map(r => { const d = r.actual - r.target; return `<tr><td>${dt(r.at)}</td><td class="num">${fmt(r.target)}</td><td class="num">${fmt(r.actual)}</td><td class="num ${d > 15 ? 'd-late' : d < -15 ? 'd-early' : ''}">${d > 0 ? '+' : d < 0 ? '-' : ''}${fmt(Math.abs(d))}</td></tr>`; }).join('')
  }</tbody></table></div></div>`;
}

/* ───────── 2. 스크립트 ───────── */
let scEditing = null, thumbObserver = null;
function renderScripts(){
  const list = $('#scList');
  if(!S.project){ list.innerHTML = noProjectHTML(); $('#scSummary').textContent = ''; return; }
  if(scEditing) return; // 편집 중에는 다시 그리지 않음
  const n = S.slides.length || Deck.pages();
  if(!n){ list.innerHTML = '<div class="empty"><b>발표자료가 아직 없습니다</b>사업 자료 관리 탭에서 발표자료(PPTX)를 올리면 슬라이드별 스크립트 칸이 생깁니다.</div>'; $('#scSummary').textContent = ''; return; }
  const done = Array.from({length:n}, (_, i) => (S.scripts[i+1] && (S.scripts[i+1].body||'').trim()) ? 1 : 0).reduce((a,b) => a+b, 0);
  const totalEst = Array.from({length:n}, (_, i) => estSec((S.scripts[i+1]||{}).body)).reduce((a,b) => a+b, 0);
  const tgt = S.project.default_min * 60 + (S.project.default_sec || 0);
  $('#scSummary').textContent = `${n}장 중 ${done}장 작성 · 스크립트를 읽으면 약 ${fmtKo(totalEst)} (발표 시간 ${fmtKo(tgt)}) · 공백 제외 분당 ${CPM}자 기준`;
  const term = $('#scSearch').value.trim().toLowerCase(), f = $('#scFilter').value;
  const visual = !!(Deck.pdf || Deck.holder);
  const items = [];
  for(let no = 1; no <= n; no++){
    const s = S.slides[no-1] || {}; const sc = S.scripts[no]; const txt = (sc && sc.body || '').trim();
    if(f === 'empty' && txt) continue; if(f === 'done' && !txt) continue;
    if(term && !((s.title||'') + ' ' + txt + ' ' + (s.body||'')).toLowerCase().includes(term)) continue;
    const thumb = visual ? `<div class="sc-thumb" data-thumb="${no}">불러오는 중</div>`
      : `<div class="sc-thumb txt"><b>${esc(s.title || '슬라이드 ' + no)}</b>${esc((s.body||'').slice(0,160))}</div>`;
    items.push(`<article class="sc-item" data-no="${no}">${thumb}<div class="sc-body">
      <div class="sc-top"><h3><span>${pad(no)}</span>${esc(s.title || '제목 없음')}</h3><div class="row">
        <button type="button" class="ghost" data-act="rh" data-no="${no}">이 장부터 리허설</button>
        <button type="button" data-act="edit" data-no="${no}">${txt ? '수정' : '작성'}</button></div></div>
      <div class="sc-text ${txt ? '' : 'none'}" id="sct-${no}">${txt ? esc(txt) : '아직 작성된 스크립트가 없습니다.'}</div>
      ${txt ? `<div class="sc-meta"><span>${nonSpace(txt)}자 · 약 ${fmtKo(estSec(txt))}</span><span>작성 ${esc(sc.author || '이름 없음')} · ${dt(sc.updated_at)}</span>${sc.src === 'pptx-notes' ? '<span class="pill">PPT 노트에서 가져옴</span>' : ''}</div>` : ''}
    </div></article>`);
  }
  list.innerHTML = items.length ? items.join('') : '<div class="empty"><b>조건에 맞는 슬라이드가 없습니다</b>검색어나 표시 범위를 바꿔 보세요.</div>';
  if(visual){
    if(thumbObserver) thumbObserver.disconnect();
    thumbObserver = 'IntersectionObserver' in window ? new IntersectionObserver(es => es.forEach(e => { if(e.isIntersecting){ fillThumb(e.target); thumbObserver.unobserve(e.target); } }), {rootMargin:'400px'}) : null;
    list.querySelectorAll('[data-thumb]').forEach(el => thumbObserver ? thumbObserver.observe(el) : fillThumb(el));
  }
}
function editScript(no){
  const el = $('#sct-' + no); if(!el) return;
  scEditing = no;
  const cur = (S.scripts[no] && S.scripts[no].body) || '';
  el.outerHTML = `<div class="field" id="sce-${no}"><textarea id="sci-${no}" rows="8" aria-label="슬라이드 ${no} 스크립트">${esc(cur)}</textarea>
    <div class="row" style="justify-content:space-between"><span class="progress-note" id="scc-${no}"></span><span class="row"><button type="button" data-act="cancel">취소</button><button type="button" class="primary" data-act="save" data-no="${no}">저장</button></span></div></div>`;
  const ta = $('#sci-' + no); const upd = () => { $('#scc-' + no).textContent = `${nonSpace(ta.value)}자 · 읽는 데 약 ${fmtKo(estSec(ta.value))}`; };
  ta.addEventListener('input', upd); upd(); ta.focus();
}
async function saveScript(no){
  const ta = $('#sci-' + no); if(!ta) return;
  const body = ta.value.replace(/\s+$/,'');
  try{
    await W(() => q(sb.from('scripts').upsert({project_id: S.pid, no, body, author: me(), src:'manual', updated_at: new Date().toISOString()})), `슬라이드 ${no} 스크립트를 저장했습니다.`);
    scEditing = null; await loadTable('scripts');
  }catch(e){}
}

/* ───────── 3. 질의 목록 ───────── */
const isAuto = x => x.src === 'auto';
function srcMatch(x, f){ return !f || (f === 'team' ? !isAuto(x) : f === 'auto' ? isAuto(x) : f === 'noanswer' ? !(x.a || '').trim() : true); }
function slideHint(page){
  const n = parseInt(String(page || ''), 10); if(!n || /제안서/.test(String(page))) return '';
  const s = S.slides[n - 1]; if(!s) return '';
  const sc = (S.scripts[n] && S.scripts[n].body || '').trim();
  const lines = String(s.body || '').split('\n').filter(Boolean).slice(0, 4).join(' / ');
  return `발표자료 p.${n} '${s.title || ''}'` + (lines ? `: ${lines}` : '') + (sc ? `\n스크립트: ${sc.length > 160 ? sc.slice(0, 160) + '…' : sc}` : '');
}
function sourcePills(x){
  return (isAuto(x) ? '<span class="pill">자동 생성</span>' : x.src === 'ai' ? '<span class="pill warn">AI 생성 · 검토 필요</span>' : '<span class="pill ok">우리 팀</span>')
    + ((x.a || '').trim() ? '' : '<span class="pill bad">답변 필요</span>');
}
let qaPanel = null, aiItems = null;
function catsInUse(){ const s = new Set(DEFAULT_CATS); S.qa.forEach(x => x.cat && s.add(x.cat)); return [...s]; }
function pageNum(p){ const m = String(p||'').match(/\d+/); return m ? +m[0] : 9999; }
function sortQA(list, how){
  const a = [...list];
  if(how === 'new') a.sort((x,y) => tms(y.created_at) - tms(x.created_at));
  else if(how === 'cat') a.sort((x,y) => (x.cat||'').localeCompare(y.cat||'') || pageNum(x.page) - pageNum(y.page));
  else a.sort((x,y) => pageNum(x.page) - pageNum(y.page) || tms(x.created_at) - tms(y.created_at));
  return a;
}
function qaMetaPills(x){
  return [x.cat ? `<span class="pill acc">${esc(x.cat)}</span>` : '',
    x.page ? `<span class="pill">${/^\d/.test(String(x.page)) ? 'p.' : ''}${esc(x.page)}</span>` : '',
    x.rfp ? `<span class="pill">${esc(x.rfp)}</span>` : ''].join('');
}
function renderQA(){
  const list = $('#qaList');
  ['qaAddBtn','qaPasteBtn','qaAutoBtn','qaAiBtn','qaCsvBtn'].forEach(id => $('#'+id).disabled = !S.project);
  if(!S.project){ list.innerHTML = noProjectHTML(); $('#qaSummary').textContent = ''; $('#qaPanel').innerHTML = ''; return; }
  const catSel = $('#qaCat'); const curCat = catSel.value;
  const cats = [...new Set(S.qa.map(x => x.cat).filter(Boolean))];
  catSel.innerHTML = '<option value="">모든 구분</option>' + cats.map(c => `<option${c === curCat ? ' selected' : ''}>${esc(c)}</option>`).join('');
  const byCat = {}; S.qa.forEach(x => byCat[x.cat || '미분류'] = (byCat[x.cat || '미분류'] || 0) + 1);
  const autoN = S.qa.filter(isAuto).length, noA = S.qa.filter(x => !(x.a || '').trim()).length;
  $('#qaSummary').textContent = `총 ${S.qa.length}개 · 우리 팀 ${S.qa.length - autoN} · 자동 생성 ${autoN} · 답변 필요 ${noA}` + (S.qa.length ? ' · ' + Object.entries(byCat).map(([k,v]) => `${k} ${v}`).join(' · ') : '');
  renderQAPanel();
  const term = $('#qaSearch').value.trim().toLowerCase(), cat = catSel.value, src = $('#qaSrc').value, showA = $('#qaShowA').checked;
  let rows = S.qa.filter(x => (!cat || x.cat === cat) && srcMatch(x, src) && (!term || [x.q,x.a,x.cat,x.kw,x.rfp,x.page].join(' ').toLowerCase().includes(term)));
  rows = sortQA(rows, $('#qaSort').value);
  if(!S.qa.length){ list.innerHTML = `<div class="empty"><b>아직 질의가 없습니다</b>${S.slides.length ? "'발표자료로 자동 생성'을 누르거나, 우리 질의를 직접 넣거나 엑셀에서 붙여넣으세요." : '사업 자료 관리 탭에서 발표자료를 올리면 예상 질의가 자동으로 만들어집니다. 우리 질의는 직접 넣거나 엑셀에서 붙여넣으세요.'}</div>`; return; }
  list.innerHTML = rows.length ? rows.map(x => `<article class="qa">
      <div class="qa-head">${qaMetaPills(x)}${sourcePills(x)}</div>
      <div class="qa-q">${esc(x.q)}</div>
      ${x.a ? `<div class="qa-a ${showA ? '' : 'hidden'}" id="qaa-${esc(x.id)}">${esc(x.a)}</div>` : `<div class="progress-note" style="white-space:pre-wrap">답변이 아직 없습니다. 수정을 눌러 답변을 적어 두세요.${slideHint(x.page) ? '\n참고 ' + esc(slideHint(x.page)) : ''}</div>`}
      <div class="qa-foot"><span>${x.kw ? '키워드: ' + esc(x.kw) + ' · ' : ''}작성 ${esc(x.author || '이름 없음')} · ${dt(x.created_at)}</span>
        <span class="acts">${x.a ? `<button type="button" class="ghost" data-act="toggleA" data-id="${esc(x.id)}">${showA ? '답변 가리기' : '답변 보기'}</button>` : ''}
        <button type="button" class="ghost" data-act="editQ" data-id="${esc(x.id)}">수정</button><button type="button" class="ghost" data-act="delQ" data-id="${esc(x.id)}">삭제</button></span></div>
    </article>`).join('') : '<div class="empty"><b>조건에 맞는 질의가 없습니다</b>필터를 바꿔 보세요.</div>';
}
function qaFormHTML(x){
  x = x || {};
  return `<div class="form-grid">
    <div class="field"><label for="fCat">구분</label><input id="fCat" type="text" list="catList" value="${esc(x.cat||'')}" placeholder="예: 기술·아키텍처"><datalist id="catList">${catsInUse().map(c => `<option value="${esc(c)}">`).join('')}</datalist></div>
    <div class="field"><label for="fPage">해당 페이지</label><input id="fPage" type="text" value="${esc(x.page||'')}" placeholder="예: 12 또는 12, 14"></div>
    <div class="field"><label for="fRfp">RFP 요구사항 번호</label><input id="fRfp" type="text" value="${esc(x.rfp||'')}" placeholder="예: SFR-012"></div>
    <div class="field"><label for="fKw">검색 키워드</label><input id="fKw" type="text" value="${esc(x.kw||'')}" placeholder="쉼표로 구분"></div>
  </div>
  <div class="field"><label for="fQ">질문</label><textarea id="fQ" rows="2">${esc(x.q||'')}</textarea></div>
  <div class="field"><label for="fA">답변 (발표자가 그대로 말할 수 있게)</label><textarea id="fA" rows="5">${esc(x.a||'')}</textarea></div>`;
}
function renderQAPanel(){
  const el = $('#qaPanel');
  if(!qaPanel){ el.innerHTML = ''; el.dataset.key = ''; return; }
  const key = JSON.stringify(qaPanel);
  if(el.dataset.key === key && el.innerHTML) return; // 입력 중인 내용 보존
  el.dataset.key = key;
  if(qaPanel.type === 'add' || qaPanel.type === 'edit'){
    const x = qaPanel.type === 'edit' ? S.qa.find(y => y.id === qaPanel.id) : null;
    el.innerHTML = `<div class="panel"><div class="panel-head"><h3>${x ? '질의 수정' : '질의 추가'}</h3><button type="button" class="ghost" data-act="closePanel">닫기</button></div>${qaFormHTML(x)}
      <div class="row" style="justify-content:flex-end"><button type="button" data-act="closePanel">취소</button><button type="button" class="primary" data-act="saveQ">${x ? '수정 저장' : '추가'}</button></div></div>`;
    $('#fQ').focus();
  } else if(qaPanel.type === 'paste'){
    el.innerHTML = `<div class="panel"><div class="panel-head"><h3>엑셀에서 붙여넣기</h3><button type="button" class="ghost" data-act="closePanel">닫기</button></div>
      <p class="progress-note">엑셀에서 표를 선택해 복사(Ctrl+C)한 뒤 아래에 붙여넣으세요. 첫 줄에 '구분 · 질문 · 해당 페이지 · 답변' 같은 제목이 있으면 열을 알아서 맞춥니다. 제목이 없으면 구분, 질문, 해당 페이지, 답변, RFP 번호, 키워드 순서로 읽습니다.</p>
      <textarea id="pasteBox" rows="8" placeholder="여기에 붙여넣기"></textarea>
      <div id="pastePreview" class="progress-note"></div>
      <div class="row" style="justify-content:flex-end"><button type="button" data-act="closePanel">취소</button><button type="button" class="primary" data-act="doPaste" id="pasteGo" disabled>추가</button></div></div>`;
    $('#pasteBox').addEventListener('input', () => {
      const rows = parseQAInput($('#pasteBox').value);
      $('#pastePreview').innerHTML = rows.length ? `질의 <b>${rows.length}개</b>를 읽었습니다. 첫 질의: ${esc(rows[0].q.slice(0,80))}` : '읽을 수 있는 질의가 없습니다.';
      $('#pasteGo').disabled = !rows.length; $('#pasteGo').textContent = rows.length ? `${rows.length}개 추가` : '추가';
    });
    $('#pasteBox').focus();
  } else if(qaPanel.type === 'auto'){
    const untouched = S.qa.filter(x => isAuto(x) && !(x.a || '').trim() && !x.updated_at).length;
    el.innerHTML = `<div class="panel"><div class="panel-head"><h3>발표자료로 예상 질의 자동 생성</h3><button type="button" class="ghost" data-act="closePanel">닫기</button></div>
      <p class="progress-note">AI 없이 규칙으로 만듭니다. 슬라이드의 RFP 요구사항 번호, 수치, '최초·유일' 같은 주장 표현, 보안·연계·일정·데이터 같은 주제어, 기술 용어를 읽어 평가위원이 물을 만한 질문을 만듭니다. 답변은 비워 두고, 답변 준비에 쓸 해당 슬라이드 내용을 함께 보여 줍니다.</p>
      <div class="form-grid">
        <div class="field"><label for="autoMax">최대 개수</label><select id="autoMax"><option>20</option><option selected>40</option><option>60</option></select></div>
        <div class="field"><label for="autoPer">슬라이드당 최대</label><select id="autoPer"><option>1</option><option>2</option><option selected>3</option></select></div>
      </div>
      <label class="chk"><input type="checkbox" id="autoProp" checked> 정성제안서에만 있는 RFP 요구사항 번호도 질의로 만들기</label>
      <p class="progress-note">다시 생성하면 답변을 적지 않은 자동 생성 질의 ${untouched}개를 새로 바꿉니다. 답변을 적었거나 수정한 질의와 우리 팀 질의는 그대로 둡니다.</p>
      <div class="row"><button type="button" class="primary" data-act="autoGo" ${S.slides.length ? '' : 'disabled'}>질의 생성</button><span class="progress-note" id="autoProg">${S.slides.length ? '' : '먼저 사업 자료 관리 탭에서 발표자료를 올려 주세요.'}</span></div></div>`;
  } else if(qaPanel.type === 'ai'){
    el.innerHTML = `<div class="panel"><div class="panel-head"><h3>AI 질의 생성용 프롬프트</h3><button type="button" class="ghost" data-act="closePanel">닫기</button></div>
      <p class="progress-note">이 사이트는 API 키 없이 동작하도록 AI를 직접 부르지 않습니다. 아래 버튼으로 발표자료와 정성제안서 내용이 담긴 프롬프트를 복사해 Claude 대화창에 붙여넣고, 받은 결과를 그대로 아래 칸에 붙여넣으면 질의 목록에 들어갑니다. 답변 속 자사 표기는 '${esc(S.project.self_name || '제안사')}'입니다.</p>
      <div class="form-grid">
        <div class="field"><label for="aiCount">개수</label><select id="aiCount"><option>10</option><option selected>15</option><option>20</option><option>30</option></select></div>
        <div class="field"><label for="aiFocus">집중할 키워드나 RFP 번호 (선택)</label><input id="aiFocus" type="text" placeholder="예: 데이터 품질, SFR-012"></div>
      </div>
      <div class="field"><label>포함할 관점</label><div class="chips" id="aiAngles">${['핵심 사업 이해','기술 구현 상세','경쟁사 대비 차별화','일정·인력·리스크','데이터·보안','기대효과·정량 목표'].map((a,i) => `<label class="chip ${i < 4 ? 'on' : ''}"><input type="checkbox" ${i < 4 ? 'checked' : ''} value="${a}" style="display:none">${a}</label>`).join('')}</div></div>
      <div class="row"><button type="button" class="primary" data-act="aiCopy">1. 프롬프트 복사</button><span class="progress-note" id="aiProg"></span></div>
      <div class="field"><label for="aiPaste">2. Claude가 준 결과를 붙여넣기</label><textarea id="aiPaste" rows="6" class="prompt-box" placeholder='[{"cat":"...","q":"...","a":"...","page":"...","rfp":"...","kw":"..."}]'></textarea></div>
      <div id="aiOut"></div></div>`;
    el.querySelectorAll('#aiAngles input').forEach(i => i.addEventListener('change', () => i.parentElement.classList.toggle('on', i.checked)));
    $('#aiPaste').addEventListener('input', () => { const items = parseQAInput($('#aiPaste').value); aiItems = items.length ? items.map(x => ({...x, pick:true})) : null; renderAiItems(); });
  }
}
async function saveQAForm(){
  const v = id => $('#' + id).value.trim();
  const body = {cat: v('fCat'), q: v('fQ'), a: v('fA'), page: v('fPage'), rfp: v('fRfp'), kw: v('fKw')};
  if(!body.q){ toast('질문을 입력하세요.'); $('#fQ').focus(); return; }
  try{
    if(qaPanel.type === 'edit') await W(() => q(sb.from('qa').update({...body, updated_at: new Date().toISOString()}).eq('id', qaPanel.id)), '질의를 수정했습니다.');
    else await W(() => q(sb.from('qa').insert({...body, project_id: S.pid, author: me(), src: 'manual'})), '질의를 추가했습니다.');
    qaPanel = null; await loadTable('qa');
  }catch(e){}
}
async function addMany(items, src){
  const rows = items.map(it => ({project_id: S.pid, cat: it.cat||'', q: it.q, a: it.a||'', page: String(it.page||''), rfp: it.rfp||'', kw: it.kw||'', author: me(), src}));
  await W(() => q(sb.from('qa').insert(rows)));
  await loadTable('qa');
  return rows.length;
}
function pickRelevant(text, focus, cap){
  if(text.length <= cap) return text;
  const parts = text.split(/(?=\[제안서 p\.\d+\])/);
  if(!focus || parts.length < 2) return text.slice(0, cap);
  const toks = focus.toLowerCase().split(/[\s,]+/).filter(Boolean);
  const scored = parts.map((t,i) => ({i, t, s: toks.reduce((a,k) => a + (t.toLowerCase().split(k).length - 1), 0)}));
  const chosen = new Set(); let len = 0;
  for(const x of [...scored].sort((a,b) => b.s - a.s)){ if(x.s === 0 || len + x.t.length > cap) continue; chosen.add(x.i); len += x.t.length; }
  for(const x of scored){ if(len + x.t.length > cap) break; if(!chosen.has(x.i)){ chosen.add(x.i); len += x.t.length; } }
  return scored.filter(x => chosen.has(x.i)).map(x => x.t).join('');
}
async function buildPrompt(){
  const count = +$('#aiCount').value, focus = $('#aiFocus').value.trim();
  const angles = [...document.querySelectorAll('#aiAngles input:checked')].map(i => i.value);
  const p = S.project, self = p.self_name || '제안사';
  let propText = '';
  try{ const r = await q(sb.from('proposal_text').select('body').eq('project_id', S.pid).maybeSingle()); propText = (r && r.body) || ''; }catch(e){}
  const deck = S.slides.map(s => `[발표자료 ${s.no}p] ${s.title}\n${s.body}`).join('\n\n').slice(0, 45000);
  const prop = pickRelevant(propText, focus, 70000);
  const existing = S.qa.map(x => '- ' + x.q).join('\n').slice(0, 8000);
  return `너는 공공기관 정보화 사업 제안 발표의 평가위원이다. 아래 [발표자료]와 [정성제안서]만 읽고, 발표 직후 질의응답 시간에 평가위원이 실제로 물을 질의와, 발표자가 그 자리에서 바로 말할 수 있는 답변을 만든다.

규칙
- 질의의 근거는 [발표자료]와 [정성제안서]에 적힌 내용뿐이다. 자료에 없는 사실·수치·제품명을 지어내지 않고, 내부 참고자료(솔루션 제조사 자료, 사내 문서 등)를 언급하지 않는다.
- 질문은 평가위원 말투로 한두 문장. 모호한 표현, 근거 없는 수치, 일정·인력의 현실성, 요구사항 충족 여부처럼 실제로 파고들 지점을 고른다.
- 답변 화자는 발표자다. 회사 이름 대신 "${self}"라고 쓴다. 첫 문장에 결론을 말하고 자료 속 근거 1~2개를 붙여 30~60초에 말할 수 있는 구어체로 쓴다.
- 자료로 답할 수 없는 질의는 답변을 "[보완 필요] "로 시작하고 무엇을 준비해야 하는지 쓴다.
- page: 근거가 있는 발표자료 페이지 번호를 "12" 또는 "12, 14"처럼 쓴다. 발표자료에 없고 제안서에만 있으면 "제안서 p.N".
- rfp: 자료에 요구사항 ID(예: SFR-001)가 나오면 그 ID, 없으면 빈 문자열.
- kw: 검색용 키워드 2~4개를 쉼표로.
- cat: ${catsInUse().join(', ')} 중 하나.
- 관점: ${angles.join(', ') || '전체'}${focus ? `\n- 이번에는 "${focus}"와 관련된 질의에 집중한다.` : ''}${p.competitor_memo ? '\n- [경쟁사·참고 메모]는 차별화 검증 질의를 고르는 데에만 쓰고, 메모 문구를 질의·답변에 그대로 옮기지 않는다.' : ''}
- [이미 있는 질의]와 겹치는 질문은 만들지 않는다.
- 정확히 ${count}개.

출력: JSON 배열만 출력한다. 설명이나 코드블록 표시 없이 [{"cat":"","q":"","a":"","page":"","rfp":"","kw":""}, ...]

[발표자료]
${deck || '(없음)'}

[정성제안서]
${prop || '(없음)'}
${p.competitor_memo ? `\n[경쟁사·참고 메모]\n${p.competitor_memo.slice(0,4000)}\n` : ''}
[이미 있는 질의]
${existing || '(없음)'}`;
}
function renderAiItems(){
  const out = $('#aiOut'); if(!out) return;
  if(!aiItems){ out.innerHTML = $('#aiPaste') && $('#aiPaste').value.trim() ? '<p class="progress-note">결과를 읽지 못했습니다. Claude가 준 JSON 배열 전체를 붙여넣었는지 확인하세요.</p>' : ''; return; }
  out.innerHTML = `<p class="progress-note">${aiItems.length}개를 읽었습니다. 넣을 질의를 고르세요.</p><div class="gen-list">${aiItems.map((x,i) => `<label class="gen-item"><input type="checkbox" data-gi="${i}" ${x.pick ? 'checked' : ''}><div><div class="row" style="gap:6px;margin-bottom:3px">${qaMetaPills(x)}</div><b>${esc(x.q)}</b><p>${esc(x.a)}</p></div></label>`).join('')}</div>
    <div class="row" style="justify-content:flex-end;margin-top:8px"><button type="button" data-act="aiAll">모두 선택</button><button type="button" data-act="aiNone">모두 해제</button><button type="button" class="primary" data-act="aiAdd">고른 질의 목록에 넣기</button></div>`;
  out.querySelectorAll('[data-gi]').forEach(c => c.addEventListener('change', () => aiItems[+c.dataset.gi].pick = c.checked));
}
async function autoGenerate({max = 40, perSlide = 3, includeProposal = true, silent = false} = {}){
  const pid = S.pid; if(!pid || !S.slides.length) return 0;
  await W(() => q(sb.from('qa').delete().eq('project_id', pid).eq('src', 'auto').eq('a', '').is('updated_at', null)));
  await loadTable('qa');
  let propText = '';
  if(includeProposal){ try{ const r = await q(sb.from('proposal_text').select('body').eq('project_id', pid).maybeSingle()); propText = (r && r.body) || ''; }catch(e){} }
  const items = window.RuleQA.generate({slides: S.slides, proposalText: propText, existing: S.qa.map(x => x.q), self: S.project.self_name || '제안사', max, perSlide, includeProposal});
  if(!items.length){ if(!silent) toast('새로 만들 질의가 없습니다.'); return 0; }
  const rows = items.map(it => ({project_id: pid, cat: it.cat, q: it.q, a: '', page: it.page, rfp: it.rfp, kw: it.kw, author: '자동 생성', src: 'auto'}));
  await W(() => q(sb.from('qa').insert(rows)));
  await loadTable('qa');
  return rows.length;
}
function exportQA(){
  const list = sortQA(S.qa, 'page');
  if(!list.length){ toast('내보낼 질의가 없습니다.'); return; }
  const head = ['구분','질문','답변','해당 페이지','RFP 요구사항','키워드','작성자'];
  const rows = list.map(x => [x.cat, x.q, x.a, x.page, x.rfp, x.kw, x.author]);
  const csv = '﻿' + [head, ...rows].map(r => r.map(c => `"${String(c ?? '').replace(/"/g,'""')}"`).join(',')).join('\r\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], {type:'text/csv;charset=utf-8'}));
  a.download = `${(S.project.name || '사업').replace(/[\\/:*?"<>|]/g,'_')}_예상질의.csv`;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

/* ───────── 4. 질의 연습 ───────── */
const D = {list:[], i:0, start:0, revealed:false, timer:0, answerSec:null, results:{}, summary:null};
function tokens(s){ return s.toLowerCase().split(/[\s,·]+/).filter(Boolean); }
function searchQA(query){
  const toks = tokens(query);
  if(!toks.length) return sortQA(S.qa, 'page').map(x => ({x, s:0}));
  const out = [];
  for(const x of S.qa){
    let score = 0, all = true;
    const pages = String(x.page||'').match(/\d+/g) || [];
    for(const t of toks){
      const pm = t.match(/^(?:p\.?|페이지)?(\d{1,3})(?:p|쪽|페이지)?$/);
      let s = 0;
      if(pm && pages.includes(String(+pm[1]))) s += 6;
      if((x.rfp||'').toLowerCase().includes(t)) s += 5;
      if((x.kw||'').toLowerCase().includes(t)) s += 4;
      if((x.q||'').toLowerCase().includes(t)) s += 3;
      if((x.cat||'').toLowerCase().includes(t)) s += 2;
      if((x.a||'').toLowerCase().includes(t)) s += 1;
      if(!s){ all = false; break; }
      score += s;
    }
    if(all) out.push({x, s:score});
  }
  return out.sort((a,b) => b.s - a.s || pageNum(a.x.page) - pageNum(b.x.page));
}
function hl(text, toks){
  let h = esc(text);
  toks.filter(t => t.length > 1 || /\d/.test(t)).forEach(t => { const re = new RegExp('(' + esc(t).replace(/[.*+?^${}()|[\]\\]/g,'\\$&') + ')', 'gi'); h = h.replace(re, '<mark>$1</mark>'); });
  return h;
}
let prCat = '';
function practiceSet(){
  const mark = $('#prMark').value;
  const src = $('#prSrc').value;
  return searchQA($('#prQ').value.trim()).map(r => r.x).filter(x => srcMatch(x, src) && (!prCat || x.cat === prCat) && (!mark || (mark === 'none' ? !S.marks[x.id] : S.marks[x.id] === mark)));
}
function renderPractice(){
  if(!$('#prDrill').classList.contains('hidden')) return;
  const res = $('#prResults');
  if(!S.project){ res.innerHTML = noProjectHTML(); $('#prCount').textContent = ''; $('#prStart').disabled = true; $('#prCats').innerHTML = ''; return; }
  const cats = [...new Set(S.qa.map(x => x.cat).filter(Boolean))];
  $('#prCats').innerHTML = cats.length ? `<button type="button" class="chip ${prCat ? '' : 'on'}" data-cat="">전체 ${S.qa.length}</button>` + cats.map(c => `<button type="button" class="chip ${prCat === c ? 'on' : ''}" data-cat="${esc(c)}">${esc(c)} ${S.qa.filter(x => x.cat === c).length}</button>`).join('') : '';
  const set = practiceSet(); const toks = tokens($('#prQ').value);
  const okN = S.qa.filter(x => S.marks[x.id] === 'ok').length, agN = S.qa.filter(x => S.marks[x.id] === 'again').length;
  $('#prCount').textContent = `${set.length}개 질의` + (S.qa.length ? ` · 내 기록: 잘 답함 ${okN} · 다시 연습 ${agN}` : '');
  $('#prStart').disabled = !set.length;
  if(!S.qa.length){ res.innerHTML = '<div class="empty"><b>연습할 질의가 아직 없습니다</b>질의 목록 탭에 질의가 모이면 여기서 검색해 연습할 수 있습니다.</div>'; return; }
  if(!set.length){ res.innerHTML = '<div class="empty"><b>검색 결과가 없습니다</b>다른 키워드나 RFP 번호, 페이지 번호(예: p.12)로 찾아보세요.</div>'; return; }
  res.innerHTML = set.slice(0, 200).map(x => `<article class="qa">
    <div class="qa-head">${qaMetaPills(x)}${sourcePills(x)}${S.marks[x.id] === 'ok' ? '<span class="pill ok">잘 답함</span>' : S.marks[x.id] === 'again' ? '<span class="pill warn">다시 연습</span>' : ''}</div>
    <div class="qa-q">${hl(x.q, toks)}</div>
    <div class="qa-a hidden" id="pra-${esc(x.id)}">${x.a ? hl(x.a, toks) : esc('답변이 아직 없습니다.' + (slideHint(x.page) ? '\n참고 ' + slideHint(x.page) : ''))}</div>
    <div class="qa-foot"><span>${x.kw ? '키워드: ' + hl(x.kw, toks) : ''}</span><span class="acts"><button type="button" class="ghost" data-act="prToggle" data-id="${esc(x.id)}">답변 보기</button><button type="button" data-act="prOne" data-id="${esc(x.id)}">이 질의만 연습</button></span></div>
  </article>`).join('') + (set.length > 200 ? `<p class="progress-note">상위 200개만 보여 줍니다. 검색어를 좁혀 보세요.</p>` : '');
}
function startDrill(list){
  if(!list.length) return;
  if($('#prOrder').value === 'random'){ list = [...list]; for(let i = list.length - 1; i > 0; i--){ const j = Math.floor(Math.random() * (i + 1)); [list[i], list[j]] = [list[j], list[i]]; } }
  D.list = list; D.i = 0; D.results = {}; D.summary = null;
  $('#prSearchView').classList.add('hidden'); $('#prDrill').classList.remove('hidden');
  showDrill();
}
function showDrill(){
  clearInterval(D.timer);
  if(D.i >= D.list.length) return drillSummary();
  const x = D.list[D.i]; D.revealed = false; D.start = performance.now(); D.answerSec = null;
  $('#drPos').textContent = `${D.i + 1} / ${D.list.length}`;
  $('#drMeta').innerHTML = qaMetaPills(x);
  $('#drQ').textContent = x.q;
  const a = $('#drA'); a.classList.add('hidden');
  const hint = slideHint(x.page);
  a.innerHTML = esc(x.a || ('등록된 답변이 없습니다. 말해 본 답변을 질의 목록에 적어 두세요.' + (hint ? '\n\n참고 ' + hint : ''))) + (x.page ? `<span class="src">근거: 발표자료 ${/^\d/.test(x.page) ? 'p.' : ''}${esc(x.page)}${x.rfp ? ' · ' + esc(x.rfp) : ''}</span>` : '');
  $('#drReveal').textContent = '답변 보기 (Space)';
  $('#drPrev').disabled = D.i === 0;
  $('#drLimitTxt').textContent = '목표 ' + fmtKo(+$('#prLimit').value);
  drillTick(); D.timer = setInterval(drillTick, 200);
}
function drillTick(){
  const lim = +$('#prLimit').value;
  const sec = D.answerSec != null ? D.answerSec : (performance.now() - D.start) / 1000;
  $('#drTime').textContent = fmt(sec);
  const bar = $('#drBar'); bar.style.width = Math.min(100, sec / lim * 100) + '%'; bar.classList.toggle('over', sec > lim);
}
function revealDrill(){
  const a = $('#drA');
  if(!D.revealed){ D.revealed = true; if(D.answerSec == null) D.answerSec = (performance.now() - D.start) / 1000; clearInterval(D.timer); drillTick(); a.classList.remove('hidden'); $('#drReveal').textContent = '답변 가리기 (Space)'; }
  else { D.revealed = false; a.classList.add('hidden'); $('#drReveal').textContent = '답변 보기 (Space)'; }
}
function markDrill(v){
  const x = D.list[D.i]; if(!x) return;
  S.marks = {...S.marks, [x.id]: v}; D.results[x.id] = v;
  if(S.pid) LS.set('marks:' + S.pid, S.marks);
  D.i++; showDrill();
}
function drillSummary(){
  clearInterval(D.timer);
  const vals = Object.values(D.results);
  const ok = vals.filter(v => v === 'ok').length, again = vals.filter(v => v === 'again').length;
  $('#drPos').textContent = '연습 완료'; $('#drMeta').innerHTML = '';
  $('#drQ').textContent = `${D.list.length}개 중 잘 답함 ${ok}개, 다시 연습 ${again}개`;
  $('#drTime').textContent = '--:--'; $('#drBar').style.width = '0%';
  const a = $('#drA'); a.classList.remove('hidden');
  const againList = D.list.filter(x => D.results[x.id] === 'again');
  a.innerHTML = againList.length ? '다시 연습할 질의<br>' + againList.map(x => '· ' + esc(x.q)).join('<br>') : '표시한 질의가 모두 잘 답함입니다.';
  $('#drReveal').textContent = againList.length ? '다시 연습만 한 번 더' : '처음부터 다시';
  D.summary = againList.length ? againList : D.list;
}
function stopDrill(silent){
  clearInterval(D.timer); D.list = []; D.summary = null;
  $('#prDrill').classList.add('hidden'); $('#prSearchView').classList.remove('hidden');
  if(!silent) renderPractice();
}

/* ───────── 사업 자료 관리 ───────── */
const setupMsg = {deck:'', prop:''};
const pinOK = () => !CONFIG.ADMIN_PIN || sessionStorage.getItem('rr:pin') === String(CONFIG.ADMIN_PIN);
function renderSetup(){
  const el = $('#setupBody');
  if(!pinOK()){
    el.innerHTML = `<div class="card pin-box"><h3>관리 번호를 입력하세요</h3><p class="progress-note">사업 자료를 바꾸는 화면이라 관리 번호로 잠겨 있습니다.</p><input id="pinIn" type="password" inputmode="numeric" autocomplete="off"><button type="button" class="primary" data-act="pin">열기</button></div>`;
    $('#pinIn').focus(); $('#pinIn').addEventListener('keydown', e => { if(e.key === 'Enter') $('[data-act=pin]').click(); });
    return;
  }
  if(document.activeElement && el.contains(document.activeElement) && /INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName)) return; // 입력 중 보존
  const p = S.project, deck = p && p.deck, prop = p && p.proposal;
  el.innerHTML = `
  <div class="card setup-sec">
    <h2>사업 자료 관리</h2>
    <ol class="steps"><li>사업을 만들고 발표 시간과 답변 속 자사 표기를 정합니다.</li><li>발표자료 PPTX를 올립니다. 슬라이드 화면은 PPTX 그대로 그리고, 슬라이드 노트는 스크립트로 가져옵니다. 화면이 원본과 다르게 보이면 같은 발표자료를 PDF로 저장해 함께 올리세요.</li><li>정성제안서를 올리면 원본은 저장소에, 글자는 AI 질의 생성 프롬프트에 씁니다.</li><li>이 사이트 주소를 팀에 공유합니다. 링크를 아는 사람은 누구나 보고 쓸 수 있습니다.</li></ol>
  </div>
  <div class="grid2">
    <div class="card setup-sec">
      <h3>새 사업 만들기</h3>
      <div class="field"><label for="newName">사업 이름</label><input id="newName" type="text" placeholder="예: 한국가스공사 Agentic AI 기반 구축"></div>
      <div class="row"><button type="button" class="primary" data-act="newProj">사업 만들기</button></div>
    </div>
    ${p ? `<div class="card setup-sec">
      <h3>현재 사업 기본값</h3>
      <div class="field"><label for="pName">사업 이름</label><input id="pName" type="text" value="${esc(p.name||'')}"></div>
      <div class="row"><div class="field"><label for="pMin">발표 시간</label><div class="mmss"><input id="pMin" type="number" min="0" max="180" value="${p.default_min ?? 20}"><span>분</span><input id="pSec" type="number" min="0" max="59" value="${p.default_sec ?? 0}"><span>초</span></div></div>
      <div class="field"><label for="pSelf">답변 속 자사 표기</label><input id="pSelf" type="text" value="${esc(p.self_name || '제안사')}" style="width:140px"></div></div>
      <div class="row"><button type="button" class="primary" data-act="saveProj">저장</button></div>
    </div>` : ''}
  </div>
  ${p ? `
  <div class="grid2">
    <div class="card setup-sec">
      <h3>발표자료</h3>
      <div class="file-status">${deck ? `<span>화면 표시: <b>${esc(deck.name)}</b> (${deck.kind === 'pdf' ? 'PDF' : 'PPTX'}) · <a href="${esc(deck.url)}" target="_blank" rel="noopener">원본 열기</a></span>` : '<span>화면 표시용 파일: 없음</span>'}${p.pptx_name ? `<span>PPTX: <b>${esc(p.pptx_name)}</b>${p.notes_count != null ? ` · 노트 있는 장 ${p.notes_count}` : ''}</span>` : ''}<span>저장된 슬라이드 정보: ${S.slides.length}장</span>${Deck.error ? '<span style="color:var(--bad)">슬라이드 화면을 그리지 못했습니다. PDF로 저장해 함께 올려 주세요.</span>' : ''}</div>
      <label class="drop" id="deckDrop" for="deckFile"><b>PPTX를 고르거나 끌어다 놓으세요</b><span class="hint">PDF를 함께 고르면 화면은 PDF로 보여 줍니다. 숨긴 슬라이드는 빠집니다.</span><input id="deckFile" type="file" accept=".pptx,.pdf" multiple></label>
      <div class="field"><label for="notesMode">PPTX 슬라이드 노트</label><select id="notesMode"><option value="empty">스크립트가 빈 슬라이드에만 가져오기</option><option value="all">모든 슬라이드 스크립트를 노트로 바꾸기</option><option value="skip">가져오지 않기</option></select></div>
      <span class="progress-note" id="deckProg">${esc(setupMsg.deck)}</span>
    </div>
    <div class="card setup-sec">
      <h3>정성제안서</h3>
      <div class="file-status">${prop ? `<span><b>${esc(prop.name)}</b>${prop.pages ? ` · ${prop.pages}쪽` : ''} · ${(prop.chars||0).toLocaleString()}자 · <a href="${esc(prop.url)}" target="_blank" rel="noopener">원본 열기</a></span><span class="hint">올린 날짜 ${dt(prop.at)}</span>` : '<span>없음</span>'}</div>
      <label class="drop" id="propDrop" for="propFile"><b>PPTX, PDF, DOCX 파일을 고르거나 끌어다 놓으세요</b><span class="hint">HWP는 PDF로 저장해서 올려 주세요.</span><input id="propFile" type="file" accept=".pptx,.pdf,.docx"></label>
      <span class="progress-note" id="propProg">${esc(setupMsg.prop)}</span>
    </div>
  </div>
  <div class="card setup-sec">
    <h3>경쟁사·참고 메모 <small style="color:var(--muted);font-weight:400">AI 질의 생성 프롬프트에만 들어가고 다른 화면에는 보이지 않습니다</small></h3>
    <textarea id="pMemo" rows="4" placeholder="예: 경쟁사 A는 현행 시스템 구축사. 강점은 현행 이해도, 약점은 확장성">${esc(p.competitor_memo||'')}</textarea>
    <div class="row"><button type="button" data-act="saveMemo">메모 저장</button></div>
  </div>
  <div class="card setup-sec">
    <h3>사업 삭제</h3>
    <p class="hint">이 사업의 발표자료·제안서 파일, 슬라이드 정보, 스크립트, 질의를 모두 지웁니다. 되돌릴 수 없습니다.</p>
    <div class="row"><button type="button" class="danger" data-act="delProj">'${esc(p.name || '이 사업')}' 삭제</button></div>
  </div>` : ''}`;
  if(p){
    bindDrop('deckDrop', 'deckFile', files => handleDeck(files));
    bindDrop('propDrop', 'propFile', files => handleProposal(files[0]));
  }
}
function bindDrop(dropId, inputId, cb){
  const d = $('#' + dropId), inp = $('#' + inputId);
  inp.addEventListener('change', () => { if(inp.files.length) cb([...inp.files]); inp.value = ''; });
  d.addEventListener('dragover', e => { e.preventDefault(); d.classList.add('over'); });
  d.addEventListener('dragleave', () => d.classList.remove('over'));
  d.addEventListener('drop', e => { e.preventDefault(); d.classList.remove('over'); if(e.dataTransfer.files.length) cb([...e.dataTransfer.files]); });
}
async function newProject(){
  const name = $('#newName').value.trim();
  if(!name){ toast('사업 이름을 입력하세요.'); $('#newName').focus(); return; }
  try{
    const row = await W(() => q(sb.from('projects').insert({name}).select().single()), `'${name}' 사업을 만들었습니다. 이제 발표자료를 올리세요.`);
    LS.set('pid', row.id); S.pid = null; $('#newName').value = '';
    await loadProjects();
  }catch(e){}
}
async function updateProject(fields, msg){ await W(() => q(sb.from('projects').update(fields).eq('id', S.pid)), msg); await loadProjects(); }
async function saveProject(){
  const fields = {name: $('#pName').value.trim() || S.project.name, default_min: Math.max(0, parseInt($('#pMin').value,10) || 0), default_sec: Math.min(59, Math.max(0, parseInt($('#pSec').value,10) || 0)), self_name: $('#pSelf').value.trim() || '제안사'};
  try{ await updateProject(fields, '저장했습니다.'); LS.set('rh:' + S.pid, null); applyRhSettings(); }catch(e){}
}
async function handleDeck(files){
  const pdfF = files.find(f => /\.pdf$/i.test(f.name)), pptF = files.find(f => /\.pptx$/i.test(f.name));
  if(!pdfF && !pptF){ toast('PPTX나 PDF 파일을 올려 주세요. 예전 형식(.ppt)은 PowerPoint에서 PPTX로 저장해 올려 주세요.', 5000); return; }
  const pid = S.pid, oldProject = S.project;
  const prog = m => { setupMsg.deck = m; const e = $('#deckProg'); if(e) e.textContent = m; };
  const mode = ($('#notesMode') || {}).value || 'empty';
  try{
    let pdfTexts = null, pptSlides = null, pdfPages = 0;
    if(pptF){ prog('PPTX를 읽는 중'); pptSlides = await parsePptx(pptF); }
    if(pdfF){
      prog('PDF를 읽는 중'); await ensureLibs('pdf');
      const pdf = await pdfjsLib.getDocument({data: new Uint8Array(await pdfF.arrayBuffer())}).promise;
      pdfPages = pdf.numPages;
      if(!pptSlides) pdfTexts = await pdfPageTexts(pdf, (i,n) => prog(`PDF 글자 읽는 중 ${i}/${n}`));
    }
    const shown = pdfF || pptF;
    prog('파일 올리는 중');
    const up = await W(() => uploadFile(pid, shown, 'deck'));
    const deck = {kind: pdfF ? 'pdf' : 'pptx', path: up.path, url: up.url, name: shown.name, size: shown.size, at: Date.now()};
    let pptxFile = null;
    if(pdfF && pptF){ pptxFile = await W(() => uploadFile(pid, pptF, 'deck-src')); deck.pptx = {path: pptxFile.path, url: pptxFile.url, name: pptF.name}; }
    const old = S.slides;
    let rows;
    if(pptSlides) rows = pptSlides.map((s,i) => ({project_id: pid, no: i+1, title: s.title, body: s.body, notes: s.notes, src: 'pptx', pptx_index: s.pptx_index}));
    else rows = Array.from({length: pdfPages}, (_, i) => {
      const o = old[i];
      if(o && o.src === 'pptx') return {project_id: pid, no: i+1, title: o.title, body: o.body, notes: o.notes, src: 'pptx', pptx_index: o.pptx_index};
      const ls = (pdfTexts[i] || '').split('\n');
      return {project_id: pid, no: i+1, title: (ls[0] || '').slice(0,200), body: ls.slice(1).join('\n').slice(0,5000), notes: '', src: 'pdf', pptx_index: null};
    });
    prog(`슬라이드 ${rows.length}장 저장 중`);
    await W(() => q(sb.from('slides').upsert(rows)));
    await W(() => q(sb.from('slides').delete().eq('project_id', pid).gt('no', rows.length)));
    let imported = 0;
    if(pptSlides && mode !== 'skip'){
      const sc = pptSlides.map((s,i) => ({no: i+1, notes: (s.notes||'').trim()})).filter(x => x.notes && (mode === 'all' || !((S.scripts[x.no] || {}).body || '').trim()))
        .map(x => ({project_id: pid, no: x.no, body: x.notes, author: me(), src: 'pptx-notes', updated_at: new Date().toISOString()}));
      if(sc.length){ prog('노트를 스크립트로 가져오는 중'); await W(() => q(sb.from('scripts').upsert(sc))); imported = sc.length; }
    }
    const fields = {deck, slide_count: rows.length};
    if(pptSlides){ fields.pptx_name = pptF.name; fields.notes_count = pptSlides.filter(s => (s.notes||'').trim()).length; }
    await W(() => q(sb.from('projects').update(fields).eq('id', pid)));
    const od = oldProject && oldProject.deck; if(od) removeFiles([od.path, od.pptx && od.pptx.path]);
    const mismatch = pdfF && pptSlides && pdfPages !== pptSlides.length ? ` PDF ${pdfPages}쪽과 PPTX ${pptSlides.length}장(숨김 제외)의 수가 다릅니다. 같은 파일에서 저장했는지 확인하세요.` : '';
    prog(`완료: 슬라이드 ${rows.length}장` + (imported ? `, 노트 ${imported}장을 스크립트로` : '') + mismatch);
    toast(`슬라이드 ${rows.length}장을 저장했습니다.` + mismatch, mismatch ? 7000 : 3500);
    await loadProjects(); await Promise.all(['slides','scripts'].map(loadTable));
    try{ const n = await autoGenerate({silent: true}); if(n){ prog(setupMsg.deck + ` · 예상 질의 ${n}개 자동 생성`); toast(`발표자료로 예상 질의 ${n}개를 자동으로 만들었습니다. 질의 목록 탭에서 확인하세요.`, 5000); } }catch(e){}
  }catch(e){ console.error(e); prog('처리하지 못했습니다: ' + ((e && e.message) || e)); }
}
async function handleProposal(file){
  if(!file) return;
  const pid = S.pid, oldProp = S.project && S.project.proposal;
  const prog = m => { setupMsg.prop = m; const e = $('#propProg'); if(e) e.textContent = m; };
  try{
    let text = '', pages = 0;
    if(/\.pdf$/i.test(file.name)){
      prog('PDF를 읽는 중'); await ensureLibs('pdf');
      const pdf = await pdfjsLib.getDocument({data: new Uint8Array(await file.arrayBuffer())}).promise;
      const t = await pdfPageTexts(pdf, (i,n) => prog(`글자 읽는 중 ${i}/${n}`)); pages = t.length;
      text = t.map((x,i) => `[제안서 p.${i+1}]\n${x}\n`).join('\n');
    } else if(/\.pptx$/i.test(file.name)){
      prog('PPTX를 읽는 중'); const sl = await parsePptx(file); pages = sl.length;
      text = sl.map((s,i) => `[제안서 p.${i+1}]\n${s.title}\n${s.body}\n${s.notes ? '(노트) ' + s.notes + '\n' : ''}`).join('\n');
    } else if(/\.docx$/i.test(file.name)){
      prog('DOCX를 읽는 중'); text = await parseDocx(file);
    } else { toast('PPTX, PDF, DOCX만 읽을 수 있습니다. HWP는 PDF로 저장해서 올려 주세요.', 5000); return; }
    if(nonSpace(text) < 200) toast('글자를 거의 찾지 못했습니다. 이미지로만 된 파일이라면 글자가 있는 PDF로 다시 저장해 주세요.', 6000);
    prog('파일 올리는 중');
    const up = await W(() => uploadFile(pid, file, 'proposal'));
    await W(() => q(sb.from('proposal_text').upsert({project_id: pid, body: text})));
    await W(() => q(sb.from('projects').update({proposal: {path: up.path, url: up.url, name: file.name, pages, chars: text.length, at: Date.now()}}).eq('id', pid)));
    if(oldProp) removeFiles([oldProp.path]);
    prog(`완료: ${pages ? pages + '쪽, ' : ''}${text.length.toLocaleString()}자 저장`); toast('정성제안서를 저장했습니다.');
    await loadProjects();
    if(S.slides.length){ try{ const n = await autoGenerate({silent: true}); if(n) prog(setupMsg.prop + ` · 예상 질의 ${n}개 다시 생성`); }catch(e){} }
  }catch(e){ console.error(e); prog('처리하지 못했습니다: ' + ((e && e.message) || e)); }
}
async function deleteProject(){
  const p = S.project; if(!p) return;
  if(!(await ask('사업을 삭제할까요?', `'${p.name}'의 발표자료·제안서 파일, 슬라이드 정보, 스크립트 ${Object.keys(S.scripts).length}장, 질의 ${S.qa.length}개를 모두 지웁니다. 되돌릴 수 없습니다.`, '삭제'))) return;
  try{
    const listed = await sb.storage.from('files').list(p.id, {limit: 1000});
    if(listed.data && listed.data.length) await removeFiles(listed.data.map(f => `${p.id}/${f.name}`));
    await W(() => q(sb.from('projects').delete().eq('id', p.id)), '삭제했습니다.');
    LS.set('pid', null); S.pid = null; await loadProjects();
  }catch(e){}
}

/* ───────── 이벤트 ───────── */
function bindUI(){
  $('#meName').addEventListener('input', e => { S.me = e.target.value; LS.set('me', S.me); });
  $('#projSel').addEventListener('change', e => { if(e.target.value) selectProject(e.target.value); });
  // 리허설
  $('#rhPresets').innerHTML = [10,15,20,30].map(m => `<button type="button" class="chip" data-min="${m}">${m}분</button>`).join('');
  $('#rhPresets').addEventListener('click', e => { const b = e.target.closest('[data-min]'); if(!b) return; $('#rhMin').value = b.dataset.min; $('#rhSec').value = 0; saveRhSettings(); tick(); });
  ['rhMin','rhSec','rhAlloc','rhWarn'].forEach(id => $('#'+id).addEventListener('input', () => { saveRhSettings(); tick(); }));
  $('#rhToggle').onclick = () => R.toggle();
  $('#rhPrev').onclick = () => R.go(R.idx - 1);
  $('#rhNext').onclick = () => R.go(R.idx + 1);
  $('#rhReset').onclick = () => { R.reset(); $('#rhResult').innerHTML = ''; renderRehearsal(); };
  $('#rhFinish').onclick = finishRehearsal;
  $('#rhScriptBtn').onclick = toggleScriptPane;
  $('#rhFull').onclick = toggleFull;
  let spSize = LS.get('spSize', 17); document.documentElement.style.setProperty('--sp-size', spSize + 'px');
  const setSp = d => { spSize = Math.max(13, Math.min(30, spSize + d)); document.documentElement.style.setProperty('--sp-size', spSize + 'px'); LS.set('spSize', spSize); };
  $('#spSmaller').onclick = () => setSp(-2); $('#spBigger').onclick = () => setSp(2);
  if(LS.get('noScript', false)) toggleScriptPane();
  document.addEventListener('fullscreenchange', () => { $('#rhFull').textContent = document.fullscreenElement ? '전체 화면 끝내기' : '전체 화면'; R.lastShown = null; setTimeout(renderRehearsal, 80); });
  let rz; window.addEventListener('resize', () => { clearTimeout(rz); rz = setTimeout(() => { R.lastShown = null; if(S.tab === 'rehearsal' && sb) renderRehearsal(); }, 200); });
  // 스크립트
  $('#scSearch').addEventListener('input', renderScripts); $('#scFilter').addEventListener('change', renderScripts);
  $('#scList').addEventListener('click', e => {
    const b = e.target.closest('[data-act]'); if(!b) return;
    const no = +b.dataset.no;
    if(b.dataset.act === 'edit'){ if(scEditing){ toast('먼저 편집 중인 스크립트를 저장하거나 취소하세요.'); return; } editScript(no); }
    else if(b.dataset.act === 'save') saveScript(no);
    else if(b.dataset.act === 'cancel'){ scEditing = null; renderScripts(); }
    else if(b.dataset.act === 'rh'){ R.reset(); R.idx = no - 1; location.hash = '#rehearsal'; }
  });
  // 질의 목록
  $('#qaAddBtn').onclick = () => { qaPanel = {type:'add'}; renderQAPanel(); };
  $('#qaPasteBtn').onclick = () => { qaPanel = {type:'paste'}; renderQAPanel(); };
  $('#qaAiBtn').onclick = () => { qaPanel = {type:'ai'}; aiItems = null; renderQAPanel(); };
  $('#qaAutoBtn').onclick = () => { qaPanel = {type:'auto'}; renderQAPanel(); };
  $('#prSrc').addEventListener('change', renderPractice);
  $('#qaCsvBtn').onclick = exportQA;
  ['qaSearch','qaCat','qaSrc','qaSort','qaShowA'].forEach(id => $('#'+id).addEventListener(id === 'qaSearch' ? 'input' : 'change', renderQA));
  $('#qaPanel').addEventListener('click', async e => {
    const b = e.target.closest('[data-act]'); if(!b) return;
    const act = b.dataset.act;
    if(act === 'closePanel'){ qaPanel = null; aiItems = null; renderQAPanel(); }
    else if(act === 'saveQ') saveQAForm();
    else if(act === 'doPaste'){
      const rows = parseQAInput($('#pasteBox').value); if(!rows.length) return;
      b.disabled = true; b.textContent = '추가하는 중';
      try{ const n = await addMany(rows, 'import'); toast(`질의 ${n}개를 추가했습니다.`); qaPanel = null; renderQA(); }catch(err){ b.disabled = false; b.textContent = '다시 시도'; }
    }
    else if(act === 'autoGo'){
      b.disabled = true; const prog = $('#autoProg'); prog.textContent = '만드는 중';
      try{ const n = await autoGenerate({max: +$('#autoMax').value, perSlide: +$('#autoPer').value, includeProposal: $('#autoProp').checked}); toast(`예상 질의 ${n}개를 만들었습니다.`); qaPanel = null; $('#qaSrc').value = 'auto'; renderQA(); }
      catch(err){ b.disabled = false; prog.textContent = '만들지 못했습니다.'; }
    }
    else if(act === 'aiCopy'){
      const prog = $('#aiProg'); prog.textContent = '프롬프트를 만드는 중';
      const text = await buildPrompt();
      try{ await navigator.clipboard.writeText(text); prog.textContent = `복사했습니다 (${text.length.toLocaleString()}자). claude.ai 새 대화에 붙여넣으세요.`; }
      catch(err){ prog.textContent = '자동 복사가 막혀 있습니다. 아래 칸의 내용을 직접 복사하세요.'; $('#aiPaste').value = text; $('#aiPaste').select(); }
    }
    else if(act === 'aiAll' || act === 'aiNone'){ aiItems.forEach(x => x.pick = act === 'aiAll'); renderAiItems(); }
    else if(act === 'aiAdd'){
      const pick = aiItems.filter(x => x.pick); if(!pick.length){ toast('넣을 질의를 고르세요.'); return; }
      b.disabled = true; b.textContent = '넣는 중';
      try{ const n = await addMany(pick, 'ai'); toast(`질의 ${n}개를 목록에 넣었습니다. 답변은 검토 후 다듬어 주세요.`, 4500); aiItems = null; qaPanel = null; renderQA(); }catch(err){ b.disabled = false; b.textContent = '다시 시도'; }
    }
  });
  $('#qaList').addEventListener('click', async e => {
    const b = e.target.closest('[data-act]'); if(!b) return;
    const id = b.dataset.id;
    if(b.dataset.act === 'toggleA'){ const a = document.getElementById('qaa-' + id); if(a){ a.classList.toggle('hidden'); b.textContent = a.classList.contains('hidden') ? '답변 보기' : '답변 가리기'; } }
    else if(b.dataset.act === 'editQ'){ qaPanel = {type:'edit', id}; renderQAPanel(); $('#qaPanel').scrollIntoView({behavior:'smooth', block:'start'}); }
    else if(b.dataset.act === 'delQ'){ const x = S.qa.find(y => y.id === id); if(x && await ask('질의를 삭제할까요?', x.q.slice(0,120), '삭제')){ try{ await W(() => q(sb.from('qa').delete().eq('id', id)), '삭제했습니다.'); await loadTable('qa'); }catch(err){} } }
  });
  // 질의 연습
  $('#prQ').addEventListener('input', renderPractice);
  $('#prClear').onclick = () => { $('#prQ').value = ''; prCat = ''; renderPractice(); $('#prQ').focus(); };
  $('#prMark').addEventListener('change', renderPractice);
  $('#prCats').addEventListener('click', e => { const b = e.target.closest('[data-cat]'); if(!b) return; prCat = b.dataset.cat; renderPractice(); });
  $('#prStart').onclick = () => startDrill(practiceSet());
  $('#prResults').addEventListener('click', e => {
    const b = e.target.closest('[data-act]'); if(!b) return;
    if(b.dataset.act === 'prToggle'){ const a = document.getElementById('pra-' + b.dataset.id); if(a){ a.classList.toggle('hidden'); b.textContent = a.classList.contains('hidden') ? '답변 보기' : '답변 가리기'; } }
    else if(b.dataset.act === 'prOne'){ const x = S.qa.find(y => y.id === b.dataset.id); if(x) startDrill([x]); }
  });
  $('#drReveal').onclick = () => { if(D.summary && D.i >= D.list.length){ const l = D.summary; D.summary = null; startDrill(l); } else revealDrill(); };
  $('#drOk').onclick = () => { if(D.i < D.list.length) markDrill('ok'); };
  $('#drAgain').onclick = () => { if(D.i < D.list.length) markDrill('again'); };
  $('#drNext').onclick = () => { if(D.i < D.list.length){ D.i++; showDrill(); } };
  $('#drPrev').onclick = () => { if(D.i > 0){ D.i--; showDrill(); } };
  $('#drExit').onclick = () => stopDrill();
  // 설정
  $('#setupBody').addEventListener('click', e => {
    const b = e.target.closest('[data-act]'); if(!b) return;
    const act = b.dataset.act;
    if(act === 'pin'){ if($('#pinIn').value === String(CONFIG.ADMIN_PIN)){ sessionStorage.setItem('rr:pin', String(CONFIG.ADMIN_PIN)); renderSetup(); } else toast('번호가 맞지 않습니다.'); }
    else if(act === 'newProj') newProject();
    else if(act === 'saveProj') saveProject();
    else if(act === 'saveMemo') updateProject({competitor_memo: $('#pMemo').value.trim()}, '메모를 저장했습니다.').catch(() => {});
    else if(act === 'delProj') deleteProject();
  });
  // 키보드
  document.addEventListener('keydown', e => {
    if(!$('#overlay').classList.contains('hidden')){ if(e.key === 'Escape') $('#ovNo').click(); return; }
    if(/INPUT|TEXTAREA|SELECT/.test(e.target.tagName || '') || e.target.isContentEditable) return;
    if(e.metaKey || e.ctrlKey || e.altKey) return;
    if(S.tab === 'rehearsal'){
      if(e.key === ' ' || e.code === 'Space'){ e.preventDefault(); R.toggle(); }
      else if(['ArrowRight','PageDown'].includes(e.key)){ e.preventDefault(); R.go(R.idx + 1); }
      else if(['ArrowLeft','PageUp'].includes(e.key)){ e.preventDefault(); R.go(R.idx - 1); }
      else if(['s','S','ㄴ'].includes(e.key)) toggleScriptPane();
      else if(['f','F','ㄹ'].includes(e.key)) toggleFull();
    } else if(S.tab === 'practice' && !$('#prDrill').classList.contains('hidden')){
      if(e.key === ' ' || e.code === 'Space'){ e.preventDefault(); $('#drReveal').click(); }
      else if(e.key === '1') $('#drOk').click();
      else if(e.key === '2') $('#drAgain').click();
      else if(e.key === 'ArrowRight') $('#drNext').click();
      else if(e.key === 'ArrowLeft') $('#drPrev').click();
      else if(e.key === 'Escape') stopDrill();
    }
  });
}
function toggleScriptPane(){
  const w = $('#stageWrap'); w.classList.toggle('noscript');
  const off = w.classList.contains('noscript');
  $('#rhScriptBtn').textContent = off ? '스크립트 보기' : '스크립트 가리기';
  LS.set('noScript', off);
}
function toggleFull(){
  const w = $('#stageWrap');
  try{
    if(document.fullscreenElement) document.exitFullscreen();
    else if(w.requestFullscreen) w.requestFullscreen().catch(() => toast('전체 화면을 쓸 수 없습니다. 브라우저의 전체 화면(F11)을 쓰세요.'));
  }catch(e){ toast('전체 화면을 쓸 수 없습니다.'); }
}

init();
