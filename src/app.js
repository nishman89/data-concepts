/* ---------- Small helpers ---------- */
const $ = (s, el=document) => el.querySelector(s);
const $$ = (s, el=document) => [...el.querySelectorAll(s)];
const h = (tag, attrs={}, ...kids) => { const e=document.createElement(tag); for (const [k,v] of Object.entries(attrs)) { if (k==='class') e.className=v; else if (k==='html') e.innerHTML=v; else if (k.startsWith('on')) e.addEventListener(k.slice(2), v); else if (v!==false && v!=null) e.setAttribute(k, v===true?'':v); } for (const k of kids.flat()) if (k!=null) e.append(k.nodeType?k:document.createTextNode(k)); return e; };
let toastT; const toast = msg => { const t=$('#toast'); t.textContent=msg; t.classList.add('on'); clearTimeout(toastT); toastT=setTimeout(()=>t.classList.remove('on'),2200); };
const SVGNS='http://www.w3.org/2000/svg';
const sv=(tag, attrs={}, ...kids)=>{ const e=document.createElementNS(SVGNS, tag); for (const [k,v] of Object.entries(attrs)) if (v!=null) e.setAttribute(k,v); for (const k of kids) if (k!=null) e.append(k.nodeType?k:document.createTextNode(k)); return e; };
// **bold** and `code` in exercise and quiz prompts
const fmtPrompt = t => { const out=[]; String(t).split(/(\*\*[^*]+\*\*|`[^`]+`)/).forEach(p=>{ if (p.startsWith('**')) out.push(h('strong',{},p.slice(2,-2))); else if (p.startsWith('`')) out.push(h('code',{},p.slice(1,-1))); else if (p) out.push(h('span',{html:p})); }); return out; };

/* ---------- Progress, saved in this browser only ---------- */
const store = {
  key:'dataConcepts.v1',
  read(){ try { return JSON.parse(localStorage.getItem(this.key)) || {}; } catch(e){ return {}; } },
  write(d){ try { localStorage.setItem(this.key, JSON.stringify(d)); } catch(e){} }
};
const P = Object.assign({ exDone:{}, drafts:{}, checks:{}, checkDone:{}, quizzes:{}, dataSheets:[], capstone:{}, mode:'here' }, store.read());
const save = () => store.write(P);

/* ---------- Header, footer and back-to-top (shared look with the ISTQB and SQL sites) ---------- */
$('#footer-year').textContent=String(new Date().getFullYear());
(()=>{ const b=$('#back-to-top'); const t=()=>{ b.hidden=window.scrollY<900; }; addEventListener('scroll',t,{passive:true}); t();
  b.addEventListener('click',()=>window.scrollTo({top:0,behavior:matchMedia('(prefers-reduced-motion: reduce)').matches?'auto':'smooth'})); })();
document.addEventListener('click',e=>{ const m=$('#mobile-menu'); if (m.open && (!m.contains(e.target) || e.target.closest('a'))) m.open=false; });

const WB = makeWorkbook(WORKBOOK);
const SHEET_NAMES = WORKBOOK.sheets.map(s=>s.name);

/* ---------- Expected answers (worked out from each model formula, once) ---------- */
const answerCol = sh => sh.cols.length + 1;
const CHECK_ROWS = 300;   // fill-down answers are checked on this many rows
const expCache = new Map();
function expected(task){
  if (expCache.has(task)) return expCache.get(task);
  const sh=WB.sheet(task.sheet), col=answerCol(sh), t={sheet:sh,row:2,col};
  let v;
  if (task.fill){ v=[]; const last=Math.min(lastRow(sh), CHECK_ROWS+1); for (let r=2;r<=last;r++) v.push(runFormula(WB, task.sol, t, {dr:r-2}).value); }
  else v=runFormula(WB, task.sol, t).value;
  expCache.set(task, v); return v;
}
const askRow = ex => ex.fill ? (ex.askRow || 2) : null;
const expectedAt = ex => ex.fill ? expected(ex)[askRow(ex)-2] : expected(ex);
function fmtVal(v, vtype){ if (typeof v==='number' && !vtype && !Number.isInteger(v)) return String(Math.round(v*10000)/10000); return displayValue(v, vtype); }

/* Marks a formula typed in the practice grid */
function markFormula(task, formula){
  const f=String(formula||'').trim();
  if (!f) return {ok:false, msg:'Type a formula first.'};
  if (!f.startsWith('=')) return {ok:false, msg:'Start your formula with an equals sign (=).', tip:'Worked it out in Excel instead? Switch to In Excel and type your answer there.'};
  const sh=WB.sheet(task.sheet), t={sheet:sh,row:2,col:answerCol(sh)};
  if (!task.fill){
    const r=runFormula(WB, f, t); if (r.error) return {ok:false, msg:`${r.error.code}: ${r.error.why}`};
    const want=expected(task);
    if (sameValue(r.value, want)) return {ok:true};
    if (isErr(r.value)) return {ok:false, msg:`Your formula gives ${r.value.code}. ${r.value.why||''}`};
    return {ok:false, msg:`Your formula gives ${fmtVal(r.value, task.vtype)}, which isn't the answer.`, tip:'Check the column, the criteria, and whether you need to leave out the heading.'};
  }
  const want=expected(task);
  for (let i=0;i<want.length;i++){
    const r=runFormula(WB, f, t, {dr:i});
    if (r.error) return {ok:false, msg:`${r.error.code}: ${r.error.why}`};
    if (!sameValue(r.value, want[i])){
      const got=isErr(r.value)?`${r.value.code}${r.value.why?' ('+r.value.why+')':''}`:fmtVal(r.value, task.vtype);
      if (i===0) return {ok:false, msg:`In row 2 your formula gives ${got}, but the answer is ${fmtVal(want[0], task.vtype)}.`};
      return {ok:false, msg:`Row 2 is right, but row ${i+2} gives ${got} instead of ${fmtVal(want[i], task.vtype)}.`, tip:'When a formula works in row 2 but not lower down, a reference has usually slid as it filled. Lock it with $ signs (F4).'};
    }
  }
  return {ok:true};
}
/* Marks a value typed in the "In Excel" box */
function markTyped(ex, text){
  const want=expectedAt(ex);
  const vt = ex.vtype || (typeof want==='boolean'?'bool':null);
  const t=readTyped(text, vt);
  if (t.value==null) return {ok:false, msg:'Type your answer first.'};
  const s=sameValue(t.value, want, {typed:true, dp:t.dp});
  if (s) return {ok:true, rounded:s==='rounded'};
  if (typeof want==='number' && typeof t.value==='number'){
    if (want!==0 && Math.abs(t.value/want-100)<0.01) return {ok:false, msg:'That looks like a percentage. Type it as a decimal, or add the % sign.'};
    return {ok:false, msg:`${text} isn't right.`, tip: Number.isInteger(want) ? 'The answer is a whole number.' : 'Give at least 1 decimal place, or more if you can.'};
  }
  return {ok:false, msg:`“${text}” isn't right.`, tip: typeof want==='string' ? 'Type the text as it appears in the data.' : ''};
}

/* ---------- The practice grid: an Excel-like sheet with a formula bar ----------
   makeSheetView(host, { sheet, col, fill, vtype, editable, onRun, onView, testid })
   The formula always goes in row 2 of a new pink column to the right of the data (the "answer column"). */
const COLOR_N = 4;
function makeSheetView(host, opts={}){
  const st={ view:opts.sheet||'Members', target:opts.sheet||'Members', col:opts.col||'Answer', fill:!!opts.fill, vtype:opts.vtype||null,
    formula:'', results:null, error:null, show:200, lastIns:null, sel:null };
  const editable = opts.editable!==false;
  const tp = opts.testid || 'sheet';
  host.innerHTML='';
  const name=h('span',{class:'xl-name','data-testid':`${tp}-name-box`});
  const input=h('input',{class:'xl-input',type:'text',spellcheck:'false',autocomplete:'off','aria-label':'Formula','data-testid':`${tp}-formula`,placeholder:editable?'Type a formula, such as =SUM(E:E)':''});
  const runBtn=h('button',{class:'btn small primary',type:'button','data-testid':`${tp}-run`},'Enter');
  const fillLbl=h('label',{class:'xl-fill',title:'Copy the formula down every row, like double-clicking the fill handle in Excel'}, h('input',{type:'checkbox','data-testid':`${tp}-fill`}), ' Fill down');
  const fillBox=$('input',fillLbl);
  const bar=h('div',{class:'xl-bar'}, name, h('span',{class:'xl-fx','aria-hidden':'true'},'fx'), input, editable?runBtn:null);
  const msg=h('div',{class:'xl-msg','aria-live':'polite','data-testid':`${tp}-result`});
  const scroll=h('div',{class:'xl-scroll',tabindex:'0','aria-label':'Sheet'});
  const tabs=h('div',{class:'xl-tabs',role:'tablist','aria-label':'Sheets'});
  const foot=h('div',{class:'xl-foot'});
  const wrap=h('div',{class:'xl'+(editable?'':' ro'),'data-testid':tp}, editable?bar:null, editable?msg:null, scroll, tabs, foot);
  host.append(wrap);
  if (!editable) input.disabled=true;

  const T=()=>WB.sheet(st.target), V=()=>WB.sheet(st.view);
  const addr=()=>`${st.target}!${colLetter(answerCol(T()))}2`;
  const quoteSheet=n=>/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n)?n:`'${n.replace(/'/g,"''")}'`;
  const prefix=()=>st.view===st.target?'':quoteSheet(st.view)+'!';

  function compute(){
    st.results=null; st.error=null;
    const f=input.value.trim(); st.formula=f;
    if (!f) return;
    if (!f.startsWith('=')){ st.results={0:{value:readTyped(f).value}}; return; }
    const sh=T(), t={sheet:sh,row:2,col:answerCol(sh)};
    if (!st.fill){ const r=runFormula(WB, f, t); if (r.error) st.error=r.error; else st.results={0:r}; return; }
    st.results={}; const last=Math.min(lastRow(sh), st.show+1);
    for (let r=2;r<=last;r++){ const x=runFormula(WB, f, t, {dr:r-2}); if (x.error){ st.error=x.error; st.results=null; return; } st.results[r-2]=x; }
  }
  function renderMsg(){
    if (!editable) return;
    msg.innerHTML=''; msg.className='xl-msg';
    if (st.error){ msg.classList.add('err'); msg.append(h('b',{},st.error.code+' '), st.error.why||''); return; }
    if (!st.results){ msg.append(h('span',{class:'muted'}, st.fill?'Your formula is filled down the pink column.':'Your answer appears in the pink cell.')); return; }
    const v=st.results[0].value;
    if (isErr(v)){ msg.classList.add('err'); msg.append(h('b',{},v.code+' '), v.why||''); return; }
    msg.append(st.fill?'Row 2: ':'Result: ', h('b',{'data-testid':`${tp}-value`}, fmtVal(v, st.vtype)));
    if (typeof v==='number' && !st.vtype && v>20000 && v<60000 && Number.isInteger(Math.round(v))) msg.append(h('span',{class:'muted'},`  (as a date: ${displayValue(v,'date')})`));
  }
  function cell(v, type){ const d=displayValue(v, type); const cls = typeof v==='number'?'n':typeof v==='boolean'?'b':''; return [d, cls]; }
  function renderGrid(){
    const sh=V(), isT=st.view===st.target, nC=sh.cols.length, ac=answerCol(sh), last=Math.min(lastRow(sh), st.show+1);
    const showAns=isT && editable;
    let html='<table class="xl-grid"><thead><tr><th class="xl-corner" aria-hidden="true"></th>';
    for (let c=1;c<=nC;c++) html+=`<th class="xl-col" data-c="${c}" scope="col">${colLetter(c)}</th>`;
    if (showAns) html+=`<th class="xl-col xl-ans" data-c="${ac}" scope="col">${colLetter(ac)}</th>`;
    html+='</tr></thead><tbody>';
    for (let r=1;r<=last;r++){
      html+=`<tr${r===1?' class="xl-h"':''}><th class="xl-row" data-r="${r}" scope="row">${r}</th>`;
      for (let c=1;c<=nC;c++){ const v=cellValue(sh,r,c); const [d,cls]=r===1?[esc(String(v)),'']:cell(v, sh.cols[c-1].type); html+=`<td data-r="${r}" data-c="${c}"${cls?` class="${cls}"`:''}>${r===1?d:esc(d)}</td>`; }
      if (showAns){
        let d='', cls='xl-ans';
        if (r===1) d=esc(st.col);
        else if (st.results && (st.fill || r===2) && st.results[r-2]){ const v=st.results[r-2].value; d=esc(isErr(v)?v.code:fmtVal(v, st.vtype)); if (typeof v==='number') cls+=' n'; if (isErr(v)) cls+=' e'; }
        if (r===2) cls+=' xl-target';
        html+=`<td data-r="${r}" data-c="${ac}" class="${cls}">${d}</td>`;
      }
      html+='</tr>';
    }
    html+='</tbody></table>';
    scroll.innerHTML=html;
    foot.innerHTML='';
    const total=lastRow(sh);
    foot.append(h('span',{}, total>last ? `Showing rows 1 to ${last} of ${total.toLocaleString('en-GB')}. Formulas use every row.` : `${(total-1).toLocaleString('en-GB')} rows of data.`));
    if (total>last) foot.append(h('button',{class:'linklike',type:'button',onclick:()=>{ st.show+=300; compute(); renderGrid(); }},'Show more rows'));
    highlight();
  }
  function renderTabs(){
    tabs.innerHTML='';
    for (const n of SHEET_NAMES){
      const b=h('button',{type:'button',role:'tab','aria-selected':String(n===st.view),class:(n===st.target&&editable?'tgt ':''),'data-testid':`${tp}-tab-${n}`}, n);
      b.addEventListener('mousedown',e=>{ if (document.activeElement===input) e.preventDefault(); });
      b.addEventListener('click',()=>{ st.view=n; st.lastIns=null; renderTabs(); renderGrid(); opts.onView&&opts.onView(n); });
      tabs.append(b);
    }
  }
  function highlight(){
    $$('.hl0,.hl1,.hl2,.hl3',scroll).forEach(e=>e.classList.remove('hl0','hl1','hl2','hl3'));
    const refs=refsIn(input.value); let k=0;
    for (const r of refs){
      const sheet = r.sheet ? r.sheet : st.target;
      if (sheet.toLowerCase()!==st.view.toLowerCase()){ k++; continue; }
      const cls='hl'+(k++%COLOR_N);
      const c1=r.t==='ref'?r.c:r.c1, c2=r.t==='ref'?r.c:r.c2, r1=r.t==='ref'?r.r:(r.whole?1:r.r1), r2=r.t==='ref'?r.r:(r.whole?1e9:r.r2);
      for (let c=Math.min(c1,c2);c<=Math.max(c1,c2);c++){
        const th=$(`th.xl-col[data-c="${c}"]`,scroll); th&&th.classList.add(cls);
        $$(`td[data-c="${c}"]`,scroll).forEach(td=>{ const rr=+td.dataset.r; if (rr>=Math.min(r1,r2)&&rr<=Math.max(r1,r2)) td.classList.add(cls); });
      }
    }
  }
  function run(){ compute(); renderMsg(); renderGrid(); if (st.view===st.target){ const th=$('th.xl-ans',scroll); if (th && th.offsetLeft+th.offsetWidth>scroll.scrollLeft+scroll.clientWidth) scroll.scrollLeft=th.offsetLeft+th.offsetWidth-scroll.clientWidth+8; } if (opts.onRun) opts.onRun(input.value, st); }
  // Select a cell or column while typing a formula to add its reference, as in Excel. Shift extends it to a range.
  function insertRef(ref, shiftAnchor){
    let s=input.value, a=input.selectionStart, b=input.selectionEnd;
    if (!s.trim()){ s='='; a=b=1; }
    const L=st.lastIns;
    if (L && a===b && a===L.end && s.slice(L.start,L.end)===L.text){
      if (shiftAnchor && L.anchor){ ref = (L.pre||'')+L.anchor+':'+ref.replace(/^.*!/,''); }
      a=L.start; b=L.end;
    }
    const text=ref;
    input.value=s.slice(0,a)+text+s.slice(b);
    const pos=a+text.length; input.setSelectionRange(pos,pos); input.focus();
    st.lastIns={start:a,end:pos,text,anchor:(shiftAnchor&&L&&L.anchor)?L.anchor:ref.replace(/^.*!/,''),pre:prefix()};
    highlight();
  }
  scroll.addEventListener('mousedown',e=>{
    const td=e.target.closest('td[data-c],th.xl-col'); if (!td) return;
    const c=+td.dataset.c, r=td.dataset.r?+td.dataset.r:null;
    if (editable && document.activeElement===input && /^=/.test(input.value.trim()||'=') && !(st.view===st.target && c===answerCol(V()))){
      e.preventDefault();
      const ref = r ? prefix()+colLetter(c)+r : prefix()+colLetter(c)+':'+colLetter(c);
      insertRef(ref, e.shiftKey);
      return;
    }
    if (r){ const v=cellValue(V(),r,c); name.textContent=`${st.view}!${colLetter(c)}${r}`; if (!editable) return;
      msg.className='xl-msg'; msg.innerHTML=''; msg.append(h('span',{class:'muted'},`${st.view}!${colLetter(c)}${r}: `), h('b',{}, r===1?String(v): c>V().cols.length ? (st.results&&st.results[r-2]?fmtVal(st.results[r-2].value):'') : displayValue(v, V().cols[c-1].type)||'(empty)'));
      setTimeout(()=>{ name.textContent=addr(); },2500); }
  });
  input.addEventListener('keydown',e=>{
    if (e.key==='Enter'){ e.preventDefault(); st.lastIns=null; run(); }
    else if (e.key==='F4'){ e.preventDefault(); toggleAbs(); }
    else if (e.key==='Escape'){ st.lastIns=null; }
  });
  input.addEventListener('input',()=>{ st.lastIns=null; highlight(); if (opts.onChange) opts.onChange(input.value); });
  runBtn.addEventListener('click',()=>{ st.lastIns=null; run(); });
  fillBox.addEventListener('change',()=>{ st.fill=fillBox.checked; run(); });
  // F4 cycles A1 → $A$1 → A$1 → $A1 → A1 on the reference at the cursor
  function toggleAbs(){
    const s=input.value, pos=input.selectionStart, re=/(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\d(])|(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9(])/g; let m;
    while ((m=re.exec(s))){
      const a=m.index, b=a+m[0].length; if (pos<a||pos>b) continue;
      if (s[a-1] && /[A-Za-z0-9_."]/.test(s[a-1])) continue;
      let t;
      if (m[2]){ const st2=(m[1]?2:0)+(m[3]?1:0); const nx={0:3,3:1,1:2,2:0}[st2]; t=(nx&2?'$':'')+m[2]+(nx&1?'$':'')+m[4]; }
      else { const on=!m[5]; t=(on?'$':'')+m[6]+':'+(on?'$':'')+m[8]; }
      input.value=s.slice(0,a)+t+s.slice(b); input.setSelectionRange(a+t.length,a+t.length); highlight(); return;
    }
  }
  const api={
    set(o={}){ if (o.sheet){ st.target=o.sheet; st.view=o.sheet; } if (o.col!=null) st.col=o.col; if (o.fill!=null){ st.fill=!!o.fill; fillBox.checked=st.fill; } if ('vtype' in o) st.vtype=o.vtype;
      if (o.formula!=null) input.value=o.formula; name.textContent=addr(); st.show=200; st.lastIns=null; if (o.run && input.value.trim()) run(); else { compute(); renderMsg(); renderGrid(); } renderTabs(); },
    get value(){ return input.value; }, set value(v){ input.value=v; },
    focus(){ input.focus(); }, run, input, state:st, fillLabel:fillLbl,
    showFill(on){ if (on) bar.insertBefore(fillLbl, runBtn); else fillLbl.remove(); },
  };
  name.textContent=addr(); renderTabs(); renderMsg(); renderGrid();
  return api;
}

/* ---------- Entity relationship diagrams (the same look as the SQL Practice Lab) ---------- */
const ROW=21, HEAD=32, CW=7.25;
const badgeW = t => t.cols.some(c=>c.pk&&c.fk) ? 58 : 30;
function measure(t){ const nameW=Math.max(...t.cols.map(c=>c.name.length),4)*CW; return {w:Math.max(150, Math.ceil(12+badgeW(t)+nameW+16), Math.ceil(t.name.length*8.6+28)), h:HEAD+t.cols.length*ROW+4}; }
function drawEnd(g, x, y, d, kind){
  if (kind==='many0'){ g.append(sv('line',{x1:x+d*13,y1:y,x2:x,y2:y-7}), sv('line',{x1:x+d*13,y1:y,x2:x,y2:y+7}), sv('line',{x1:x+d*13,y1:y,x2:x,y2:y}), sv('circle',{cx:x+d*20,cy:y,r:4})); }
  else { g.append(sv('line',{x1:x+d*8,y1:y-7,x2:x+d*8,y2:y+7}), sv('line',{x1:x+d*14,y1:y-7,x2:x+d*14,y2:y+7})); }
}
function renderERD(host, model, {testPrefix='erd', label='Entity relationship diagram'}={}){
  host.innerHTML='';
  const tables=model.tables.map(t=>({name:t.name, x:t.x, y:t.y, fact:t.fact, cols:t.cols.map(([n,k,ref])=>({name:n, pk:k==='PK', fk:k==='FK', ref}))}));
  const box={}; let maxX=0, maxY=0;
  for (const t of tables){ const m=measure(t); box[t.name]={x:t.x,y:t.y,...m}; maxX=Math.max(maxX,t.x+m.w); maxY=Math.max(maxY,t.y+m.h); }
  const pad=30;
  const svg=sv('svg',{class:'erd',role:'img','aria-label':label, viewBox:`${-pad} ${-pad} ${maxX+pad*2} ${maxY+pad*2}`, width:'100%'});
  svg.style.minWidth=Math.round((maxX+pad*2)*0.72)+'px'; svg.style.maxWidth=Math.round(maxX+pad*2)+'px';
  const relLayer=sv('g'), tblLayer=sv('g'); svg.append(relLayer, tblLayer);
  const rowY=(t,i)=>box[t.name].y+HEAD+i*ROW+ROW/2+2;
  const rels=[];
  for (const t of tables) t.cols.forEach((c,i)=>{
    if (!c.fk||!c.ref) return; const p=tables.find(x=>x.name===c.ref); if (!p) return;
    const pk=p.cols.findIndex(x=>x.pk); const A=box[t.name], B=box[p.name];
    const sy=rowY(t,i), ey=rowY(p,Math.max(0,pk)); let sx,ex,d1,d2,path;
    if (A.x+A.w+30<=B.x){ sx=A.x+A.w; d1=1; ex=B.x; d2=-1; const o=Math.max(30,(ex-sx)/2); path=`M${sx},${sy} C${sx+o},${sy} ${ex-o},${ey} ${ex},${ey}`; }
    else if (B.x+B.w+30<=A.x){ sx=A.x; d1=-1; ex=B.x+B.w; d2=1; const o=Math.max(30,(sx-ex)/2); path=`M${sx},${sy} C${sx-o},${sy} ${ex+o},${ey} ${ex},${ey}`; }
    else { sx=A.x+A.w; d1=1; ex=B.x+B.w; d2=1; const rx=Math.max(sx,ex)+40; path=`M${sx},${sy} C${rx},${sy} ${rx},${ey} ${ex},${ey}`; }
    const g=sv('g',{class:'rel','data-a':t.name,'data-b':p.name}); g.append(sv('title',{},`${t.name}.${c.name} links to ${p.name}`), sv('path',{d:path}));
    drawEnd(g,sx,sy,d1,'many0'); drawEnd(g,ex,ey,d2,'one1'); relLayer.append(g); rels.push(g);
  });
  for (const t of tables){
    const b=box[t.name];
    const g=sv('g',{class:'tbl'+(t.fact?' fact':''),'data-t':t.name,'data-testid':`${testPrefix}-table-${t.name.toLowerCase()}`,transform:`translate(${b.x},${b.y})`,tabindex:'0','aria-label':`${t.name}: ${t.cols.map(c=>c.name).join(', ')}`});
    g.append(sv('rect',{class:'tbl-box',x:0,y:0,width:b.w,height:b.h,rx:5}), sv('path',{class:'tbl-head',d:`M0,5 a5,5 0 0 1 5,-5 h${b.w-10} a5,5 0 0 1 5,5 v${HEAD-5} h${-b.w}z`}), sv('text',{class:'tbl-name',x:12,y:21},t.name));
    t.cols.forEach((c,i)=>{ const y=HEAD+i*ROW+2, cy=y+ROW/2; if (i>0) g.append(sv('line',{class:'rowline',x1:8,x2:b.w-8,y1:y,y2:y})); let bx=10;
      if (c.pk){ g.append(sv('rect',{class:'k-pk',x:bx,y:cy-7,width:24,height:14,rx:3}), sv('text',{class:'k-pk-t',x:bx+12,y:cy+3.3,'text-anchor':'middle'},'PK')); bx+=28; }
      if (c.fk) g.append(sv('rect',{class:'k-fk',x:bx+.5,y:cy-6.5,width:23,height:13,rx:3}), sv('text',{class:'k-fk-t',x:bx+12,y:cy+3.3,'text-anchor':'middle'},'FK'));
      g.append(sv('text',{class:'col',x:12+badgeW(t),y:cy+4},c.name)); });
    tblLayer.append(g);
    const focus=on=>{ svg.classList.toggle('focusing',on); if (!on){ $$('.hot',svg).forEach(x=>x.classList.remove('hot')); return; } g.classList.add('hot');
      for (const r of rels) if (r.dataset.a===t.name||r.dataset.b===t.name){ r.classList.add('hot'); const o=r.dataset.a===t.name?r.dataset.b:r.dataset.a; const og=tblLayer.querySelector(`[data-t="${o}"]`); og&&og.classList.add('hot'); } };
    g.addEventListener('mouseenter',()=>focus(true)); g.addEventListener('mouseleave',()=>focus(false)); g.addEventListener('focus',()=>focus(true)); g.addEventListener('blur',()=>focus(false));
  }
  host.append(svg); return svg;
}

/* ---------- Interactive pieces in lessons (W('name') in the content) ---------- */
const colNums = (sheet, name) => { const sh=WB.sheet(sheet); const i=sh.cols.findIndex(c=>c.name===name); return sh.rows.map(r=>r[i]).filter(v=>typeof v==='number'); };
const meanOf = a => a.reduce((s,x)=>s+x,0)/a.length;
const medianOf = a => { const s=[...a].sort((x,y)=>x-y), m=s.length/2; return s.length%2?s[Math.floor(m)]:(s[m-1]+s[m])/2; };
const sdOf = a => { const m=meanOf(a); return Math.sqrt(a.reduce((s,x)=>s+(x-m)**2,0)/a.length); };
const r2 = x => Math.round(x*100)/100;
function chartFrame(W,H,m){ return {W,H,m, x0:m.l, x1:W-m.r, y0:H-m.b, y1:m.t}; }
function niceTicks(lo, hi, n=5){ const span=hi-lo||1, step0=span/n, mag=10**Math.floor(Math.log10(step0)), step=[1,2,2.5,5,10].map(k=>k*mag).find(s=>s>=step0); const out=[]; for (let v=Math.ceil(lo/step)*step; v<=hi+1e-9; v+=step) out.push(Math.round(v*1e6)/1e6); return out; }
const WIDGETS = {
  sheet(el){ const v=makeSheetView(el,{sheet:el.dataset.sheet, editable:false, testid:'widget-sheet'}); v.set({sheet:el.dataset.sheet}); },
  typerecord(el){
    const typeOf = (sheet, col) => DATA_TYPES.find(t=>t.cols.some(([s2,c2])=>s2===sheet&&c2===col));
    const SHEETS=['Members','Visits','Bookings','Sites','Classes','Plans','SiteMonths'];
    let sheet='Members', row=0;
    const tabs=h('div',{class:'dtabs',role:'group','aria-label':'Sheet'});
    SHEETS.forEach(n=>tabs.append(h('button',{type:'button','aria-pressed':String(n===sheet),'data-testid':`typerecord-${n}`,onclick:()=>{ sheet=n; row=0; $$('button',tabs).forEach(b=>b.setAttribute('aria-pressed',String(b.textContent===n))); draw(); }},n)));
    const box=h('div',{class:'trec','data-testid':'typerecord'});
    const more=h('button',{class:'btn small',type:'button','data-testid':'typerecord-another',onclick:()=>{ const sh=WB.sheet(sheet); row=(row+1+Math.floor(Math.random()*Math.min(sh.rows.length-1,400)))%sh.rows.length; draw(); }},'Show another row');
    el.append(tabs, box, h('p',{}, more));
    const draw=()=>{
      const sh=WB.sheet(sheet), r=sh.rows[row]; box.innerHTML='';
      box.append(h('div',{class:'trec-h'}, `${sheet}, row ${row+2}`));
      sh.cols.forEach((c,i)=>{ const t=typeOf(sheet,c.name), v=r[i];
        box.append(h('div',{class:'trec-r'}, h('code',{class:'trec-c'},c.name), h('span',{class:'trec-v'+(v==null||v===''?' blank':'')}, v==null||v===''?'(blank)':displayValue(v,c.type)), t?h('span',{class:'tchip t-'+t.key},t.name):h('span')));
      });
    };
    draw();
  },
  typecards(el){
    const grid=h('div',{class:'tcards'});
    const samples=(sheet,col,n)=>{ const sh=WB.sheet(sheet), i=sh.cols.findIndex(c=>c.name===col), out=[], seen=new Set();
      for (const r of sh.rows){ const v=r[i]; if (v==null||v==='') continue; const d=displayValue(v, sh.cols[i].type); if (seen.has(d)) continue; seen.add(d); out.push(d); if (out.length>=n) break; } return out; };
    for (const t of DATA_TYPES){
      const list=h('ul',{class:'tcard-cols'});
      const show=(sheet,col,li)=>{ $$('li',list).forEach(x=>x.classList.toggle('on',x===li)); vals.innerHTML=''; vals.append(h('span',{class:'small muted'},`${sheet}.${col}: `), ...samples(sheet,col,6).map(v=>h('code',{},v))); };
      const vals=h('div',{class:'tcard-vals','aria-live':'polite'});
      t.cols.forEach(([sheet,col],k)=>{ const li=h('li',{}, h('button',{type:'button',class:'linklike','data-testid':`typecard-${t.key}-${col}`,onclick:()=>show(sheet,col,li)}, col), h('span',{class:'small muted'},` ${sheet}`)); list.append(li); if (k===0) setTimeout(()=>show(sheet,col,li)); });
      grid.append(h('article',{class:'tcard t-'+t.key,'data-testid':`typecard-${t.key}`}, h('h3',{},t.name), h('p',{class:'small'},t.blurb), list, vals));
    }
    el.append(grid);
  },
  formats(el){
    const ms=WB.sheet('Members').rows.slice(0,3).map(r=>({id:r[0], first:r[1], last:r[2], plan:r[7], joined:fmtDate(r[8],'yyyy-mm-dd')}));
    const txt=ms.map(m=>`${m.first} ${m.last} joined on ${m.joined} on the ${m.plan} plan (member ${m.id}).`).join('\n');
    const csv='MemberID,FirstName,LastName,PlanCode,JoinDate\n'+ms.map(m=>[m.id,m.first,m.last,m.plan,m.joined].join(',')).join('\n');
    const json='[\n'+ms.map(m=>`  {"memberId": "${m.id}", "firstName": "${m.first}", "lastName": "${m.last}", "plan": "${m.plan}", "joined": "${m.joined}"}`).join(',\n')+'\n]';
    const yml='members:\n'+ms.map(m=>`  - member_id: ${m.id}\n    first_name: ${m.first}\n    last_name: ${m.last}\n    plan: ${m.plan}\n    joined: ${m.joined}`).join('\n');
    const xml='<?xml version="1.0" encoding="UTF-8"?>\n<members>\n'+ms.map(m=>`  <member id="${m.id}">\n    <firstName>${m.first}</firstName>\n    <lastName>${m.last}</lastName>\n    <plan>${m.plan}</plan>\n    <joined>${m.joined}</joined>\n  </member>`).join('\n')+'\n</members>';
    el.innerHTML=TABS([['.txt',txt],['.csv',csv],['.json',json],['.yml',yml],['.xml',xml]].map(([label,code])=>({label, html:`<pre class="code">${esc(code)}</pre>`})));
    wireTabs(el);
  },
  histogram(el){
    const ages = (()=>{ const sh=WB.sheet('Members'), end=serialOf(2025,12,31); return sh.rows.map(r=>{ const b=fromSerial(r[4]), e=fromSerial(end); return e.y-b.y-((e.m<b.m||(e.m===b.m&&e.d<b.d))?1:0); }); })();
    const SRC={ visits:{label:'Visit length (minutes)', data:colNums('Visits','DurationMins'), trim:true}, ages:{label:'Member age on 31 Dec 2025 (years)', data:ages}, joins:{label:'New members per site per month', data:colNums('SiteMonths','NewMembers')} };
    let src='visits', bins=25, trim=false;
    const sel=h('select',{'aria-label':'Data','data-testid':'hist-data'}, Object.entries(SRC).map(([k,s])=>h('option',{value:k},s.label)));
    const rng=h('input',{type:'range',min:'4',max:'60',value:String(bins),'aria-label':'Number of bins','data-testid':'hist-bins'});
    const rl=h('span',{class:'small'}); const tr=h('label',{class:'small'}, h('input',{type:'checkbox','data-testid':'hist-trim'}), ' Leave out visits over 300 minutes');
    const fig=h('div',{class:'w-fig'}), stats=h('p',{class:'w-stats','data-testid':'hist-stats'});
    el.append(h('div',{class:'w-controls'}, h('label',{class:'small'},'Data ',sel), h('label',{class:'small'},'Bins ',rng,rl), tr), fig, stats);
    const draw=()=>{
      let d=SRC[src].data; if (src==='visits'&&trim) d=d.filter(x=>x<=300); tr.hidden=src!=='visits';
      const lo=Math.min(...d), hi=Math.max(...d), w=(hi-lo)/bins||1; const counts=new Array(bins).fill(0); d.forEach(x=>counts[Math.min(bins-1,Math.floor((x-lo)/w))]++);
      const F=chartFrame(600,260,{l:46,r:12,t:12,b:40}), maxC=Math.max(...counts), X=v=>F.x0+(v-lo)/(hi-lo||1)*(F.x1-F.x0), Y=c=>F.y0-(c/maxC)*(F.y0-F.y1);
      const svg=sv('svg',{viewBox:`0 0 ${F.W} ${F.H}`,class:'chart',role:'img','aria-label':`Histogram of ${SRC[src].label} with ${bins} bins`});
      niceTicks(0,maxC,4).forEach(t=>svg.append(sv('line',{x1:F.x0,x2:F.x1,y1:Y(t),y2:Y(t),class:'grid'}), sv('text',{x:F.x0-6,y:Y(t)+4,'text-anchor':'end',class:'ax'},String(t))));
      counts.forEach((c,i)=>svg.append(sv('rect',{x:X(lo+i*w)+0.5,y:Y(c),width:Math.max(0.5,X(lo+(i+1)*w)-X(lo+i*w)-1),height:F.y0-Y(c),class:'bar'}, sv('title',{},`${r2(lo+i*w)} to ${r2(lo+(i+1)*w)}: ${c}`))));
      niceTicks(lo,hi,6).forEach(t=>svg.append(sv('text',{x:X(t),y:F.y0+16,'text-anchor':'middle',class:'ax'},String(t))));
      svg.append(sv('text',{x:(F.x0+F.x1)/2,y:F.H-4,'text-anchor':'middle',class:'ax-t'},SRC[src].label));
      const mn=meanOf(d), md=medianOf(d);
      svg.append(sv('line',{x1:X(mn),x2:X(mn),y1:F.y1,y2:F.y0,class:'mark-mean'}), sv('text',{x:X(mn)+4,y:F.y1+12,class:'mark-t mean'},'mean'));
      svg.append(sv('line',{x1:X(md),x2:X(md),y1:F.y1+18,y2:F.y0,class:'mark-med'}), sv('text',{x:X(md)-4,y:F.y1+30,'text-anchor':'end',class:'mark-t'},'median'));
      fig.innerHTML=''; fig.append(svg); rl.textContent=String(bins);
      const s=sdOf(d), sk=d.reduce((a,x)=>a+((x-mn)/s)**3,0)/d.length;
      stats.innerHTML=`<b>${d.length.toLocaleString('en-GB')}</b> values. Mean <b>${r2(mn)}</b>, median <b>${r2(md)}</b>, standard deviation <b>${r2(s)}</b>, skew <b>${r2(sk)}</b>.`;
    };
    sel.onchange=()=>{ src=sel.value; draw(); }; rng.oninput=()=>{ bins=+rng.value; draw(); }; $('input',tr).onchange=e=>{ trim=e.target.checked; draw(); };
    draw();
  },
  scatter(el){
    const sh=WB.sheet('SiteMonths'), col=n=>sh.cols.findIndex(c=>c.name===n);
    const XS={MarketingSpend:'Marketing spend (£)', AvgTempC:'Average temperature (°C)'}, YS={NewMembers:'New members', Visits:'Visits'};
    let xk='MarketingSpend', yk='NewMembers';
    const xs=h('select',{'aria-label':'x axis','data-testid':'scatter-x'},Object.entries(XS).map(([k,l])=>h('option',{value:k},l)));
    const ys=h('select',{'aria-label':'y axis','data-testid':'scatter-y'},Object.entries(YS).map(([k,l])=>h('option',{value:k},l)));
    const fig=h('div',{class:'w-fig'}), stats=h('p',{class:'w-stats','data-testid':'scatter-stats'});
    el.append(h('div',{class:'w-controls'}, h('label',{class:'small'},'Across ',xs), h('label',{class:'small'},'Up ',ys)), fig, stats);
    const draw=()=>{
      const pts=sh.rows.map(r=>[r[col(xk)], r[col(yk)], r[0], r[1]]);
      const X=pts.map(p=>p[0]), Y=pts.map(p=>p[1]), mx=meanOf(X), my=meanOf(Y);
      let sxy=0,sxx=0,syy=0; pts.forEach(([x,y])=>{ sxy+=(x-mx)*(y-my); sxx+=(x-mx)**2; syy+=(y-my)**2; }); const r=sxy/Math.sqrt(sxx*syy), b=sxy/sxx, a=my-b*mx;
      const F=chartFrame(600,300,{l:50,r:14,t:14,b:42}); const xl=Math.min(...X), xh=Math.max(...X), yl=0, yh=Math.max(...Y)*1.05;
      const PX=v=>F.x0+(v-xl)/(xh-xl)*(F.x1-F.x0), PY=v=>F.y0-(v-yl)/(yh-yl)*(F.y0-F.y1);
      const svg=sv('svg',{viewBox:`0 0 ${F.W} ${F.H}`,class:'chart',role:'img','aria-label':`Scatter plot of ${YS[yk]} against ${XS[xk]}: r = ${r2(r)}`});
      niceTicks(yl,yh,5).forEach(t=>svg.append(sv('line',{x1:F.x0,x2:F.x1,y1:PY(t),y2:PY(t),class:'grid'}), sv('text',{x:F.x0-6,y:PY(t)+4,'text-anchor':'end',class:'ax'},String(t))));
      niceTicks(xl,xh,6).forEach(t=>svg.append(sv('text',{x:PX(t),y:F.y0+16,'text-anchor':'middle',class:'ax'},String(t))));
      pts.forEach(([x,y,s,m])=>svg.append(sv('circle',{cx:PX(x),cy:PY(y),r:4.5,class:'dot'}, sv('title',{},`${s}, ${fmtDate(m,'mmm yyyy')}: ${x}, ${y}`))));
      svg.append(sv('line',{x1:PX(xl),y1:PY(a+b*xl),x2:PX(xh),y2:PY(a+b*xh),class:'trend'}));
      svg.append(sv('text',{x:(F.x0+F.x1)/2,y:F.H-4,'text-anchor':'middle',class:'ax-t'},XS[xk]));
      fig.innerHTML=''; fig.append(svg);
      stats.innerHTML=`One dot per site per month. r = <b>${r2(r)}</b>, R² = <b>${Math.round(r*r*100)}%</b>.`;
    };
    xs.onchange=()=>{ xk=xs.value; draw(); }; ys.onchange=()=>{ yk=ys.value; draw(); }; draw();
  },
  anscombe(el){
    const x=[10,8,13,9,11,14,6,4,12,7,5], x4=[8,8,8,8,8,8,8,19,8,8,8];
    const D=[[x,[8.04,6.95,7.58,8.81,8.33,9.96,7.24,4.26,10.84,4.82,5.68]],[x,[9.14,8.14,8.74,8.77,9.26,8.1,6.13,3.1,9.13,7.26,4.74]],[x,[7.46,6.77,12.74,7.11,7.81,8.84,6.08,5.39,8.15,6.42,5.73]],[x4,[6.58,5.76,7.71,8.84,8.47,7.04,5.25,12.5,5.56,7.91,6.89]]];
    const grid=h('div',{class:'w-quad'});
    D.forEach(([xs,ys],k)=>{ const F=chartFrame(260,200,{l:28,r:8,t:22,b:24}), PX=v=>F.x0+(v-2)/18*(F.x1-F.x0), PY=v=>F.y0-(v-2)/12*(F.y0-F.y1);
      const svg=sv('svg',{viewBox:'0 0 260 200',class:'chart',role:'img','aria-label':`Anscombe dataset ${['I','II','III','IV'][k]}`});
      svg.append(sv('text',{x:130,y:14,'text-anchor':'middle',class:'ax-t'},`Dataset ${['I','II','III','IV'][k]}`));
      [4,8,12].forEach(t=>svg.append(sv('line',{x1:F.x0,x2:F.x1,y1:PY(t),y2:PY(t),class:'grid'}), sv('text',{x:F.x0-4,y:PY(t)+4,'text-anchor':'end',class:'ax'},String(t))));
      svg.append(sv('line',{x1:PX(2),y1:PY(4),x2:PX(20),y2:PY(13),class:'trend'}));
      xs.forEach((v,i)=>svg.append(sv('circle',{cx:PX(v),cy:PY(ys[i]),r:4,class:'dot'})));
      grid.append(svg); });
    el.append(grid);
  },
  truncate(el){
    const sh=WB.sheet('Visits'); const cnt={}; sh.rows.forEach(r=>cnt[r[2]]=(cnt[r[2]]||0)+1);
    const sites=WB.sheet('Sites').rows.map(r=>[r[0],r[2]]).filter(([id])=>['S01','S02','S05'].includes(id));
    const cb=h('input',{type:'checkbox','data-testid':'truncate-zero'}); const fig=h('div',{class:'w-fig'});
    el.append(h('div',{class:'w-controls'}, h('label',{class:'small'},cb,' Start the axis at zero')), fig);
    const draw=()=>{ const vals=sites.map(([id])=>cnt[id]); const lo=cb.checked?0:Math.floor(Math.min(...vals)*0.97/50)*50, hi=Math.max(...vals)*1.03;
      const F=chartFrame(520,230,{l:110,r:20,t:10,b:30}), PX=v=>F.x0+(v-lo)/(hi-lo)*(F.x1-F.x0);
      const svg=sv('svg',{viewBox:`0 0 ${F.W} ${F.H}`,class:'chart',role:'img','aria-label':`Visits by site, axis starting at ${lo}`});
      niceTicks(lo,hi,5).forEach(t=>svg.append(sv('line',{x1:PX(t),x2:PX(t),y1:F.y1,y2:F.y0,class:'grid'}), sv('text',{x:PX(t),y:F.y0+16,'text-anchor':'middle',class:'ax'},t.toLocaleString('en-GB'))));
      sites.forEach(([id,town],i)=>{ const y=F.y1+10+i*62; svg.append(sv('rect',{x:F.x0,y,width:Math.max(1,PX(cnt[id])-F.x0),height:40,class:'bar'}), sv('text',{x:F.x0-8,y:y+25,'text-anchor':'end',class:'ax'},town)); });
      fig.innerHTML=''; fig.append(svg); };
    cb.onchange=draw; draw();
    el.append(h('p',{class:'small muted'},'With the axis cut off, one site looks several times busier than another. Tick the box to see the real difference.'));
  },
  colourblind(el){
    const months=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const sm=WB.sheet('SiteMonths').rows; const series=['S01','S03'].map(s=>sm.filter(r=>r[0]===s).map(r=>r[4]));
    const MATS={normal:null, protan:'0.567 0.433 0 0 0  0.558 0.442 0 0 0  0 0.242 0.758 0 0  0 0 0 1 0', deutan:'0.625 0.375 0 0 0  0.7 0.3 0 0 0  0 0.3 0.7 0 0  0 0 0 1 0', grey:'0.299 0.587 0.114 0 0  0.299 0.587 0.114 0 0  0.299 0.587 0.114 0 0  0 0 0 1 0'};
    let mode='normal', safe=false;
    const btns=h('div',{class:'dtabs',role:'group','aria-label':'Simulate'}); const fig=h('div',{class:'w-fig'}); const sf=h('label',{class:'small'},h('input',{type:'checkbox','data-testid':'cb-safe'}),' Use a colour-blind-safe palette (blue and orange)');
    [['normal','Normal vision'],['protan','Protanopia'],['deutan','Deuteranopia'],['grey','Greyscale']].forEach(([k,l])=>btns.append(h('button',{type:'button','aria-pressed':String(k===mode),'data-testid':`cb-${k}`,onclick:()=>{ mode=k; $$('button',btns).forEach(b=>b.setAttribute('aria-pressed',String(b.dataset.testid===`cb-${k}`))); draw(); }},l)));
    el.append(btns, sf, fig);
    const draw=()=>{ const F=chartFrame(560,240,{l:46,r:90,t:12,b:30}); const all=series.flat(), hi=Math.max(...all)*1.1; const PX=i=>F.x0+i/11*(F.x1-F.x0), PY=v=>F.y0-v/hi*(F.y0-F.y1);
      const cols = safe ? ['#1f5fbf','#e07b00'] : ['#d62728','#2ca02c'];
      const id='cbm'+Math.random().toString(36).slice(2,7);
      const svg=sv('svg',{viewBox:`0 0 ${F.W} ${F.H}`,class:'chart',role:'img','aria-label':'Monthly visits at Telford Central and Wolverhampton City'});
      if (MATS[mode]){ const f=sv('filter',{id}); f.append(sv('feColorMatrix',{type:'matrix',values:MATS[mode]})); const d=sv('defs'); d.append(f); svg.append(d); }
      const g=sv('g',MATS[mode]?{filter:`url(#${id})`}:{});
      niceTicks(0,hi,4).forEach(t=>g.append(sv('line',{x1:F.x0,x2:F.x1,y1:PY(t),y2:PY(t),class:'grid'}), sv('text',{x:F.x0-6,y:PY(t)+4,'text-anchor':'end',class:'ax'},String(t))));
      months.forEach((m,i)=>g.append(sv('text',{x:PX(i),y:F.y0+16,'text-anchor':'middle',class:'ax'},m)));
      series.forEach((s,k)=>{ g.append(sv('path',{d:'M'+s.map((v,i)=>`${PX(i)} ${PY(v)}`).join(' L'),stroke:cols[k],'stroke-width':3,fill:'none'})); g.append(sv('text',{x:F.x1+6,y:PY(s[11])+4,fill:cols[k],class:'lbl'},['Telford','Wolverhampton'][k])); });
      svg.append(g); fig.innerHTML=''; fig.append(svg); };
    $('input',sf).onchange=e=>{ safe=e.target.checked; draw(); }; draw();
  },
  schema(el){
    const a=h('div'), b=h('div');
    el.innerHTML=TABS([{label:'Star schema',html:'<div class="erd-scroll" data-slot="star"></div>'},{label:'Snowflake schema',html:'<div class="erd-scroll" data-slot="snow"></div>'}]);
    wireTabs(el);
    renderERD($('[data-slot=star]',el), SCHEMA_STAR, {testPrefix:'star', label:'Star schema: the FactVisit table links to four dimension tables'});
    renderERD($('[data-slot=snow]',el), SCHEMA_SNOW, {testPrefix:'snow', label:'Snowflake schema: dimension tables are split into further tables'});
  },
};
function wireWidgets(root){ $$('[data-widget]',root).forEach(el=>{ if (el.dataset.ready) return; el.dataset.ready='1'; const f=WIDGETS[el.dataset.widget]; if (f) f(el); }); }
function wireTabs(root){
  $$('[data-ltabs]',root).forEach(box=>{ if (box.dataset.wired) return; box.dataset.wired='1';
    const tabs=$$('[role=tab]',box), panels=$$(':scope > [role=tabpanel]',box);
    const pick=k=>tabs.forEach((t,i)=>{ t.setAttribute('aria-selected',String(i===k)); t.tabIndex=i===k?0:-1; panels[i].hidden=i!==k; });
    tabs.forEach((t,i)=>{ t.addEventListener('click',()=>pick(i)); t.addEventListener('keydown',e=>{ if (e.key==='ArrowRight'||e.key==='ArrowLeft'){ e.preventDefault(); const k=(i+(e.key==='ArrowRight'?1:tabs.length-1))%tabs.length; pick(k); tabs[k].focus(); } }); });
  });
}

/* ---------- Progress through the path ---------- */
const FOUND = {
  intro:{need:'Read through, then press Got it, start learning.', prog:()=>[P.introDone?1:0,1], label:'Read'},
  data:{need:'Open at least three sheets.', prog:()=>[Math.min(3,(P.dataSheets||[]).length),3], label:'Explore the data'},
  capstone:{need:'Tick every task in the checklist (Power BI is optional).', prog:()=>{ const req=CAPSTONE_TASKS.filter(t=>!t[2]); return [req.filter(t=>P.capstone[t[0]]).length, req.length]; }, label:'Project'},
};
function lessonProg(id){
  if (FOUND[id]) return FOUND[id].prog();
  const L=LESSONS[id]; let a=0, b=0;
  if (L.ex){ b+=L.ex.length; a+=L.ex.filter((_,i)=>P.exDone[id+'.'+i]).length; }
  if (L.check){ b+=1; a+=P.checkDone[id]?1:0; }
  if (L.tasks){ b+=L.tasks.length; a+=L.tasks.filter((_,i)=>(P.tasks||{})[id+'.'+i]).length; }
  return [a,b];
}
const markedDone = id => !!(P.markedDone||{})[id];
const autoDone = id => { const [a,b]=lessonProg(id); return a>=b; };
const lessonDone = id => autoDone(id) || markedDone(id);
const REQUIRED = PATH.filter(x=>!x.unit.optional);
const nextLesson = () => REQUIRED.find(x=>!lessonDone(x.id));

/* ---------- Routing ---------- */
const VIEWS=['path','intro','data','capstone','lesson','practice','quizzes','quiz'];
let current={view:'path',lesson:null};
function route(){
  const parts=location.hash.replace(/^#\/?/,'').split('/');
  let view='path', lid=null;
  if (parts[0]==='lesson' && LESSONS[parts[1]]){ lid=parts[1]; view=LESSONS[lid].kind||'lesson'; }
  else if (parts[0]==='practice') view='practice';
  else if (parts[0]==='quiz') view = parts[1] ? 'quiz' : 'quizzes';
  if (view!=='quiz') stopTimer();
  current={view,lesson:lid};
  VIEWS.forEach(v=>$('#v-'+v).classList.toggle('on', v===view));
  const section = ['intro','data','capstone','practice'].includes(view)?view : (view==='quiz'||view==='quizzes')?'quiz' : 'path';
  $$('[data-nav]').forEach(b=>{ if (b.dataset.nav===section) b.setAttribute('aria-current','page'); else b.removeAttribute('aria-current'); });
  $('#pathHero').hidden = view!=='path'; $('#main').classList.toggle('on-path', view==='path');
  if (lid){ P.last=lid; save(); }
  if (view==='path') renderPath();
  if (view==='intro') drawIntro();
  if (view==='data') drawData();
  if (view==='capstone') drawCapstone();
  if (view==='lesson') openLesson(lid);
  if (view==='practice') drawPractice();
  if (view==='quizzes') renderQuizList();
  if (view==='quiz') openQuiz(parts[1]);
  renderChrome();
  window.scrollTo({top:0});
  const h1=view==='path'?$('#h-path'):$('#v-'+view+' h1'); if (h1){ h1.setAttribute('tabindex','-1'); h1.focus({preventScroll:true}); }
}
addEventListener('hashchange', route);
function goto(to){ const hs = (to==='path'||to==='practice'||to==='quiz'||to.startsWith('quiz/')) ? '#/'+to : '#/lesson/'+to; if (location.hash===hs) route(); else location.hash=hs; }
document.addEventListener('click', e=>{ const b=e.target.closest('[data-go]'); if (b){ e.preventDefault(); goto(b.dataset.go); } });

function renderChrome(){
  const cr=$('#crumbs'), foot=$('#lessonFoot'), topNav=$('#lessonNav'); topNav.innerHTML=''; topNav.hidden=true;
  const id=current.lesson;
  if (!id){
    foot.innerHTML='';
    const PAGE={practice:'Excel practice',quizzes:'Quizzes'}[current.view];
    if (!PAGE){ cr.hidden=true; return; }
    cr.hidden=false; cr.innerHTML=''; cr.append(h('a',{href:'#/path'},'Home'), h('span',{class:'sep','aria-hidden':'true'},'/'), h('span',{},PAGE)); return;
  }
  const k=PATH.findIndex(x=>x.id===id), item=PATH[k], un=UNITS.indexOf(item.unit)+1;
  cr.hidden=false; cr.innerHTML='';
  cr.append(h('a',{href:'#/path','data-testid':'breadcrumb-home'},'Home'), h('span',{class:'sep','aria-hidden':'true'},'/'), h('span',{'data-testid':'breadcrumb-unit'},`Unit ${un}: ${item.unit.title}`), h('span',{class:'pos','data-testid':'lesson-position'}, `Lesson ${k+1} of ${PATH.length}`));
  topNav.hidden=false;
  const prevL=k>0?PATH[k-1].id:null, nextL=k<PATH.length-1?PATH[k+1].id:null;
  topNav.append(
    prevL ? h('button',{class:'ln-prev',type:'button','data-go':prevL,title:LESSONS[prevL].title}, h('span',{class:'arr','aria-hidden':'true'},'← '), h('span',{class:'lbl'},'Previous'), h('span',{class:'ttl'},': '+LESSONS[prevL].title)) : h('span',{class:'ln-gap'}),
    nextL ? h('button',{class:'ln-next',type:'button','data-go':nextL,title:LESSONS[nextL].title}, h('span',{class:'lbl'},'Next'), h('span',{class:'ttl'},': '+LESSONS[nextL].title), h('span',{class:'arr','aria-hidden':'true'},' →')) : h('span',{class:'ln-gap'}));
  foot.innerHTML='';
  const L=LESSONS[id], done=lessonDone(id), [a,b]=lessonProg(id), auto=autoDone(id), manual=markedDone(id);
  const setMark = v => { P.markedDone=P.markedDone||{}; if (v) P.markedDone[id]=true; else delete P.markedDone[id]; save(); refreshChrome(); toast(v?'Marked as complete':'No longer marked as complete'); };
  const mark = auto ? null : manual ? h('button',{class:'btn small ghost',type:'button','data-testid':'unmark-complete',onclick:()=>setMark(false)},'Undo: mark as not complete')
    : h('button',{class:'btn small mark',type:'button','data-testid':'mark-complete',onclick:()=>setMark(true)},'✓ Mark lesson as complete');
  const need = FOUND[id] ? FOUND[id].need : [L.tasks?'tick everything in the checklist':null, L.ex?`solve all ${L.ex.length} ${L.noGrid?'tasks':'exercises'} (${L.ex.filter((_,i)=>P.exDone[id+'.'+i]).length} done)`:null, L.check?'answer every Check yourself question correctly':null].filter(Boolean).join(', and ')+'.';
  const state = done ? h('div',{class:'state ok','data-testid':'lesson-status','data-done':'true'}, manual&&!auto?'Marked as complete.':'Lesson complete.', mark?' ':'', mark||'')
    : h('div',{class:'state','data-testid':'lesson-status','data-done':'false'}, FOUND[id]?`To complete this lesson: ${need}`:`To complete this lesson, ${need}`, h('span',{class:'or'},' Or '), mark);
  const nexts=h('div',{class:'nexts'}, h('button',{class:'btn ghost',type:'button','data-go':'path'},'Home'));
  if (k>0) nexts.append(h('button',{class:'btn ghost',type:'button','data-go':PATH[k-1].id,'data-testid':'previous-lesson'},'Previous lesson'));
  nexts.append(k<PATH.length-1 ? h('button',{class:'btn'+(done?' primary':''),type:'button','data-go':PATH[k+1].id,'data-testid':'next-lesson'}, `Next: ${LESSONS[PATH[k+1].id].title}`) : h('button',{class:'btn primary',type:'button','data-go':'path'},'Back to home'));
  foot.append(h('div',{class:'lfoot','data-testid':'lesson-footer'}, state, nexts));
}
function refreshChrome(){ renderChrome(); if (current.view==='lesson') renderLessonTabs(); }

/* ---------- Home: the learning path ---------- */
function renderPath(){
  const v=$('#v-path'); v.innerHTML='';
  const doneL=REQUIRED.filter(x=>lessonDone(x.id)).length, nxt=nextLesson();
  const exTotal=REQUIRED.reduce((n,x)=>n+((LESSONS[x.id].ex||[]).length),0), exDone=Object.keys(P.exDone).filter(k=>P.exDone[k]&&REQUIRED.some(x=>x.id===k.split('.')[0])).length;
  const card=h('div',{class:'hero-card','data-testid':'continue-card'});
  card.append(h('div',{class:'k'}, nxt?(doneL?'Continue with':'Start with'):'Path complete'), h('div',{class:'next'}, nxt?LESSONS[nxt.id].title:'You\'ve finished every lesson. Try the final quiz.'));
  const trail=h('div',{class:'trail','aria-hidden':'true'});
  UNITS.filter(u=>!u.optional).forEach((u,ui)=>{ if (ui) trail.append(h('b')); u.lessons.forEach(id=>trail.append(h('i',{class:(lessonDone(id)?'d ':'')+(nxt&&nxt.id===id?'n':''),title:LESSONS[id].title}))); });
  card.append(trail, h('div',{class:'hero-stats','data-testid':'overall-progress'}, `${doneL} of ${REQUIRED.length} lessons complete. ${exDone} of ${exTotal} Excel exercises solved.`),
    h('div',{}, nxt ? h('button',{class:'btn primary',type:'button','data-go':nxt.id,'data-testid':'continue'}, doneL?'Continue':'Start the path') : h('button',{class:'btn primary',type:'button','data-go':'quiz/final'},'Final quiz')));
  v.append(h('div',{class:'path-hero'}, card));
  const ha=$('#heroActions'); ha.innerHTML='';
  ha.append(h('button',{class:'btn primary',type:'button','data-go':nxt?nxt.id:'quiz','data-testid':'hero-start'}, nxt?(doneL?'Continue learning':'Start the learning path'):'Take a quiz'),
    h('button',{class:'btn light',type:'button','data-go':'intro'},'How this site works'),
    h('button',{class:'btn light',type:'button','data-go':'data'},'Meet the data'),
    h('button',{class:'btn light',type:'button','data-go':'practice'},'Excel practice'),
    h('button',{class:'btn light',type:'button','data-go':'quiz'},'Take a quiz'));
  let num=0;
  UNITS.forEach((u,ui)=>{
    const ol=h('ol',{class:'route'});
    u.lessons.forEach(id=>{
      const L=LESSONS[id]; num++;
      const isDone=lessonDone(id), isNext=nxt&&nxt.id===id, [a,b]=lessonProg(id);
      const status = isDone?'Done' : isNext?'Up next' : a>0 ? (FOUND[id]?'Started':`${a} of ${b} done`) : FOUND[id]?FOUND[id].label : L.tasks ? 'Set-up checklist' : L.ex ? `${L.ex.length} ${L.noGrid?'tasks':'exercises'}${L.excel||L.noGrid?' in Excel':''}` : `Read and ${L.check.length} questions`;
      const node=h('span',{class:'node','aria-hidden':'true'}, isDone?h('span',{class:'tick'}):h('span',{},String(num)));
      ol.append(h('li',{class:(isDone?'done ':'')+(isNext?'next':''),'data-testid':`lesson-${id}`},
        h('a',{href:'#/lesson/'+id}, node, h('span',{class:'t'}, L.title, h('small',{},L.summary)), h('span',{class:'s','data-testid':`lesson-status-${id}`},status))));
    });
    const uq=QUIZZES.find(q=>q.unitQuiz===ui+1), rec=uq&&P.quizzes[uq.id];
    v.append(h('section',{class:'unit'+(u.optional?' optional':''),'aria-label':`Unit ${ui+1}: ${u.title}`,'data-testid':`unit-${ui+1}`},
      h('div',{class:'unit-h'}, h('div',{class:'n'},`Unit ${ui+1}`, u.optional?h('span',{class:'opt-badge'},'Optional'):null), h('h2',{},u.title), h('p',{},u.blurb), h('div',{class:'meta'}, `${u.lessons.filter(lessonDone).length} of ${u.lessons.length} lessons complete`),
        uq ? h('div',{class:'unit-quiz'}, h('button',{class:'btn small',type:'button','data-go':'quiz/'+uq.id,'data-testid':`unit-quiz-${ui+1}`},`Unit ${ui+1} quiz`), h('span',{class:'small'}, rec?`Best ${rec.best.percent}%`:`${uq.questions.length} questions`)) : null),
      ol));
  });
  v.append(h('p',{class:'progress-note'}, 'Your progress is saved in this browser only. ', doneL||exDone ? h('button',{class:'linklike',type:'button',onclick:resetAll,'data-testid':'reset-progress'},'Reset my progress') : null));
}
function resetAll(){
  if (!confirm('Reset all your progress and answers? This can\'t be undone.')) return;
  for (const k of Object.keys(P)) delete P[k];
  Object.assign(P,{exDone:{},drafts:{},checks:{},checkDone:{},quizzes:{},dataSheets:[],capstone:{},mode:'here'}); save(); route(); toast('Progress reset');
}

/* ---------- Check yourself questions ---------- */
function renderCheck(form, Q, tp, saved, onAnswer){
  form.innerHTML='';
  Q.forEach((item,qi)=>{
    const fs=h('fieldset',{class:'q','data-testid':`${tp}-question-${qi+1}`}, h('legend',{html:`${qi+1}. ${item.q}`}));
    const fb=h('div',{class:'fb','aria-live':'polite','data-testid':`${tp}-question-${qi+1}-feedback`});
    const mark=(ai, silent)=>{
      $$('label',fs).forEach((l,i)=>{ l.classList.remove('right','wrong'); if (i===ai) l.classList.add(ai===item.c?'right':'wrong'); });
      $$('input',fs)[ai].checked=true;
      fb.className='fb '+(ai===item.c?'good':'bad'); fb.innerHTML=ai===item.c?'Correct. '+item.why:'Not quite. Try another answer.';
      if (!silent) onAnswer(qi, ai);
    };
    item.a.forEach((txt,ai)=>{ const inp=h('input',{type:'radio',name:`${tp}-q${qi}`,value:ai,'data-testid':`${tp}-question-${qi+1}-option-${'abcd'[ai]}-input`}); inp.addEventListener('change',()=>mark(ai));
      fs.append(h('label',{'data-testid':`${tp}-question-${qi+1}-option-${'abcd'[ai]}`}, inp, h('span',{html:txt}))); });
    fs.append(fb); form.append(fs);
    if (saved[qi]!=null) mark(saved[qi], true);
  });
}

/* ---------- Foundations pages ---------- */
function drawIntro(){ $('#introBody').innerHTML=SITE_INTRO; $('#introStart').onclick=()=>{ P.introDone=true; save(); goto('data'); }; }
let dataView=null;
function drawData(){
  const b=$('#dataBody'); if (!b.childElementCount) b.innerHTML=DATA_INTRO;
  const e=$('#dataErd'); if (!e.childElementCount) renderERD(e, DATA_ERD, {testPrefix:'data-erd', label:'How the Sparta Leisure sheets link'});
  if (!dataView){ dataView=makeSheetView($('#dataSheets'),{sheet:'Sites', editable:false, testid:'data-sheet', onView:n=>{ if (!P.dataSheets.includes(n)){ P.dataSheets.push(n); save(); renderChrome(); } }}); }
  dataView.set({sheet:'Sites'});
  if (!P.dataSheets.includes('Sites')){ P.dataSheets.push('Sites'); save(); }
}
function drawCapstone(){
  const b=$('#capstoneBody'); if (!b.childElementCount) b.innerHTML=CAPSTONE_HTML;
  const ul=$('#capChecklist'); ul.innerHTML='';
  CAPSTONE_TASKS.forEach(([id,label,optional])=>{
    const cb=h('input',{type:'checkbox',id:`cap-${id}`,'data-testid':`capstone-task-${id}`,checked:!!P.capstone[id]});
    cb.addEventListener('change',()=>{ P.capstone[id]=cb.checked; save(); refreshChrome(); });
    ul.append(h('li',{class:optional?'optional':''}, h('label',{for:`cap-${id}`}, cb, h('span',{},label), optional?h('em',{class:'opt'},'optional'):'')));
  });
}
$('#capPrint').addEventListener('click',()=>window.print());

/* ---------- Lessons ---------- */
let lView=null, curLesson=null, curTab='learn', reveal={hint:false,sol:false};
function openLesson(id){
  const L=LESSONS[id];
  if (curLesson!==id){ curLesson=id; curTab='learn'; }
  $('#lTitle').textContent=L.title; $('#lSummary').textContent=L.summary;
  const excelLesson=!!L.ex && !L.noGrid;
  $('#lGrid').classList.toggle('single', !excelLesson); $('#lSide').hidden=!excelLesson;
  if (excelLesson && !lView) lView=makeSheetView($('#lSheet'),{sheet:L.sheet||'Members', testid:'lesson-sheet', onRun:onLessonRun, onChange:v=>{ P.drafts[draftKey()]=v; save(); }});
  renderLessonTabs(); renderLessonBody();
}
const draftKey=()=>curLesson+':'+curTab;
function renderLessonTabs(){
  const L=LESSONS[curLesson], t=$('#lTabs'); t.innerHTML='';
  t.hidden=!L.ex;
  if (!L.ex) return;
  const word=L.noGrid?'Task':'Exercise';
  ['learn',...L.ex.map((_,i)=>i)].forEach(k=>{
    const done = k!=='learn' && P.exDone[curLesson+'.'+k];
    const b=h('button',{type:'button',role:'tab','aria-selected':String(curTab===k),class:done?'done':'','data-testid':k==='learn'?'tab-learn':`tab-exercise-${k+1}`},
      k==='learn'?'Learn':[h('span',{class:'dotc','aria-hidden':'true'}), `${word} ${k+1}`, done?h('span',{class:'sr'},' (solved)'):null]);
    b.addEventListener('click',()=>setTab(k)); t.append(b);
  });
}
function setTab(k){ curTab=k; renderLessonTabs(); renderLessonBody(); if (innerWidth<1020) $('#lTabs').scrollIntoView({block:'start'}); }
function renderLessonBody(){
  const L=LESSONS[curLesson], body=$('#lBody'); body.innerHTML=''; reveal={hint:false,sol:false};
  if (curTab==='learn'){
    body.innerHTML='<div class="prose">'+L.learn+'</div>'; wireTabs(body); wireWidgets(body);
    $$('.fx-ex',body).forEach((el,n)=>{
      const f=el.dataset.formula, sheet=el.dataset.sheet, fill=!!el.dataset.fill, col=el.dataset.col||'Try it';
      const note=$('.fx-note',el);
      el.innerHTML=''; el.append(h('code',{class:'fx-f'},f), h('span',{class:'fx-on'},`on ${sheet}`));
      if (lView) el.append(h('button',{class:'btn small run',type:'button','data-testid':`try-example-${n+1}`,onclick:()=>{ lView.set({sheet, col, fill, formula:f, run:true, vtype:null}); if (innerWidth<1020) $('#lSheet').scrollIntoView({block:'start'}); }},'Try this'));
      if (note) el.append(note);
    });
    if (L.check){
      const form=h('form',{class:'quiz',onsubmit:e=>e.preventDefault(),'data-testid':'lesson-check'});
      body.append(h('div',{class:'rule'}), h('h2',{class:'dot'},'Check yourself'), form);
      const saved=(P.checks[curLesson]=P.checks[curLesson]||{}), id=curLesson;
      renderCheck(form, L.check, 'lesson-check', saved, (qi,ai)=>{ saved[qi]=ai; P.checkDone[id]=L.check.every((x,i)=>saved[i]===x.c); save(); refreshChrome(); });
    }
    if (L.tasks){
      const ul=h('ul',{class:'cap-checklist','data-testid':'setup-checklist'}); P.tasks=P.tasks||{};
      L.tasks.forEach((t,k)=>{ const key=curLesson+'.'+k, cb=h('input',{type:'checkbox',id:'task-'+k,'data-testid':'setup-task-'+(k+1),checked:!!P.tasks[key]});
        cb.addEventListener('change',()=>{ P.tasks[key]=cb.checked; save(); refreshChrome(); }); ul.append(h('li',{}, h('label',{for:'task-'+k}, cb, h('span',{html:t})))); });
      body.append(h('div',{class:'rule'}), h('h2',{class:'dot'},'Your checklist'), ul);
    }
    if (L.ex){
      body.append(h('p',{style:'margin-top:26px'}, h('button',{class:'btn primary',type:'button','data-testid':'start-exercises',onclick:()=>setTab(0)}, `Start the ${L.noGrid?'tasks':'exercises'} (${L.ex.length})`)));
      if (L.noGrid) return;
      const ex0=$('.fx-ex',body);
      lView.set({sheet:L.sheet||L.ex[0].sheet, col:'Try it', fill:false, formula:P.drafts[draftKey()] ?? '', vtype:null});
      lView.showFill(true);
    }
    return;
  }
  const i=curTab, ex=L.ex[i], mode = L.noGrid ? 'excel' : (ex.excel||L.excel) ? (P.modeSet&&P.modeSet[curLesson+'.'+i] || 'excel') : P.mode;
  body.dataset.solution=ex.sol;
  body.append(h('div',{class:'ex-num','data-testid':'exercise-number'}, `${L.noGrid?'Task':'Exercise'} ${i+1} of ${L.ex.length}`));
  if (P.exDone[curLesson+'.'+i]) body.append(h('div',{class:'solved','data-testid':'solved'},'Solved'));
  body.append(h('p',{class:'ex-prompt','data-testid':'exercise-prompt'}, ...fmtPrompt(ex.p)));
  const toggle=h('div',{class:'mode-toggle',role:'group','aria-label':'How do you want to answer?'});
  [['here','Practise here'],['excel','In Excel']].forEach(([k,l])=>toggle.append(h('button',{type:'button','aria-pressed':String(mode===k),'data-testid':`mode-${k}`,onclick:()=>{ if (ex.excel||L.excel){ P.modeSet=P.modeSet||{}; P.modeSet[curLesson+'.'+i]=k; } else P.mode=k; save(); renderLessonBody(); }},l)));
  if (!L.noGrid) body.append(toggle);
  if (ex.steps) body.append(h('ol',{class:'steps','data-testid':'task-steps'}, ex.steps.map(t=>h('li',{html:t}))));
  const panel=h('div',{class:'mode-panel','data-testid':`mode-panel-${mode}`});
  body.append(panel);
  if (mode==='here'){
    panel.append(h('p',{class:'small'}, ex.fill ? `Type your formula for row 2 of the pink ${ex.col||'Answer'} column on the ${ex.sheet} sheet. It's filled down and every row is checked.` : `Type your formula for the pink answer cell on the ${ex.sheet} sheet, then press Enter.`, ex.excel?' This task is meant for Excel, but you can also work it out with a formula here.':''));
    panel.append(h('div',{id:'lVerdict'}));
  } else {
    const ask = ex.fill ? (()=>{ const sh=WB.sheet(ex.sheet), r=askRow(ex); return `Fill your formula down in Excel. What does it give in row ${r} (${sh.cols[0].name} ${cellValue(sh,r,1)})?`; })() : 'Type the answer you get in Excel.';
    const want=expectedAt(ex);
    const ph = ex.vtype==='date' ? 'For example 31/12/2025' : typeof want==='boolean' ? 'TRUE or FALSE' : typeof want==='string' ? 'Type the text' : 'Type a number';
    const inp=h('input',{type:'text',class:'answer-input',id:'excelAnswer','aria-label':'Your answer','data-testid':'excel-answer',placeholder:ph,value:P.drafts[draftKey()+':x']||''});
    const go=()=>{ P.drafts[draftKey()+':x']=inp.value; save(); const v=markTyped(ex, inp.value); showVerdict(v); };
    inp.addEventListener('keydown',e=>{ if (e.key==='Enter') go(); });
    panel.append(h('div',{class:'excel-box'}, L.noGrid ? null : h('p',{}, DL_WORKBOOK_NODE(), ' ', h('span',{class:'small muted'},`Open the ${ex.sheet} sheet.`)), h('label',{for:'excelAnswer',class:'q-instr'}, ex.ask || ask), h('div',{class:'answer-row'}, inp, h('button',{class:'btn primary',type:'button','data-testid':'check-answer',onclick:go},'Check'))), h('div',{id:'lVerdict'}));
  }
  body.append(h('div',{class:'ex-actions'},
    h('button',{class:'btn small',type:'button',id:'bHint','data-testid':'show-hint',onclick:()=>{ reveal.hint=!reveal.hint; drawReveal(); }},'Show a hint'),
    h('button',{class:'btn small',type:'button',id:'bSol','data-testid':'show-answer',onclick:()=>{ reveal.sol=!reveal.sol; drawReveal(); }},'Show the answer')));
  body.append(h('div',{id:'lReveal'}));
  if (i===0 && !L.noGrid) body.append(h('p',{class:'small'}, 'Stuck? The ', h('button',{type:'button',class:'linkish',onclick:()=>setTab('learn')},'Learn tab'), ' has worked examples you can try.'));
  if (L.noGrid) return;
  lView.set({sheet:ex.sheet, col:ex.col||'Answer', fill:!!ex.fill, vtype:ex.vtype||null, formula:P.drafts[draftKey()] ?? ''});
  lView.showFill(false);
}
const DL_WORKBOOK_NODE = () => h('a',{class:'btn small',href:'downloads/sparta-leisure.xlsx',download:'','data-testid':'download-workbook'},'Download the workbook');
function drawReveal(){
  const ex=LESSONS[curLesson].ex[curTab], r=$('#lReveal'); r.innerHTML='';
  $('#bHint').textContent=reveal.hint?'Hide the hint':'Show a hint'; $('#bSol').textContent=reveal.sol?'Hide the answer':'Show the answer';
  if (reveal.hint) r.append(h('div',{class:'reveal','data-testid':'hint'}, h('b',{},'Hint: '), ...fmtPrompt(ex.hint)));
  if (reveal.sol){
    const want=expectedAt(ex);
    if (ex.how){ r.append(h('div',{class:'reveal','data-testid':'answer'}, h('b',{},'How to do it: '), h('span',{html:ex.how}), h('p',{class:'small'}, 'The answer is ', h('b',{},fmtVal(want, ex.vtype)), '.'))); return; }
    r.append(h('div',{class:'reveal','data-testid':'answer'}, h('b',{},'One possible formula'), h('pre',{},ex.sol),
      h('p',{class:'small'}, ex.fill?`In row ${askRow(ex)} it gives `:'It gives ', h('b',{},fmtVal(want, ex.vtype)), '.'),
      h('button',{class:'btn small',type:'button','data-testid':'copy-answer',onclick:()=>{ P.mode==='excel'&&!(ex.excel)?null:null; lView.set({formula:ex.sol}); lView.focus(); }},'Copy into the grid')));
  }
}
function nextUnsolved(after){ const L=LESSONS[curLesson]; for (let k=1;k<=L.ex.length;k++){ const j=(after+k)%L.ex.length; if (!P.exDone[curLesson+'.'+j]) return j; } return null; }
function onLessonRun(formula){
  if (curTab==='learn' || current.view!=='lesson') return;
  const L=LESSONS[curLesson], ex=L.ex[curTab];
  if (!$('[data-testid="mode-panel-here"]')) return;   // In Excel mode, the grid is just for looking
  showVerdict(markFormula(ex, formula));
}
function showVerdict(v){
  const out=$('#lVerdict'); if (!out) return; out.innerHTML='';
  const L=LESSONS[curLesson], i=curTab;
  const box=h('div',{class:'verdict '+(v.ok?'good':'bad'),role:'status','data-testid':'verdict','data-correct':String(v.ok)});
  if (v.ok){
    const was=lessonDone(curLesson); if (!P.exDone[curLesson+'.'+i]){ P.exDone[curLesson+'.'+i]=true; save(); }
    const k=PATH.findIndex(x=>x.id===curLesson); let btn=null, sub=v.rounded?'That\'s right, to the precision you gave.':'That matches the expected answer.';
    if (lessonDone(curLesson) && !was){ sub='That\'s the last one. Lesson complete.'; btn=k<PATH.length-1?h('button',{class:'btn small primary',type:'button','data-go':PATH[k+1].id},'Next lesson'):null; toast('Lesson complete'); }
    else { const j=nextUnsolved(i); if (j!=null) btn=h('button',{class:'btn small primary',type:'button',onclick:()=>setTab(j)},`${L.noGrid?'Task':'Exercise'} ${j+1}`); else if (k<PATH.length-1) btn=h('button',{class:'btn small primary',type:'button','data-go':PATH[k+1].id},'Next lesson'); }
    box.append(h('div',{}, 'Correct.', h('span',{},sub)));
    if (btn){ btn.setAttribute('data-testid','verdict-next'); btn.style.marginLeft='auto'; box.append(btn); box.style.alignItems='center'; }
    if (!$('.solved',$('#lBody'))) $('.ex-num',$('#lBody')).after(h('div',{class:'solved','data-testid':'solved'},'Solved'));
    refreshChrome();
  } else box.append(h('div',{}, 'Not yet. '+v.msg, v.tip?h('span',{},v.tip):null));
  out.append(box);
}

/* ---------- Excel practice page and the data pop-up ---------- */
let pView=null;
function drawPractice(){ if (!pView){ pView=makeSheetView($('#pSheet'),{sheet:'Members', testid:'practice', onChange:v=>{ P.drafts.practice=v; save(); }}); pView.showFill(true); pView.set({sheet:'Members', col:'Try it', formula:P.drafts.practice||'=COUNTIF(J:J,"Active")', run:true}); } }
let dView=null;
function openData(sheet){ const d=$('#dataDialog'); if (!dView) dView=makeSheetView($('#dialogSheet'),{sheet, testid:'dialog-sheet'}); dView.showFill(true); dView.set({sheet, col:'Try it', formula:''}); d.showModal(); }
$('#data-dialog-close').addEventListener('click',()=>$('#dataDialog').close());
$('#dataDialog').addEventListener('click',e=>{ if (e.target.id==='dataDialog') e.target.close(); });

const setMast=()=>document.documentElement.style.setProperty('--mast-h', $('.mast').offsetHeight+'px');
new ResizeObserver(setMast).observe($('.mast'));
function boot(){ if (!location.hash) history.replaceState(null,'','#/path'); route(); }
