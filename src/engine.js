/* ---------- Excel formula engine ----------
   Evaluates Excel formulas against the Sparta Leisure workbook. No DOM here, so scripts/check.mjs runs the same
   code in Node. A sheet is { name, cols:[{name,type}], rows:[[…]] }: row 1 is the header, data starts on row 2.
   Values are numbers, strings, booleans, null (blank) or XErr. Dates are Excel serial numbers (days since 30/12/1899). */
class XErr { constructor(code, why){ this.code=code; this.why=why||''; } toString(){ return this.code; } }
const ERR = (c, why) => new XErr(c, why);
const isErr = v => v instanceof XErr;

/* ----- Workbook access ----- */
function makeWorkbook(data){
  const byName = new Map();
  for (const s of data.sheets) byName.set(s.name.toLowerCase(), s);
  return { sheets:data.sheets, sheet:n => byName.get(String(n).toLowerCase()) };
}
function cellValue(sh, r, c){                       // r and c are 1-based, as in Excel
  if (!sh || r < 1 || c < 1 || c > sh.cols.length) return null;
  if (r === 1){ const n=sh.cols[c-1].name; return sh.headerNumbers && /^\d+$/.test(n) ? +n : n; }
  const row = sh.rows[r-2]; if (!row) return null;
  const v = row[c-1]; return v === undefined || v === '' ? null : v;
}
const lastRow = sh => sh.rows.length + 1;
const colLetter = n => { let s=''; while (n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s; };
const colNumber = s => { let n=0; for (const ch of s.toUpperCase()) n=n*26+(ch.charCodeAt(0)-64); return n; };

/* ----- Tokenizer ----- */
const REF_RE = /^(?:('(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})(?![A-Za-z0-9_(.])/;
function tokenize(src){
  const toks=[]; let i=0;
  while (i < src.length){
    const ch=src[i], rest=src.slice(i);
    if (/\s/.test(ch)){ i++; continue; }
    if (ch === '"'){ let j=i+1, s=''; for(;;){ if (j>=src.length) throw ERR('#ERROR', 'A text value is missing its closing quote mark (")'); if (src[j]==='"'){ if (src[j+1]==='"'){ s+='"'; j+=2; continue; } break; } s+=src[j++]; } toks.push({t:'str', v:s}); i=j+1; continue; }
    if (src[i]==='“'||src[i]==='”') throw ERR('#ERROR', 'Use straight quote marks (") rather than curly ones (“ ”). Curly quotes often appear when a formula is copied from a document.');
    if (ch === '#'){ const m=rest.match(/^#(N\/A|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|NULL!)/i); if (m){ toks.push({t:'err', v:m[0].toUpperCase()}); i+=m[0].length; continue; } }
    let m;
    if ((m = rest.match(/^\d+(\.\d+)?([eE][+-]?\d+)?|^\.\d+/)) && !/^\d+:/.test(rest) && !/^\d+[A-Za-z]/.test(rest)){ toks.push({t:'num', v:parseFloat(m[0])}); i+=m[0].length; continue; }
    if ((m = rest.match(REF_RE))){
      const before = toks[toks.length-1];
      // A bare name like "SUM" followed by "(" is a function, which REF_RE excludes with its lookahead
      toks.push({t:'ref', sheet:m[1]?m[1].replace(/^'|'$/g,'').replace(/''/g,"'"):null, ref:m[2], text:m[0]}); i+=m[0].length; continue;
    }
    if ((m = rest.match(/^(?:_xlfn\.|_xlws\.)?([A-Za-z][A-Za-z0-9_.]*)\s*\(/))){ toks.push({t:'fn', v:m[1].toUpperCase()}); i+=m[0].length; continue; }
    if ((m = rest.match(/^(TRUE|FALSE)(?![A-Za-z0-9_(])/i))){ toks.push({t:'bool', v:m[1].toUpperCase()==='TRUE'}); i+=m[0].length; continue; }
    if ((m = rest.match(/^(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!/))) throw ERR('#REF!', `“${m[0]}” isn't followed by a cell or range, such as ${m[0]}A2 or ${m[0]}B:B`);
    if ((m = rest.match(/^<>|^<=|^>=|^[-+*/^&=<>%(),:;{}]/))){ toks.push({t:'op', v:m[0]===';'?',':m[0]}); i+=m[0].length; continue; }
    if ((m = rest.match(/^[A-Za-z_][A-Za-z0-9_.]*/))) throw ERR('#NAME?', `Excel doesn't recognise “${m[0]}”. Text needs double quotes ("${m[0]}"), and a function needs brackets after its name.`);
    throw ERR('#ERROR', `Unexpected character “${ch}”`);
  }
  return toks;
}
function parseRef(sheet, ref){
  const one = s => { const m=s.match(/^(\$?)([A-Za-z]{1,3})(\$?)(\d*)$/); return { ca:!!m[1], c:colNumber(m[2]), ra:!!m[3], r:m[4]?+m[4]:null }; };
  const [a,b] = ref.split(':'); const A=one(a), B=b?one(b):null;
  if (!B) return { t:'ref', sheet, r:A.r, c:A.c, ra:A.ra, ca:A.ca };
  return { t:'range', sheet, r1:A.r, c1:A.c, r2:B.r, c2:B.c, ra1:A.ra, ca1:A.ca, ra2:B.ra, ca2:B.ca, whole:A.r==null };
}

/* ----- Parser (Excel precedence: comparison < & < + - < * / < ^ < % < unary minus) ----- */
function parse(src){
  let s = String(src).trim();
  if (s.startsWith('=')) s=s.slice(1); else if (s.startsWith('+')) s=s.slice(1);
  if (!s) throw ERR('#ERROR', 'The formula is empty.');
  const toks=tokenize(s); let p=0;
  const peek=()=>toks[p], next=()=>toks[p++];
  const isOp=v=>peek()&&peek().t==='op'&&peek().v===v;
  const expect=v=>{ if (!isOp(v)) throw ERR('#ERROR', v===')'?'A bracket is missing: check every ( has a matching ).':`Expected “${v}”.`); p++; };
  function cmp(){ let a=cat(); while (peek()&&peek().t==='op'&&['=','<>','<','>','<=','>='].includes(peek().v)){ const op=next().v; a={t:'bin',op,a,b:cat()}; } return a; }
  function cat(){ let a=add(); while (isOp('&')){ p++; a={t:'bin',op:'&',a,b:add()}; } return a; }
  function add(){ let a=mul(); while (isOp('+')||isOp('-')){ const op=next().v; a={t:'bin',op,a,b:mul()}; } return a; }
  function mul(){ let a=pow(); while (isOp('*')||isOp('/')){ const op=next().v; a={t:'bin',op,a,b:pow()}; } return a; }
  function pow(){ let a=pct(); while (isOp('^')){ p++; a={t:'bin',op:'^',a,b:pct()}; } return a; }
  function pct(){ let a=un(); while (isOp('%')){ p++; a={t:'pct',a}; } return a; }
  function un(){ if (isOp('-')){ p++; return {t:'neg',a:un()}; } if (isOp('+')){ p++; return un(); } return prim(); }
  function prim(){
    const t=next(); if (!t) throw ERR('#ERROR', 'The formula ends too early: something is missing after the last operator.');
    if (t.t==='num') return {t:'num',v:t.v};
    if (t.t==='str') return {t:'str',v:t.v};
    if (t.t==='bool') return {t:'bool',v:t.v};
    if (t.t==='err') return {t:'errv',v:t.v};
    if (t.t==='ref') return parseRef(t.sheet, t.ref);
    if (t.t==='fn'){
      const args=[];
      if (isOp(')')){ p++; return {t:'fn',name:t.v,args}; }
      for(;;){
        if (isOp(',')||isOp(')')) args.push({t:'missing'}); else args.push(cmp());
        if (isOp(',')){ p++; continue; } expect(')'); break;
      }
      return {t:'fn',name:t.v,args};
    }
    if (t.t==='op' && t.v==='('){ const e=cmp(); expect(')'); return e; }
    if (t.t==='op' && t.v==='{'){ const rows=[[]]; for(;;){ const v=un(); if (!['num','str','bool','neg'].includes(v.t)) throw ERR('#ERROR','An array constant can only hold values.'); rows[rows.length-1].push(v.t==='neg'?-v.a.v:v.v); if (isOp(',')){ p++; continue; } if (isOp(';')){ p++; rows.push([]); continue; } expect('}'); break; } return {t:'arrc',rows}; }
    throw ERR('#ERROR', `Unexpected “${t.v ?? t.t}”.`);
  }
  const e=cmp();
  if (p<toks.length) throw ERR('#ERROR', toks[p].t==='op'&&toks[p].v===')' ? 'There\'s a ) with no matching (.' : `Unexpected “${toks[p].text||toks[p].v}”. Check for a missing operator or comma.`);
  return e;
}

/* ----- Values ----- */
class Range { constructor(wb, sh, r1, c1, r2, c2){ this.wb=wb; this.sh=sh; this.r1=Math.min(r1,r2); this.r2=Math.max(r1,r2); this.c1=Math.min(c1,c2); this.c2=Math.max(c1,c2); }
  get rows(){ return this.r2-this.r1+1; } get cols(){ return this.c2-this.c1+1; }
  at(i, j){ return cellValue(this.sh, this.r1+i, this.c1+j); }
  *values(){ const end=Math.min(this.r2, lastRow(this.sh)); for (let r=this.r1;r<=end;r++) for (let c=this.c1;c<=this.c2;c++) yield cellValue(this.sh,r,c); }
  // Whole-column ranges stop at the last row of data, so they're fast
  get usedRows(){ return Math.max(0, Math.min(this.r2, lastRow(this.sh)) - this.r1 + 1); }
}
class Arr { constructor(rows){ this.data=rows; } get rows(){ return this.data.length; } get cols(){ return this.data[0]?this.data[0].length:0; } at(i,j){ return this.data[i][j]; } *values(){ for (const r of this.data) yield* r; } get usedRows(){ return this.rows; } }
const isRange = v => v instanceof Range || v instanceof Arr;
const toArr = v => { if (v instanceof Arr) return v; if (v instanceof Range){ const n=v.usedRows, rows=[]; for (let i=0;i<n;i++){ const r=[]; for (let j=0;j<v.cols;j++) r.push(v.at(i,j)); rows.push(r); } return new Arr(rows); } return new Arr([[v]]); };

function numberToText(n){
  if (Number.isInteger(n)) return String(n);
  let s=Number(n.toPrecision(15)).toString(); if (/e/.test(s)) s=n.toPrecision(15).replace(/\.?0+$/,''); return s;
}
function toText(v){ if (isErr(v)) return v; if (v==null) return ''; if (typeof v==='boolean') return v?'TRUE':'FALSE'; if (typeof v==='number') return numberToText(v); return String(v); }
function toNum(v){
  if (isErr(v)) return v; if (v==null) return 0; if (typeof v==='number') return v; if (typeof v==='boolean') return v?1:0;
  const s=String(v).trim().replace(/^£/,'').replace(/,/g,'');
  if (s!=='' && /^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?%?$/.test(s)) return s.endsWith('%') ? parseFloat(s)/100 : parseFloat(s);
  const d=parseDateText(String(v)); if (d!=null) return d;
  return ERR('#VALUE!', `“${v}” is text, so it can't be used as a number`);
}
function toBool(v){ if (isErr(v)) return v; if (typeof v==='boolean') return v; if (v==null) return false; if (typeof v==='number') return v!==0; const s=String(v).toUpperCase(); if (s==='TRUE') return true; if (s==='FALSE') return false; return ERR('#VALUE!', `“${v}” isn't TRUE or FALSE`); }

/* ----- Dates (serial numbers, UK formats) ----- */
const DAY_MS=86400000, EPOCH=Date.UTC(1899,11,30);
const serialOf = (y,m,d) => Math.round((Date.UTC(y,m-1,d)-EPOCH)/DAY_MS);
const fromSerial = s => { const d=new Date(EPOCH+Math.floor(s)*DAY_MS); return {y:d.getUTCFullYear(), m:d.getUTCMonth()+1, d:d.getUTCDate(), dow:d.getUTCDay()}; };
const MONTHS=['january','february','march','april','may','june','july','august','september','october','november','december'];
function parseDateText(s){
  s=String(s).trim(); let m;
  if ((m=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/))) return valid(+m[1],+m[2],+m[3]) + timeFrac(m[4],m[5],m[6]);
  if ((m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/))) return valid(fixYear(+m[3]),+m[2],+m[1]) + timeFrac(m[4],m[5],m[6]);
  if ((m=s.match(/^(\d{1,2})[ -]([A-Za-z]{3,9})[ -](\d{2}|\d{4})$/))){ const k=MONTHS.findIndex(x=>x.startsWith(m[2].toLowerCase())); if (k>=0 && (m[2].length===3||MONTHS[k]===m[2].toLowerCase())) return valid(fixYear(+m[3]),k+1,+m[1]); }
  return null;
  function fixYear(y){ return y<100 ? (y<30?2000+y:1900+y) : y; }
  function valid(y,mo,d){ if (mo<1||mo>12||d<1||d>31) return NaN; const x=serialOf(y,mo,d); return fromSerial(x).d===d ? x : NaN; }
  function timeFrac(h,mi,se){ return h==null?0:((+h)*3600+(+mi)*60+(+(se||0)))/86400; }
}
const pd = s => { const v=parseDateText(s); return v==null||isNaN(v) ? null : v; };
parseDateText.safe = pd;

/* ----- Formatting values for display ----- */
function fmtNum(n, dp){ return n.toLocaleString('en-GB',{minimumFractionDigits:dp,maximumFractionDigits:dp}); }
function displayValue(v, type){
  if (isErr(v)) return v.code;
  if (v==null) return '';
  if (typeof v==='boolean') return v?'TRUE':'FALSE';
  if (typeof v==='number'){
    if (type==='date') return fmtDate(v,'dd/mm/yyyy');
    if (type==='datetime') return fmtDate(v,'dd/mm/yyyy hh:mm');
    if (type==='time') return fmtDate(v,'hh:mm');
    if (type==='month') return fmtDate(v,'mmm yyyy');
    if (type==='money') return '£'+fmtNum(v,2);
    if (type==='money0') return '£'+fmtNum(v,0);
    if (type==='int') return String(Math.round(v));
    if (Number.isInteger(v)) return String(v);
    const a=Math.abs(v); return a>=1e11||a<1e-9 ? v.toPrecision(6) : Number(v.toPrecision(10)).toString();
  }
  return String(v);
}
function fmtDate(s, f){
  const {y,m,d,dow}=fromSerial(s); const frac=s-Math.floor(s); const mins=Math.round(frac*1440); const hh=Math.floor(mins/60)%24, mi=mins%60;
  const D=['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  return f.replace(/yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|nn|MM/gi, tok=>{
    const t=tok.toLowerCase();
    if (t==='yyyy') return String(y); if (t==='yy') return String(y).slice(-2);
    if (t==='mmmm') return MONTHS[m-1][0].toUpperCase()+MONTHS[m-1].slice(1);
    if (t==='mmm') return MONTHS[m-1][0].toUpperCase()+MONTHS[m-1].slice(1,3);
    if (t==='mm') return /h/i.test(f.slice(0,f.toLowerCase().indexOf(tok))) && f.slice(Math.max(0,f.toLowerCase().indexOf(tok)-1), f.toLowerCase().indexOf(tok)) === ':' ? String(mi).padStart(2,'0') : String(m).padStart(2,'0');
    if (t==='m') return String(m);
    if (t==='dddd') return D[dow]; if (t==='ddd') return D[dow].slice(0,3);
    if (t==='dd') return String(d).padStart(2,'0'); if (t==='d') return String(d);
    if (t==='hh') return String(hh).padStart(2,'0'); if (t==='h') return String(hh);
    if (t==='nn') return String(mi).padStart(2,'0'); if (t==='ss') return '00';
    return tok;
  });
}
function textFormat(v, f){
  if (isErr(v)) return v;
  const n=toNum(v); if (isErr(n)) return toText(v);
  const fl=f.toLowerCase();
  if (/[dy]/.test(fl) || /^m+$/.test(fl) || /^mmm/.test(fl) || /h/.test(fl)) return fmtDate(n, f);
  const pct=f.includes('%'); let x=pct?n*100:n;
  const cur=(f.match(/^[£$€]/)||[''])[0];
  const body=f.replace(/^[£$€]/,'').replace('%','');
  const dec=(body.split('.')[1]||'').replace(/[^0#]/g,'').length;
  const comma=body.includes(',');
  const zeros=(body.split('.')[0].match(/0/g)||[]).length;
  let s=Math.abs(x).toFixed(dec);
  let [ip,fp]=s.split('.');
  if (zeros>1) ip=ip.padStart(zeros,'0');
  if (comma) ip=ip.replace(/\B(?=(\d{3})+(?!\d))/g,',');
  return (x<0?'-':'')+cur+ip+(fp?'.'+fp:'')+(pct?'%':'');
}

/* ----- Criteria for COUNTIF, SUMIFS and friends ----- */
function wildRe(s){ return new RegExp('^'+s.replace(/[.+^${}()|[\]\\]/g,'\\$&').replace(/~\*/g,'\u0001').replace(/~\?/g,'\u0002').replace(/\*/g,'.*').replace(/\?/g,'.').replace(/\u0001/g,'\\*').replace(/\u0002/g,'\\?')+'$','i'); }
function makeCriterion(c){
  if (isErr(c)) return ()=>false;
  if (typeof c==='number') return v=>typeof v==='number' ? Math.abs(v-c)<1e-9 : (typeof v==='string' && v.trim()!=='' && !isNaN(+v) && +v===c);
  if (typeof c==='boolean') return v=>v===c;
  if (c==null) return v=>v==null||v==='';
  const m=String(c).match(/^(<>|<=|>=|=|<|>)?([\s\S]*)$/); const op=m[1]||'=', rhs=m[2];
  const rnum = rhs.trim()!=='' && !isNaN(+rhs) ? +rhs : (pd(rhs)!=null ? pd(rhs) : null);
  const rbool = /^(true|false)$/i.test(rhs) ? rhs.toUpperCase()==='TRUE' : null;
  if (op==='=' || op==='<>'){
    let test;
    if (rhs==='') test = v=>v==null||v==='';
    else if (rnum!=null) test = v=>typeof v==='number' ? Math.abs(v-rnum)<1e-9 : (typeof v==='string' && v.trim().toLowerCase()===rhs.trim().toLowerCase());
    else if (rbool!=null) test = v=>v===rbool;
    else { const re=wildRe(rhs); test = v=>typeof v==='string' && re.test(v); }
    return op==='=' ? test : v=>!test(v);
  }
  if (rnum!=null) return v=>{ if (typeof v!=='number') return false; return op==='<'?v<rnum: op==='>'?v>rnum: op==='<='?v<=rnum: v>=rnum; };
  const r=rhs.toLowerCase(); return v=>{ if (typeof v!=='string') return false; const x=v.toLowerCase(); return op==='<'?x<r: op==='>'?x>r: op==='<='?x<=r: x>=r; };
}
function ifsMask(pairs, rows, cols){
  const mask=new Array(rows*cols).fill(true);
  for (const [rng, crit] of pairs){
    if (!isRange(rng)) return ERR('#VALUE!', 'The criteria range must be a range of cells, such as Members!J:J');
    if (rng.rows!==rows && !(rng instanceof Range && rng.whole)) { /* ranges of different sizes */ }
    const test=makeCriterion(crit);
    const a=rng instanceof Range ? rng : toArr(rng);
    const n=Math.min(rows, rng.usedRows ?? rows);
    for (let i=0;i<rows;i++) for (let j=0;j<cols;j++){ const k=i*cols+j; if (!mask[k]) continue; mask[k] = i<n && test(a.at(i,j)); }
  }
  return mask;
}
function sameShape(a, b){ if (!isRange(a)||!isRange(b)) return true; return a.rows===b.rows && a.cols===b.cols; }

/* ----- Comparison (Excel order: numbers < text < booleans; text ignores case) ----- */
function compareVals(a, b){
  const rank=v=>v==null?-1:typeof v==='number'?0:typeof v==='string'?1:2;
  if (a==null) a = typeof b==='string' ? '' : typeof b==='boolean' ? false : 0;
  if (b==null) b = typeof a==='string' ? '' : typeof a==='boolean' ? false : 0;
  const ra=rank(a), rb=rank(b); if (ra!==rb) return ra-rb;
  if (typeof a==='string'){ const x=a.toLowerCase(), y=b.toLowerCase(); return x<y?-1:x>y?1:0; }
  if (typeof a==='boolean') return (a?1:0)-(b?1:0);
  return Math.abs(a-b)<1e-12?0:a<b?-1:1;
}
function lookupEq(target, v){
  if (target==null) return v==null;
  if (typeof target==='string'){ if (typeof v!=='string') return false; return /[*?]/.test(target) ? wildRe(target).test(v) : v.toLowerCase()===target.toLowerCase(); }
  if (typeof target==='number') return typeof v==='number' && Math.abs(v-target)<1e-9;
  return v===target;
}

/* ----- Evaluation ----- */
function evaluate(node, ctx){
  switch (node.t){
    case 'num': case 'str': case 'bool': return node.v;
    case 'errv': return ERR(node.v);
    case 'missing': return null;
    case 'arrc': return new Arr(node.rows);
    case 'ref': {
      const sh=sheetOf(node.sheet, ctx); if (isErr(sh)) return sh;
      const r=node.r+(node.ra?0:ctx.dr), c=node.c+(node.ca?0:ctx.dc);
      if (r<1||c<1) return ERR('#REF!','The reference has moved off the top or left edge of the sheet');
      if (ctx.self && sh===ctx.self.sh && r===ctx.self.r && c===ctx.self.c) return ERR('#REF!','Circular reference: the formula refers to its own cell');
      if (c>sh.cols.length && ctx.extra && sh===ctx.extra.sh && c===ctx.extra.c) return ctx.extra.get(r);
      return cellValue(sh, r, c);
    }
    case 'range': {
      const sh=sheetOf(node.sheet, ctx); if (isErr(sh)) return sh;
      if (node.whole){ const rg=new Range(ctx.wb, sh, 1, node.c1+(node.ca1?0:ctx.dc), 1048576, node.c2+(node.ca2?0:ctx.dc)); rg.whole=true; return rg; }
      const r1=node.r1+(node.ra1?0:ctx.dr), r2=node.r2+(node.ra2?0:ctx.dr), c1=node.c1+(node.ca1?0:ctx.dc), c2=node.c2+(node.ca2?0:ctx.dc);
      if (r1<1||r2<1||c1<1||c2<1) return ERR('#REF!','The range has moved off the edge of the sheet');
      return new Range(ctx.wb, sh, r1, c1, r2, c2);
    }
    case 'neg': return mapUnary(evaluate(node.a, ctx), v=>{ const n=toNum(v); return isErr(n)?n:-n; });
    case 'pct': return mapUnary(evaluate(node.a, ctx), v=>{ const n=toNum(v); return isErr(n)?n:n/100; });
    case 'bin': return binop(node.op, evaluate(node.a, ctx), evaluate(node.b, ctx), ctx);
    case 'fn': {
      const f=FUNCS[node.name];
      if (!f){
        const why = NOT_HERE[node.name] || `There's no function called ${node.name}. Check the spelling.`;
        return ERR('#NAME?', why);
      }
      if (f.lazy) return f.fn(node.args, ctx);
      const args=node.args.map(a=>evaluate(a, ctx));
      if (f.min!=null && args.length<f.min) return ERR('#ERROR', `${node.name} needs at least ${f.min} argument${f.min>1?'s':''}.`);
      if (f.max!=null && args.length>f.max) return ERR('#ERROR', `${node.name} takes at most ${f.max} argument${f.max>1?'s':''}.`);
      return f.fn(args, ctx);
    }
  }
  return ERR('#VALUE!');
}
function sheetOf(name, ctx){
  if (!name) return ctx.sheet;
  const sh=ctx.wb.sheet(name);
  return sh || ERR('#REF!', `There's no sheet called “${name}”. The sheets are ${ctx.wb.sheets.map(s=>s.name).join(', ')}.`);
}
function mapUnary(v, f){ if (isRange(v)){ const a=toArr(v); return new Arr(a.data.map(r=>r.map(f))); } return f(v); }
function binop(op, a, b, ctx){
  if (isRange(a)||isRange(b)){
    const A=isRange(a)?toArr(a):null, B=isRange(b)?toArr(b):null;
    const rows=Math.max(A?A.rows:1, B?B.rows:1), cols=Math.max(A?A.cols:1, B?B.cols:1);
    const get=(X,x,i,j)=>X ? (X.rows===1&&X.cols===1?X.at(0,0): X.at(Math.min(i,X.rows-1), Math.min(j,X.cols-1))) : x;
    const out=[]; for (let i=0;i<rows;i++){ const r=[]; for (let j=0;j<cols;j++) r.push(scalarOp(op, get(A,a,i,j), get(B,b,i,j))); out.push(r); }
    return new Arr(out);
  }
  return scalarOp(op, a, b);
}
function scalarOp(op, a, b){
  if (isErr(a)) return a; if (isErr(b)) return b;
  if (op==='&') return toText(a)+toText(b);
  if (['=','<>','<','>','<=','>='].includes(op)){
    const c=compareVals(a,b);
    return op==='='?c===0: op==='<>'?c!==0: op==='<'?c<0: op==='>'?c>0: op==='<='?c<=0: c>=0;
  }
  const x=toNum(a); if (isErr(x)) return x; const y=toNum(b); if (isErr(y)) return y;
  switch(op){
    case '+': return x+y; case '-': return x-y; case '*': return x*y;
    case '/': return y===0 ? ERR('#DIV/0!','Dividing by zero (or by an empty cell)') : x/y;
    case '^': return Math.pow(x,y);
  }
}
/* A range used where one value is needed: take the cell on the same row (implicit intersection) */
function scalar(v, ctx){
  if (!isRange(v)) return v;
  if (v.rows===1 && v.cols===1) return v.at(0,0);
  if (v instanceof Range && v.cols===1 && ctx && ctx.row!=null && v.sh===ctx.sheet && ctx.row>=v.r1 && ctx.row<=v.r2) return v.at(ctx.row-v.r1, 0);
  return ERR('#VALUE!', 'A range was used where one value is needed');
}

/* ----- Collecting numbers for SUM, AVERAGE and friends ----- */
// Excel: numbers typed straight into the arguments count, text and TRUE in cells are skipped
function nums(args, {countText=false}={}){
  const out=[];
  for (const a of args){
    if (isErr(a)) return a;
    if (isRange(a)){ for (const v of a.values()){ if (isErr(v)) return v; if (typeof v==='number') out.push(v); else if (countText && v!=null) out.push(typeof v==='boolean'?(v?1:0):0); } }
    else if (a==null) continue;
    else { const n=toNum(a); if (isErr(n)) return n; out.push(n); }
  }
  return out;
}
function flat(args){ const out=[]; for (const a of args){ if (isRange(a)) for (const v of a.values()) out.push(v); else out.push(a); } return out; }
const firstErr = arr => arr.find(isErr);
const mean = a => a.reduce((s,x)=>s+x,0)/a.length;
function variance(a, sample){ const m=mean(a); return a.reduce((s,x)=>s+(x-m)**2,0)/(a.length-(sample?1:0)); }
function percentile(a, k){ a=[...a].sort((x,y)=>x-y); const pos=(a.length-1)*k, lo=Math.floor(pos), hi=Math.ceil(pos); return a[lo]+(a[hi]-a[lo])*(pos-lo); }
function pairs(xr, yr){
  if (!isRange(xr)||!isRange(yr)) return ERR('#VALUE!','Both arguments must be ranges');
  const X=[...xr.values()], Y=[...yr.values()];
  if (X.length!==Y.length) return ERR('#N/A','The two ranges must be the same size');
  const xs=[], ys=[]; for (let i=0;i<X.length;i++) if (typeof X[i]==='number' && typeof Y[i]==='number'){ xs.push(X[i]); ys.push(Y[i]); }
  return [xs, ys];
}
function correl(xs, ys){ const mx=mean(xs), my=mean(ys); let sxy=0,sxx=0,syy=0; for (let i=0;i<xs.length;i++){ sxy+=(xs[i]-mx)*(ys[i]-my); sxx+=(xs[i]-mx)**2; syy+=(ys[i]-my)**2; } return sxy/Math.sqrt(sxx*syy); }
const roundTo = (x, d) => { const f=Math.pow(10,d); const y=Math.abs(x)*f; const r=Math.round(y+1e-9*Math.max(1,y)); return Math.sign(x)*r/f; };

/* ----- Function library ----- */
const NOT_HERE = {
  FILTER:'FILTER returns a list of results (it "spills"), which this practice grid can\'t show. Try it in Excel.',
  UNIQUE:'UNIQUE returns a list of results (it "spills"), which this practice grid can\'t show. Try it in Excel.',
  SORT:'SORT returns a list of results (it "spills"), which this practice grid can\'t show. Use Data > Sort in Excel.',
  SORTBY:'SORTBY spills a list of results, which this practice grid can\'t show. Try it in Excel.',
  SEQUENCE:'SEQUENCE spills a list of results, which this practice grid can\'t show. Try it in Excel.',
  LET:'LET isn\'t supported on this site. Write the formula out in full here, or try LET in Excel.',
  LAMBDA:'LAMBDA isn\'t supported on this site.', FREQUENCY:'FREQUENCY spills a list of results, which this grid can\'t show. Try it in Excel.',
  GETPIVOTDATA:'Pivot tables only exist in Excel. Use SUMIFS or COUNTIFS here instead.',
};
const FL = {};
const def = (names, min, max, fn, extra={}) => names.split(" ").forEach(n=>FL[n]={min,max,fn,...extra});
const num1 = (fn) => (args, ctx) => { const v=scalar(args[0],ctx); const n=toNum(v); return isErr(n)?n:fn(n, args, ctx); };

// Maths
def('SUM', 1, 255, a=>{ const n=nums(a); return isErr(n)?n:n.reduce((s,x)=>s+x,0); });
def('PRODUCT', 1, 255, a=>{ const n=nums(a); return isErr(n)?n:n.reduce((s,x)=>s*x,1); });
def('ROUND', 2, 2, (a,c)=>{ const x=toNum(scalar(a[0],c)), d=toNum(scalar(a[1],c)); if (isErr(x)) return x; if (isErr(d)) return d; return roundTo(x, Math.trunc(d)); });
def('ROUNDUP', 2, 2, (a,c)=>{ const x=toNum(scalar(a[0],c)), d=Math.trunc(toNum(scalar(a[1],c))); if (isErr(x)) return x; const f=10**d; return Math.sign(x)*Math.ceil(Math.abs(x)*f-1e-9)/f; });
def('ROUNDDOWN', 2, 2, (a,c)=>{ const x=toNum(scalar(a[0],c)), d=Math.trunc(toNum(scalar(a[1],c))); if (isErr(x)) return x; const f=10**d; return Math.sign(x)*Math.floor(Math.abs(x)*f+1e-9)/f; });
def('INT', 1, 1, num1(x=>Math.floor(x)));
def('TRUNC', 1, 2, (a,c)=>{ const x=toNum(scalar(a[0],c)); if (isErr(x)) return x; const d=a[1]==null?0:toNum(scalar(a[1],c)); const f=10**d; return Math.trunc(x*f)/f; });
def('ABS', 1, 1, num1(x=>Math.abs(x)));
def('SQRT', 1, 1, num1(x=>x<0?ERR('#NUM!','You can\'t take the square root of a negative number'):Math.sqrt(x)));
def('POWER', 2, 2, (a,c)=>scalarOp('^', scalar(a[0],c), scalar(a[1],c)));
def('MOD', 2, 2, (a,c)=>{ const x=toNum(scalar(a[0],c)), y=toNum(scalar(a[1],c)); if (isErr(x)) return x; if (isErr(y)) return y; if (y===0) return ERR('#DIV/0!'); return x-y*Math.floor(x/y); });
def('PI', 0, 0, ()=>Math.PI);
def('SUMPRODUCT', 1, 255, a=>{
  const arrs=a.map(toArr); const r=arrs[0].rows, c=arrs[0].cols;
  if (arrs.some(x=>x.rows!==r||x.cols!==c)) return ERR('#VALUE!','Every range in SUMPRODUCT must be the same size');
  let s=0; for (let i=0;i<r;i++) for (let j=0;j<c;j++){ let p=1; for (const x of arrs){ const v=x.at(i,j); if (isErr(v)) return v; p*= typeof v==='number'?v: typeof v==='boolean'?(v?1:0):0; } s+=p; } return s; });

// Statistics
def('AVERAGE', 1, 255, a=>{ const n=nums(a); if (isErr(n)) return n; return n.length?mean(n):ERR('#DIV/0!','There are no numbers to average'); });
def('AVERAGEA', 1, 255, a=>{ const n=nums(a,{countText:true}); if (isErr(n)) return n; return n.length?mean(n):ERR('#DIV/0!'); });
def('MEDIAN', 1, 255, a=>{ const n=nums(a); if (isErr(n)) return n; if (!n.length) return ERR('#NUM!'); return percentile(n, 0.5); });
def('MODE MODE.SNGL', 1, 255, a=>{ const n=nums(a); if (isErr(n)) return n; const cnt=new Map(); let best=null, bc=1; for (const x of n){ const k=cnt.get(x)+1||1; cnt.set(x,k); }
  for (const x of n){ const k=cnt.get(x); if (k>bc){ bc=k; best=x; } } return best==null?ERR('#N/A','No value appears more than once'):best; });
def('MIN', 1, 255, a=>{ const n=nums(a); return isErr(n)?n: n.length?Math.min(...n):0; });
def('MAX', 1, 255, a=>{ const n=nums(a); return isErr(n)?n: n.length?Math.max(...n):0; });
def('COUNT', 1, 255, a=>{ let k=0; for (const x of a){ if (isRange(x)){ for (const v of x.values()) if (typeof v==='number') k++; } else if (typeof x==='number' || (x!=null && !isErr(toNum(x)) && typeof x!=='string')) k++; else if (typeof x==='string' && !isErr(toNum(x))) k++; } return k; });
def('COUNTA', 1, 255, a=>{ let k=0; for (const x of a){ if (isRange(x)){ for (const v of x.values()) if (v!=null && v!=='') k++; } else if (x!=null) k++; } return k; });
def('COUNTBLANK', 1, 1, a=>{ const r=a[0]; if (!isRange(r)) return ERR('#VALUE!','COUNTBLANK needs a range'); let k=0; const n=r instanceof Range ? Math.min(r.r2,lastRow(r.sh))-r.r1+1 : r.rows; for (let i=0;i<n;i++) for (let j=0;j<r.cols;j++){ const v=r.at(i,j); if (v==null||v==='') k++; } return k; });
def('STDEV STDEV.S', 1, 255, a=>{ const n=nums(a); if (isErr(n)) return n; return n.length<2?ERR('#DIV/0!','Need at least two numbers'):Math.sqrt(variance(n,true)); });
def('STDEVP STDEV.P', 1, 255, a=>{ const n=nums(a); if (isErr(n)) return n; return n.length<1?ERR('#DIV/0!'):Math.sqrt(variance(n,false)); });
def('VAR VAR.S', 1, 255, a=>{ const n=nums(a); if (isErr(n)) return n; return n.length<2?ERR('#DIV/0!'):variance(n,true); });
def('VARP VAR.P', 1, 255, a=>{ const n=nums(a); if (isErr(n)) return n; return n.length<1?ERR('#DIV/0!'):variance(n,false); });
def('GEOMEAN', 1, 255, a=>{ const n=nums(a); if (isErr(n)) return n; if (n.some(x=>x<=0)) return ERR('#NUM!','GEOMEAN only works with positive numbers'); return Math.exp(mean(n.map(Math.log))); });
def('HARMEAN', 1, 255, a=>{ const n=nums(a); if (isErr(n)) return n; if (n.some(x=>x<=0)) return ERR('#NUM!'); return n.length/n.reduce((s,x)=>s+1/x,0); });
def('LARGE', 2, 2, (a,c)=>{ const n=nums([a[0]]); if (isErr(n)) return n; const k=toNum(scalar(a[1],c)); n.sort((x,y)=>y-x); return k<1||k>n.length?ERR('#NUM!'):n[Math.ceil(k)-1]; });
def('SMALL', 2, 2, (a,c)=>{ const n=nums([a[0]]); if (isErr(n)) return n; const k=toNum(scalar(a[1],c)); n.sort((x,y)=>x-y); return k<1||k>n.length?ERR('#NUM!'):n[Math.ceil(k)-1]; });
def('PERCENTILE PERCENTILE.INC', 2, 2, (a,c)=>{ const n=nums([a[0]]); if (isErr(n)) return n; const k=toNum(scalar(a[1],c)); if (k<0||k>1) return ERR('#NUM!','The percentile must be between 0 and 1'); return percentile(n,k); });
def('QUARTILE QUARTILE.INC', 2, 2, (a,c)=>{ const n=nums([a[0]]); if (isErr(n)) return n; const q=Math.trunc(toNum(scalar(a[1],c))); if (q<0||q>4) return ERR('#NUM!','The quartile must be 0 to 4'); return percentile(n,q/4); });
def('CORREL PEARSON', 2, 2, a=>{ const p=pairs(a[0],a[1]); if (isErr(p)) return p; return p[0].length<2?ERR('#DIV/0!'):correl(p[0],p[1]); });
def('RSQ', 2, 2, a=>{ const p=pairs(a[1],a[0]); if (isErr(p)) return p; return correl(p[0],p[1])**2; });
def('SLOPE', 2, 2, a=>{ const p=pairs(a[1],a[0]); if (isErr(p)) return p; const [xs,ys]=p, mx=mean(xs), my=mean(ys); let sxy=0,sxx=0; xs.forEach((x,i)=>{ sxy+=(x-mx)*(ys[i]-my); sxx+=(x-mx)**2; }); return sxy/sxx; });
def('INTERCEPT', 2, 2, a=>{ const p=pairs(a[1],a[0]); if (isErr(p)) return p; const [xs,ys]=p, mx=mean(xs), my=mean(ys); let sxy=0,sxx=0; xs.forEach((x,i)=>{ sxy+=(x-mx)*(ys[i]-my); sxx+=(x-mx)**2; }); return my-sxy/sxx*mx; });
def('SKEW', 1, 255, a=>{ const n=nums(a); if (isErr(n)) return n; const k=n.length; if (k<3) return ERR('#DIV/0!'); const m=mean(n), s=Math.sqrt(variance(n,true)); return k/((k-1)*(k-2))*n.reduce((t,x)=>t+((x-m)/s)**3,0); });
def('STANDARDIZE', 3, 3, (a,c)=>{ const [x,m,s]=a.map(v=>toNum(scalar(v,c))); return firstErr([x,m,s]) || (s<=0?ERR('#NUM!'):(x-m)/s); });
def('RANK RANK.EQ', 2, 3, (a,c)=>{ const x=toNum(scalar(a[0],c)); const n=nums([a[1]]); if (isErr(x)) return x; if (isErr(n)) return n; const asc=a[2]!=null&&toNum(a[2])!==0; if (!n.some(v=>Math.abs(v-x)<1e-9)) return ERR('#N/A','The number isn\'t in the list'); return 1+n.filter(v=> asc? v<x-1e-9 : v>x+1e-9).length; });

// Conditional aggregates
function ifsCore(name, sumRange, pairList, ctx, reducer){
  for (let i=0;i<pairList.length;i++) if (!isRange(pairList[i][0])) return ERR('#VALUE!', `${name}: each criteria range must be a range of cells (for example Members!H:H)`);
  const base = sumRange || pairList[0][0];
  const rows = Math.max(base.usedRows ?? base.rows, ...pairList.map(p=>p[0].usedRows ?? p[0].rows)), cols=base.cols;
  for (const [r] of pairList) if (!(r instanceof Range && r.whole) && !(base instanceof Range && base.whole) && (r.rows!==base.rows||r.cols!==base.cols)) return ERR('#VALUE!', `${name}: every range must be the same size. Whole columns, like Visits!C:C, avoid this.`);
  const mask=ifsMask(pairList.map(([r,c])=>[r, scalar(c,ctx)]), rows, cols); if (isErr(mask)) return mask;
  const vals=[]; for (let i=0;i<rows;i++) for (let j=0;j<cols;j++) if (mask[i*cols+j]){ const v=base.at(i,j); if (isErr(v)) return v; vals.push(v); }
  return reducer(vals);
}
const numsOf = vals => vals.filter(v=>typeof v==='number');
def('COUNTIF', 2, 2, (a,c)=>ifsCore('COUNTIF', null, [[a[0],a[1]]], c, v=>v.length));
def('COUNTIFS', 2, 254, (a,c)=>{ if (a.length%2) return ERR('#ERROR','COUNTIFS takes pairs: a range, then its criteria.'); const p=[]; for (let i=0;i<a.length;i+=2) p.push([a[i],a[i+1]]); return ifsCore('COUNTIFS', null, p, c, v=>v.length); });
def('SUMIF', 2, 3, (a,c)=>ifsCore('SUMIF', a[2]||a[0], [[a[0],a[1]]], c, v=>numsOf(v).reduce((s,x)=>s+x,0)));
def('SUMIFS', 3, 255, (a,c)=>{ if (a.length%2===0) return ERR('#ERROR','SUMIFS takes the range to add up, then pairs of a range and its criteria.'); const p=[]; for (let i=1;i<a.length;i+=2) p.push([a[i],a[i+1]]); return ifsCore('SUMIFS', a[0], p, c, v=>numsOf(v).reduce((s,x)=>s+x,0)); });
def('AVERAGEIF', 2, 3, (a,c)=>ifsCore('AVERAGEIF', a[2]||a[0], [[a[0],a[1]]], c, v=>{ const n=numsOf(v); return n.length?mean(n):ERR('#DIV/0!','No rows match the criteria'); }));
def('AVERAGEIFS', 3, 255, (a,c)=>{ if (a.length%2===0) return ERR('#ERROR','AVERAGEIFS takes the range to average, then pairs of a range and its criteria.'); const p=[]; for (let i=1;i<a.length;i+=2) p.push([a[i],a[i+1]]); return ifsCore('AVERAGEIFS', a[0], p, c, v=>{ const n=numsOf(v); return n.length?mean(n):ERR('#DIV/0!','No rows match the criteria'); }); });
def('MAXIFS', 3, 255, (a,c)=>{ const p=[]; for (let i=1;i<a.length;i+=2) p.push([a[i],a[i+1]]); return ifsCore('MAXIFS', a[0], p, c, v=>{ const n=numsOf(v); return n.length?Math.max(...n):0; }); });
def('MINIFS', 3, 255, (a,c)=>{ const p=[]; for (let i=1;i<a.length;i+=2) p.push([a[i],a[i+1]]); return ifsCore('MINIFS', a[0], p, c, v=>{ const n=numsOf(v); return n.length?Math.min(...n):0; }); });
/* A range as the criteria (=COUNTIF(A2:A52,A2:A52)) gives one result per cell, as in Excel 365 */
function vecCrit(name, critIdx){
  const f=FL[name], inner=f.fn;
  f.fn=(a,c)=>{
    const k=critIdx(a).find(i=>isRange(a[i]) && (a[i].rows>1||a[i].cols>1) && !(a[i] instanceof Range && a[i].whole));
    if (k==null) return inner(a,c);
    const A=toArr(a[k]); return new Arr(A.data.map(row=>row.map(v=>{ const b=[...a]; b[k]=v; return inner(b,c); })));
  };
}
vecCrit('COUNTIF', ()=>[1]); vecCrit('SUMIF', ()=>[1]); vecCrit('AVERAGEIF', ()=>[1]);
vecCrit('COUNTIFS', a=>a.map((_,i)=>i).filter(i=>i%2===1));
['SUMIFS','AVERAGEIFS','MAXIFS','MINIFS'].forEach(n=>vecCrit(n, a=>a.map((_,i)=>i).filter(i=>i>=2&&i%2===0)));

// Logical (IF and friends only evaluate the branch they need, like Excel)
def('IF', 1, 3, (args, ctx)=>{
  const t=scalar(evaluate(args[0],ctx),ctx); if (isErr(t)) return t; const b=toBool(t); if (isErr(b)) return b;
  if (b) return args[1] ? scalar(evaluate(args[1],ctx),ctx) ?? 0 : true;
  return args[2] ? (args[2].t==='missing'?0:scalar(evaluate(args[2],ctx),ctx) ?? 0) : false; }, {lazy:true});
def('IFS', 2, 254, (args, ctx)=>{ for (let i=0;i<args.length;i+=2){ const t=toBool(scalar(evaluate(args[i],ctx),ctx)); if (isErr(t)) return t; if (t) return scalar(evaluate(args[i+1],ctx),ctx) ?? 0; } return ERR('#N/A','None of the IFS conditions is TRUE. End with TRUE, "something" as a catch-all.'); }, {lazy:true});
def('IFERROR', 2, 2, (args, ctx)=>{ const v=scalar(evaluate(args[0],ctx),ctx); return isErr(v) ? (scalar(evaluate(args[1],ctx),ctx) ?? 0) : (v ?? 0); }, {lazy:true});
def('IFNA', 2, 2, (args, ctx)=>{ const v=scalar(evaluate(args[0],ctx),ctx); return isErr(v)&&v.code==='#N/A' ? (scalar(evaluate(args[1],ctx),ctx) ?? 0) : v; }, {lazy:true});
def('SWITCH', 3, 254, (args, ctx)=>{ const x=scalar(evaluate(args[0],ctx),ctx); if (isErr(x)) return x; let i=1; for (; i+1<args.length; i+=2){ if (compareVals(x, scalar(evaluate(args[i],ctx),ctx))===0) return scalar(evaluate(args[i+1],ctx),ctx); } return i<args.length ? scalar(evaluate(args[i],ctx),ctx) : ERR('#N/A','No value matched, and there\'s no default'); }, {lazy:true});
function logicArgs(a){ const out=[]; for (const x of a){ if (isRange(x)){ for (const v of x.values()){ if (isErr(v)) return v; if (typeof v==='boolean') out.push(v); else if (typeof v==='number') out.push(v!==0); } } else { const b=toBool(x); if (isErr(b)) return b; out.push(b); } } return out; }
def('AND', 1, 255, a=>{ const b=logicArgs(a); return isErr(b)?b:b.every(Boolean); });
def('OR', 1, 255, a=>{ const b=logicArgs(a); return isErr(b)?b:b.some(Boolean); });
def('XOR', 1, 255, a=>{ const b=logicArgs(a); return isErr(b)?b:b.filter(Boolean).length%2===1; });
def('NOT', 1, 1, (a,c)=>{ const b=toBool(scalar(a[0],c)); return isErr(b)?b:!b; });
def('TRUE', 0, 0, ()=>true); def('FALSE', 0, 0, ()=>false);
def('NA', 0, 0, ()=>ERR('#N/A'));
def('ISBLANK', 1, 1, (a,c)=>{ const v=scalar(a[0],c); return v==null; });
def('ISNUMBER', 1, 1, (a,c)=>typeof scalar(a[0],c)==='number');
def('ISTEXT', 1, 1, (a,c)=>typeof scalar(a[0],c)==='string');
def('ISERROR', 1, 1, (a,c)=>isErr(scalar(a[0],c)));
def('ISNA', 1, 1, (a,c)=>{ const v=scalar(a[0],c); return isErr(v)&&v.code==='#N/A'; });
def('ISLOGICAL', 1, 1, (a,c)=>typeof scalar(a[0],c)==='boolean');

// Lookups
def('VLOOKUP', 3, 4, (a,c)=>{
  const x=scalar(a[0],c), tbl=a[1]; if (isErr(x)) return x;
  if (!isRange(tbl)) return ERR('#VALUE!','The second argument must be the table to search, such as Plans!A:C');
  const k=Math.trunc(toNum(scalar(a[2],c))); if (isErr(k)) return k;
  if (k<1) return ERR('#VALUE!','The column number must be 1 or more');
  if (k>tbl.cols) return ERR('#REF!', `Column ${k} is outside the table, which is only ${tbl.cols} column${tbl.cols>1?'s':''} wide`);
  const approx = a[3]==null ? true : toBool(scalar(a[3],c));
  const n=tbl.usedRows;
  if (!approx){ for (let i=0;i<n;i++) if (lookupEq(x, tbl.at(i,0))) return tbl.at(i,k-1) ?? 0; return ERR('#N/A', `${toText(x)} wasn't found in the first column of the table`); }
  let hit=-1; for (let i=0;i<n;i++){ const v=tbl.at(i,0); if (v==null) continue; if (typeof v!==typeof x) continue; if (compareVals(v,x)<=0) hit=i; else break; }
  return hit<0 ? ERR('#N/A','No value is smaller than or equal to the one you looked up') : (tbl.at(hit,k-1) ?? 0);
});
def('HLOOKUP', 3, 4, (a,c)=>{
  const x=scalar(a[0],c), tbl=a[1]; if (isErr(x)) return x;
  if (!isRange(tbl)) return ERR('#VALUE!','The second argument must be the table to search, such as PriceHistory!A1:E6');
  const k=Math.trunc(toNum(scalar(a[2],c))); if (k<1) return ERR('#VALUE!');
  const rows = tbl.usedRows; if (k>rows) return ERR('#REF!', `Row ${k} is outside the table, which is only ${rows} rows tall`);
  const approx = a[3]==null ? true : toBool(scalar(a[3],c));
  if (!approx){ for (let j=0;j<tbl.cols;j++) if (lookupEq(x, tbl.at(0,j))) return tbl.at(k-1,j) ?? 0; return ERR('#N/A', `${toText(x)} wasn't found in the first row of the table`); }
  let hit=-1; for (let j=0;j<tbl.cols;j++){ const v=tbl.at(0,j); if (v==null||typeof v!==typeof x) continue; if (compareVals(v,x)<=0) hit=j; else break; }
  return hit<0 ? ERR('#N/A') : (tbl.at(k-1,hit) ?? 0);
});
def('XLOOKUP', 3, 6, (a,c)=>{
  const x=scalar(a[0],c), look=a[1], ret=a[2]; if (isErr(x)) return x;
  if (!isRange(look)||!isRange(ret)) return ERR('#VALUE!','XLOOKUP needs a range to search and a range to return from');
  const mode = a[4]==null?0:toNum(scalar(a[4],c)), search = a[5]==null?1:toNum(scalar(a[5],c));
  const vertical = look.cols===1;
  const n = vertical ? look.usedRows : look.cols;
  if (vertical && !(ret.usedRows>=n || (ret instanceof Range && ret.whole))) return ERR('#VALUE!','The search range and the return range must be the same height');
  const val = i => vertical ? look.at(i,0) : look.at(0,i);
  const order = search===-1 ? [...Array(n).keys()].reverse() : [...Array(n).keys()];
  let hit=-1;
  if (mode===0 || mode===2){ for (const i of order) if (mode===2 ? lookupEq(x, val(i)) : (typeof x==='string' ? (typeof val(i)==='string' && val(i).toLowerCase()===x.toLowerCase()) : lookupEq(x,val(i)))){ hit=i; break; } }
  else { let best=null; for (const i of order){ const v=val(i); if (v==null||typeof v!==typeof x) continue; const cmp=compareVals(v,x); if (cmp===0){ hit=i; break; } if (mode===-1 && cmp<0 && (best==null||compareVals(v,val(best))>0)) best=i; if (mode===1 && cmp>0 && (best==null||compareVals(v,val(best))<0)) best=i; } if (hit<0 && best!=null) hit=best; }
  if (hit<0) return a[3]!=null && a[3]!==undefined && !(a[3]===null) ? scalar(a[3],c) : ERR('#N/A', `${toText(x)} wasn't found`);
  if (vertical){ if (ret.cols===1) return ret.at(hit,0) ?? 0; const row=[]; for (let j=0;j<ret.cols;j++) row.push(ret.at(hit,j)); return new Arr([row]); }
  if (ret.rows===1) return ret.at(0,hit) ?? 0; const col=[]; for (let i=0;i<ret.rows;i++) col.push([ret.at(i,hit)]); return new Arr(col);
});
def('MATCH', 2, 3, (a,c)=>{
  const x=scalar(a[0],c), rg=a[1]; if (isErr(x)) return x; if (!isRange(rg)) return ERR('#N/A');
  const type = a[2]==null?1:toNum(scalar(a[2],c)); const vertical=rg.cols===1; const n=vertical?rg.usedRows:rg.cols;
  const val=i=>vertical?rg.at(i,0):rg.at(0,i);
  if (type===0){ for (let i=0;i<n;i++) if (lookupEq(x,val(i))) return i+1; return ERR('#N/A', `${toText(x)} wasn't found`); }
  let hit=-1; for (let i=0;i<n;i++){ const v=val(i); if (v==null||typeof v!==typeof x) continue; const cmp=compareVals(v,x); if (type>0 ? cmp<=0 : cmp>=0) hit=i; else break; }
  return hit<0?ERR('#N/A'):hit+1;
});
def('INDEX', 2, 3, (a,c)=>{
  const rg=a[0]; if (!isRange(rg)) return ERR('#VALUE!','INDEX needs a range first');
  let r=Math.trunc(toNum(scalar(a[1],c))), k=a[2]==null?0:Math.trunc(toNum(scalar(a[2],c)));
  if (isErr(r)) return r; if (isErr(k)) return k;
  if (rg.cols===1 && k===0) k=1; if (rg.rows===1 && a[2]==null){ k=r; r=1; }
  if (r<1||k<1) return ERR('#VALUE!','INDEX needs a row and column number of 1 or more');
  if (r>(rg instanceof Range && rg.whole ? 1048576 : rg.rows) || k>rg.cols) return ERR('#REF!','The row or column number is outside the range');
  return rg.at(r-1,k-1) ?? 0;
});
def('ROW', 0, 1, (a,c)=>a.length&&a[0] instanceof Range ? a[0].r1 : (c.row||1));
def('COLUMN', 0, 1, (a,c)=>a.length&&a[0] instanceof Range ? a[0].c1 : (c.col||1));
def('ROWS', 1, 1, a=>isRange(a[0])?(a[0] instanceof Range&&a[0].whole?1048576:a[0].rows):1);

// Text
const txt = (v,c) => { const s=scalar(v,c); return isErr(s)?s:toText(s); };
def('LEFT', 1, 2, (a,c)=>{ const s=txt(a[0],c); if (isErr(s)) return s; const n=a[1]==null?1:toNum(scalar(a[1],c)); return n<0?ERR('#VALUE!'):s.slice(0,n); });
def('RIGHT', 1, 2, (a,c)=>{ const s=txt(a[0],c); if (isErr(s)) return s; const n=a[1]==null?1:toNum(scalar(a[1],c)); return n<0?ERR('#VALUE!'):(n===0?'':s.slice(-n)); });
def('MID', 3, 3, (a,c)=>{ const s=txt(a[0],c); if (isErr(s)) return s; const st=toNum(scalar(a[1],c)), n=toNum(scalar(a[2],c)); if (st<1||n<0) return ERR('#VALUE!'); return s.substr(st-1,n); });
def('LEN', 1, 1, (a,c)=>{ const s=txt(a[0],c); return isErr(s)?s:s.length; });
def('TRIM', 1, 1, (a,c)=>{ const s=txt(a[0],c); return isErr(s)?s:s.replace(/ +/g,' ').trim(); });
def('CLEAN', 1, 1, (a,c)=>{ const s=txt(a[0],c); return isErr(s)?s:s.replace(/[\x00-\x1f]/g,''); });
def('UPPER', 1, 1, (a,c)=>{ const s=txt(a[0],c); return isErr(s)?s:s.toUpperCase(); });
def('LOWER', 1, 1, (a,c)=>{ const s=txt(a[0],c); return isErr(s)?s:s.toLowerCase(); });
def('PROPER', 1, 1, (a,c)=>{ const s=txt(a[0],c); return isErr(s)?s:s.toLowerCase().replace(/(^|[^a-z])([a-z])/g,(m,p,ch)=>p+ch.toUpperCase()); });
def('CONCAT', 1, 255, a=>{ let s=''; for (const x of a){ if (isRange(x)){ for (const v of x.values()){ if (isErr(v)) return v; s+=toText(v); } } else { if (isErr(x)) return x; s+=toText(x); } } return s; });
def('CONCATENATE', 1, 255, (a,c)=>{ let s=''; for (const x of a){ const v=scalar(x,c); if (isErr(v)) return v; s+=toText(v); } return s; });
def('TEXTJOIN', 3, 255, (a,c)=>{ const d=txt(a[0],c), ign=toBool(scalar(a[1],c)); const parts=[]; for (const v of flat(a.slice(2))){ if (isErr(v)) return v; const t=toText(v); if (ign && t==='') continue; parts.push(t); } return parts.join(d); });
def('SUBSTITUTE', 3, 4, (a,c)=>{ const s=txt(a[0],c), o=txt(a[1],c), n=txt(a[2],c); if (isErr(s)) return s; if (!o) return s; if (a[3]==null) return s.split(o).join(n); const k=toNum(scalar(a[3],c)); let i=-1; for (let j=0;j<k;j++){ i=s.indexOf(o,i+1); if (i<0) return s; } return s.slice(0,i)+n+s.slice(i+o.length); });
def('REPLACE', 4, 4, (a,c)=>{ const s=txt(a[0],c); const st=toNum(scalar(a[1],c)), n=toNum(scalar(a[2],c)), r=txt(a[3],c); return s.slice(0,st-1)+r+s.slice(st-1+n); });
def('FIND', 2, 3, (a,c)=>{ const f=txt(a[0],c), s=txt(a[1],c); const st=a[2]==null?1:toNum(scalar(a[2],c)); const i=s.indexOf(f, st-1); return i<0?ERR('#VALUE!',`“${f}” wasn't found (FIND is case-sensitive)`):i+1; });
def('SEARCH', 2, 3, (a,c)=>{ const f=txt(a[0],c), s=txt(a[1],c); const st=a[2]==null?1:toNum(scalar(a[2],c)); const i=s.toLowerCase().indexOf(f.toLowerCase(), st-1); return i<0?ERR('#VALUE!',`“${f}” wasn't found`):i+1; });
def('EXACT', 2, 2, (a,c)=>txt(a[0],c)===txt(a[1],c));
def('REPT', 2, 2, (a,c)=>{ const s=txt(a[0],c); const n=toNum(scalar(a[1],c)); return s.repeat(Math.max(0,Math.trunc(n))); });
def('VALUE', 1, 1, (a,c)=>{ const v=scalar(a[0],c); if (typeof v==='number') return v; const n=toNum(v); return isErr(n)?ERR('#VALUE!',`“${v}” can't be turned into a number`):n; });
def('TEXT', 2, 2, (a,c)=>{ const v=scalar(a[0],c), f=txt(a[1],c); if (isErr(v)) return v; return textFormat(v, f); });

// Dates
def('DATE', 3, 3, (a,c)=>{ const [y,m,d]=a.map(v=>toNum(scalar(v,c))); const e=firstErr([y,m,d]); if (e) return e; return Math.round((Date.UTC(y,m-1,d)-EPOCH)/DAY_MS); });
const dpart = fn => (a,c)=>{ const v=scalar(a[0],c); if (isErr(v)) return v; if (typeof v==='string'){ const n=pd(v); if (n==null) return ERR('#VALUE!',`“${v}” isn't a date Excel recognises`); return fn(n); } const n=toNum(v); return isErr(n)?n:fn(n); };
def('YEAR', 1, 1, dpart(n=>fromSerial(n).y));
def('MONTH', 1, 1, dpart(n=>fromSerial(n).m));
def('DAY', 1, 1, dpart(n=>fromSerial(n).d));
def('HOUR', 1, 1, dpart(n=>Math.floor(Math.round((n-Math.floor(n))*1440)/60)%24));
def('MINUTE', 1, 1, dpart(n=>Math.round((n-Math.floor(n))*1440)%60));
def('WEEKDAY', 1, 2, (a,c)=>{ const n=toNum(scalar(a[0],c)); if (isErr(n)) return n; const t=a[1]==null?1:toNum(scalar(a[1],c)); const dow=fromSerial(n).dow; if (t===2) return dow===0?7:dow; if (t===3) return dow===0?6:dow-1; return dow+1; });
def('TIME', 3, 3, (a,c)=>{ const [h,m,s]=a.map(v=>toNum(scalar(v,c))); return ((h*3600+m*60+s)/86400)%1; });
def('DATEVALUE', 1, 1, (a,c)=>{ const s=scalar(a[0],c); if (typeof s==='number') return ERR('#VALUE!','DATEVALUE needs text. This cell already holds a real date (a number).'); const n=pd(String(s)); return n==null?ERR('#VALUE!',`“${s}” isn't a date Excel recognises`):Math.floor(n); });
def('TODAY', 0, 0, ()=>Math.floor((Date.now()-EPOCH)/DAY_MS));
def('NOW', 0, 0, ()=>(Date.now()-EPOCH)/DAY_MS);
def('EDATE', 2, 2, (a,c)=>{ const n=toNum(scalar(a[0],c)), k=toNum(scalar(a[1],c)); const {y,m,d}=fromSerial(n); const t=new Date(Date.UTC(y,m-1+k,1)); const last=new Date(Date.UTC(t.getUTCFullYear(),t.getUTCMonth()+1,0)).getUTCDate(); return serialOf(t.getUTCFullYear(), t.getUTCMonth()+1, Math.min(d,last)); });
def('EOMONTH', 2, 2, (a,c)=>{ const n=toNum(scalar(a[0],c)), k=toNum(scalar(a[1],c)); const {y,m}=fromSerial(n); return Math.round((Date.UTC(y,m+k,0)-EPOCH)/DAY_MS); });
def('DATEDIF', 3, 3, (a,c)=>{
  const s=toNum(scalar(a[0],c)), e=toNum(scalar(a[1],c)), u=txt(a[2],c).toUpperCase(); if (isErr(s)) return s; if (isErr(e)) return e;
  if (s>e) return ERR('#NUM!','The start date is after the end date');
  const A=fromSerial(s), B=fromSerial(e);
  let months=(B.y-A.y)*12+(B.m-A.m); if (B.d<A.d) months--;
  if (u==='Y') return Math.floor(months/12); if (u==='M') return months; if (u==='D') return Math.floor(e)-Math.floor(s);
  if (u==='YM') return months%12;
  if (u==='MD'){ let d=B.d-A.d; if (d<0){ const pl=new Date(Date.UTC(B.y,B.m-1,0)).getUTCDate(); d+=pl; } return d; }
  if (u==='YD'){ let st=serialOf(B.y,A.m,A.d); if (st>e) st=serialOf(B.y-1,A.m,A.d); return Math.floor(e)-st; }
  return ERR('#NUM!','The unit must be "Y", "M", "D", "YM", "MD" or "YD"'); });
def('NETWORKDAYS', 2, 2, (a,c)=>{ let s=Math.floor(toNum(scalar(a[0],c))), e=Math.floor(toNum(scalar(a[1],c))); let sign=1; if (s>e){ [s,e]=[e,s]; sign=-1; } let k=0; for (let d=s; d<=e; d++){ const w=fromSerial(d).dow; if (w!==0&&w!==6) k++; } return sign*k; });
def('YEARFRAC', 2, 3, (a,c)=>{ const s=toNum(scalar(a[0],c)), e=toNum(scalar(a[1],c)); return Math.abs(e-s)/365.25; });
const FUNCS = FL;

/* ----- Public API ----- */
// target: { sheet: sheet object, row, col } where the formula lives. dr/dc shift relative references when a formula is filled down.
function runFormula(wb, formula, target, {dr=0, dc=0, extra=null}={}){
  let ast; try { ast=parse(formula); } catch(e){ if (isErr(e)) return {error:e}; throw e; }
  const ctx={wb, sheet:target.sheet, row:target.row+dr, col:target.col+dc, dr, dc, extra, self:{sh:target.sheet, r:target.row+dr, c:target.col+dc}};
  let v;
  try { v=evaluate(ast, ctx); } catch(e){ if (isErr(e)) return {error:e}; if (e instanceof RangeError) return {error:ERR('#NUM!','The calculation is too big')}; throw e; }
  if (isRange(v)){
    if (v.rows===1 && v.cols===1) v=v.at(0,0);
    else { const sv=scalar(v, ctx); if (!isErr(sv) && v instanceof Range) v=sv; else return {value:ERR('#SPILL', `This formula returns ${v instanceof Range?v.usedRows:v.rows} × ${v.cols} values (in Excel it would “spill” into the cells around it). Wrap it in a function such as SUM or COUNTA to get one answer.`)}; }
  }
  if (v==null) v=0;
  if (typeof v==='number' && !isFinite(v)) v=ERR('#NUM!');
  return {value:v, ast};
}
// Cell references used in a formula, for highlighting and for "did they use the right sheet" hints
function refsIn(formula){ try { const out=[]; const walk=n=>{ if (!n) return; if (n.t==='ref'||n.t==='range') out.push(n); (n.args||[]).forEach(walk); walk(n.a); walk(n.b); }; walk(parse(formula)); return out; } catch(e){ return []; } }
const fnNamesIn = formula => { try { const out=new Set(); const walk=n=>{ if (!n) return; if (n.t==='fn') out.add(n.name); (n.args||[]).forEach(walk); walk(n.a); walk(n.b); }; walk(parse(formula)); return [...out]; } catch(e){ return []; } };

/* Compares a learner's value with the expected one. Numbers: within 0.005, or the same once rounded to the learner's own
   precision (at least 1 decimal place for a non-whole answer). Text: ignores case and spaces at the ends. */
function sameValue(got, want, {typed=false, dp=null}={}){
  if (isErr(want)) return isErr(got) && got.code===want.code;
  if (isErr(got)) return false;
  if (typeof want==='number'){
    let g=got; if (typeof g==='boolean') return false; if (typeof g==='string'){ const n=toNum(g); if (isErr(n)) return false; g=n; }
    if (Math.abs(g-want) <= 0.005 + 1e-9*Math.abs(want)) return true;
    if (Number.isInteger(want)) return false;
    const d = dp!=null ? dp : (()=>{ const s=String(got); const m=s.match(/\.(\d+)/); return m?m[1].length:0; })();
    if (d>=1 && roundTo(want, d) === roundTo(g, d) && typed) return 'rounded';
    return false;
  }
  if (typeof want==='boolean') return typeof got==='boolean' ? got===want : String(got).trim().toUpperCase()===(want?'TRUE':'FALSE');
  return String(got ?? '').trim().toLowerCase() === String(want ?? '').trim().toLowerCase();
}
// Reads what a learner typed in the "In Excel" box: numbers with £, commas or %, UK dates, TRUE/FALSE or text
function readTyped(s, wantType){
  const t=String(s).trim(); if (t==='') return {value:null, dp:0};
  if (/^(true|false)$/i.test(t)) return {value:t.toUpperCase()==='TRUE'};
  if (wantType==='date' || wantType==='month'){ const d=pd(t); if (d!=null) return {value:Math.floor(d)}; }
  const clean=t.replace(/^£/,'').replace(/,/g,'').replace(/\s/g,'');
  if (/^[-+]?(\d+\.?\d*|\.\d+)%$/.test(clean)){ const raw=clean.slice(0,-1); const dp=(raw.split('.')[1]||'').length+2; return {value:parseFloat(raw)/100, dp}; }
  if (/^[-+]?(\d+\.?\d*|\.\d+)$/.test(clean)) return {value:parseFloat(clean), dp:(clean.split('.')[1]||'').length};
  return {value:t};
}

if (typeof module!=='undefined') module.exports = { makeWorkbook, runFormula, parse, tokenize, displayValue, cellValue, colLetter, colNumber, sameValue, readTyped, refsIn, fnNamesIn, isErr, XErr, FUNCS, fmtDate, serialOf, fromSerial, parseDateText, lastRow, textFormat };
