/* ---------- Helpers for writing lesson content ---------- */
const esc = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const TIP = t => `<div class="tip">${t}</div>`;
const NOTE = (title, t) => `<div class="focus-note"><b>${title}</b> ${t}</div>`;
/* A formula example. The app adds a "Try this" button that puts it in the practice grid beside the lesson.
   F('=AVERAGE(E:E)', 'Visits') runs once; pass {fill:true} to fill it down a new column, and {col:'Age'} to name it. */
const F = (formula, sheet, o={}) => `<div class="fx-ex" data-formula="${esc(formula)}" data-sheet="${sheet}"${o.fill?' data-fill="1"':''}${o.col?` data-col="${esc(o.col)}"`:''}>${o.note?`<span class="fx-note">${o.note}</span>`:''}</div>`;
/* Tabs inside a lesson: TABS([{label, html}]) */
let TABS_N = 0;
const TABS = tabs => { const n=++TABS_N;
  return `<div class="ltabs-box" data-ltabs><div class="ltabs-list" role="tablist">${tabs.map((t,i)=>`<button type="button" role="tab" id="lt${n}-${i}" aria-controls="lp${n}-${i}" aria-selected="${i===0}" tabindex="${i===0?0:-1}">${t.label}</button>`).join('')}</div>${tabs.map((t,i)=>`<div class="ltabs-panel" role="tabpanel" id="lp${n}-${i}" aria-labelledby="lt${n}-${i}"${i===0?'':' hidden'}>${t.html}</div>`).join('')}</div>`; };
/* An interactive piece drawn by the app (see WIDGETS in app.js) */
const W = (name, attrs={}) => `<div class="widget" data-widget="${name}"${Object.entries(attrs).map(([k,v])=>` data-${k}="${esc(v)}"`).join('')}></div>`;
const STEPS = items => `<ol class="steps">${items.map(s=>`<li>${s}</li>`).join('')}</ol>`;
const KEYS = (...k) => k.map(x=>`<kbd>${x}</kbd>`).join('+');
const DL_WORKBOOK = `<a class="btn small" href="downloads/sparta-leisure.xlsx" download data-testid="download-workbook">Download the workbook</a>`;

/* ----- Static figures ----- */
// Normal distribution with the 68 / 95 / 99.7 rule
const FIG_NORMAL = (() => {
  const W=560, H=230, base=190, cx=W/2, sd=70, pdf=x=>Math.exp(-0.5*((x-cx)/sd)**2);
  const pts=[]; for (let x=20;x<=W-20;x+=4) pts.push([x, base-150*pdf(x)]);
  const area=(k)=>{ const a=cx-k*sd, b=cx+k*sd; const p=pts.filter(([x])=>x>=a&&x<=b); return `M${a} ${base} `+p.map(([x,y])=>`L${x} ${y.toFixed(1)}`).join(' ')+` L${b} ${base} Z`; };
  const labels=[-3,-2,-1,0,1,2,3].map(k=>`<text x="${cx+k*sd}" y="${base+18}" text-anchor="middle" class="ax">${k===0?'mean':(k>0?'+':'−')+Math.abs(k)+' SD'}</text>`).join('');
  return `<figure class="figure"><svg viewBox="0 0 ${W} ${H}" role="img" aria-label="A bell-shaped normal distribution. 68% of values lie within 1 standard deviation of the mean, 95% within 2, and 99.7% within 3.">
<path d="${area(3)}" class="nd3"/><path d="${area(2)}" class="nd2"/><path d="${area(1)}" class="nd1"/>
<path d="M${pts.map(([x,y])=>`${x} ${y.toFixed(1)}`).join(' L')}" class="nd-line"/>
<line x1="20" y1="${base}" x2="${W-20}" y2="${base}" class="axis"/>${labels}
<text x="${cx}" y="${base-40}" text-anchor="middle" class="nd-t">68%</text><text x="${cx-1.5*sd}" y="${base-14}" text-anchor="middle" class="nd-t sm">95%</text><text x="${cx+2.55*sd}" y="${base-6}" text-anchor="middle" class="nd-t sm">99.7%</text>
</svg><figcaption>In a normal distribution, 68% of values lie within 1 standard deviation (SD) of the mean, 95% within 2 SDs and 99.7% within 3.</figcaption></figure>`;
})();
// Positive and negative skew side by side
const FIG_SKEW = (() => {
  const curve=(skew)=>{ const pts=[]; for (let i=0;i<=100;i++){ const t=i/100; const x=skew>0?t:1-t; const y=Math.pow(x,1.6)*Math.pow(1-x,4.5); pts.push([20+t*220, y]); } const m=Math.max(...pts.map(p=>p[1])); return 'M'+pts.map(([x,y])=>`${x.toFixed(1)} ${(150-120*y/m).toFixed(1)}`).join(' L'); };
  const panel=(dx, skew, title, order)=>`<g transform="translate(${dx},0)"><text x="130" y="18" text-anchor="middle" class="sk-h">${title}</text><path d="${curve(skew)}" class="nd-line"/><line x1="20" y1="150" x2="240" y2="150" class="axis"/><text x="130" y="172" text-anchor="middle" class="ax">${order}</text></g>`;
  return `<figure class="figure"><svg viewBox="0 0 560 182" role="img" aria-label="Positive skew has its long tail to the right; negative skew has its long tail to the left.">
${panel(10, 1, 'Positive skew: tail to the right', 'Mode < Median < Mean (usually)')}${panel(290, -1, 'Negative skew: tail to the left', 'Mean < Median < Mode (usually)')}</svg>
<figcaption>The tail points the way of the skew. Visit lengths at Sparta Leisure are positively skewed: most visits are about an hour, and a few are much longer.</figcaption></figure>`;
})();
// ETL and ELT pipelines
const FIG_ETL = (() => {
  const box=(x,y,w,t,cls='')=>`<g class="pipe ${cls}"><rect x="${x}" y="${y}" width="${w}" height="40" rx="9"/><text x="${x+w/2}" y="${y+25}" text-anchor="middle">${t}</text></g>`;
  const arrow=(x1,y,x2)=>`<path d="M${x1} ${y} H${x2-8}" class="pipe-arrow" marker-end="url(#arr)"/>`;
  const row=(y,label,steps)=>`<text x="0" y="${y+25}" class="pipe-l">${label}</text>`+steps.map((s,i)=>box(70+i*125,y,110,s,i===steps.length-1?'dest':'')+(i<steps.length-1?arrow(180+i*125,y+20,195+i*125):'')).join('');
  return `<figure class="figure"><svg viewBox="0 0 570 130" role="img" aria-label="ETL: extract, transform, then load into a warehouse. ELT: extract, load into a lake or warehouse, then transform there.">
<defs><marker id="arr" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="7" markerHeight="7" orient="auto"><path d="M0 0L10 5L0 10z" class="pipe-head"/></marker></defs>
${row(10,'ETL',['Extract','Transform','Load','Warehouse'])}${row(75,'ELT',['Extract','Load','Transform','In the lake'])}</svg>
<figcaption>ETL cleans and reshapes data before it's stored. ELT stores the raw data first and transforms it later, inside the storage system.</figcaption></figure>`;
})();
