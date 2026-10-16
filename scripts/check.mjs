// Checks every lesson, exercise and quiz question before a build. Any problem stops the build with a message
// naming the lesson and the exercise, so a broken answer never reaches learners.
import { load } from './content.mjs';
const S = load();
const WB = S.makeWorkbook(S.WORKBOOK);
const errors = [], warnings = [];
const err = (where, msg) => errors.push(`${where}: ${msg}`);
const seen = new Set();
for (const u of S.UNITS) for (const id of u.lessons){
  if (seen.has(id)) err(id, 'is listed twice in UNITS'); seen.add(id);
  if (!S.LESSONS[id]) err(id, 'is in UNITS but not in LESSONS');
}
for (const id of Object.keys(S.LESSONS)) if (!seen.has(id)) warnings.push(`${id}: isn't in any unit, so it won't appear on the path`);
function checkTask(where, t){
  const sh = WB.sheet(t.sheet); if (!sh) return err(where, `sheet "${t.sheet}" doesn't exist`);
  if (!String(t.sol||'').startsWith('=')) return err(where, 'sol must be a formula starting with =');
  const target = { sheet: sh, row: 2, col: sh.cols.length + 1 };
  const rows = t.fill ? Math.min(S.lastRow(sh), 301) - 1 : 1;
  for (let i = 0; i < rows; i++){
    const r = S.runFormula(WB, t.sol, target, { dr: i });
    if (r.error) return err(where, `the model formula doesn't run: ${r.error.code} ${r.error.why}`);
    if (S.isErr(r.value)){ if (!t.fill) return err(where, `the model formula gives ${r.value.code} ${r.value.why}`); continue; }
  }
  if (t.askRow && (!t.fill || t.askRow < 2 || t.askRow > rows + 1)) err(where, 'askRow must be a data row, and only for fill exercises');
  if (t.fill){ const r = S.runFormula(WB, t.sol, target, { dr: (t.askRow || 2) - 2 }); if (S.isErr(r.value)) err(where, `the row learners are asked about gives ${r.value.code}`); }
}
for (const [id, L] of Object.entries(S.LESSONS)){
  if (L.kind) continue;
  if (!L.title || !L.summary) err(id, 'needs a title and a summary');
  if (!L.learn) err(id, 'needs a learn tab');
  if (!L.ex && !L.check) err(id, 'needs exercises (ex) or questions (check)');
  (L.ex || []).forEach((ex, i) => { const w = `${id} exercise ${i + 1}`; if (!ex.p || !ex.hint) err(w, 'needs p and hint'); checkTask(w, ex); if (ex.fill && !ex.col) warnings.push(`${w}: has no col name, so the column is called Answer`); });
  (L.check || []).forEach((q, i) => { const w = `${id} question ${i + 1}`; if (!q.q || !q.why || !Array.isArray(q.a) || q.a.length < 2) err(w, 'needs q, why and at least two options in a'); if (!(q.c >= 0 && q.c < (q.a||[]).length)) err(w, 'c must be the index of the right option'); if (new Set(q.a).size !== (q.a||[]).length) err(w, 'has two identical options'); });
  // Formula examples in the Learn tab must run too
  for (const m of L.learn.matchAll(/data-formula="([^"]*)" data-sheet="([^"]*)"(?: data-fill="1")?/g)){
    const f = m[1].replace(/&quot;/g, '"').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
    const sh = WB.sheet(m[2]); if (!sh){ err(id, `example ${f} uses unknown sheet ${m[2]}`); continue; }
    const r = S.runFormula(WB, f, { sheet: sh, row: 2, col: sh.cols.length + 1 });
    if (r.error || (S.isErr(r.value) && !/VLOOKUP/.test(f))) err(id, `example ${f} gives ${(r.error || r.value).code} ${(r.error || r.value).why}`);
  }
}
const ids = new Set();
for (const q of S.QUIZZES){
  if (ids.has(q.id)) err(q.id, 'two quizzes share this id'); ids.add(q.id);
  if (!q.questions.length) err(q.id, 'has no questions');
  if (!q.timeLimits.includes(q.defaultLimit)) err(q.id, 'defaultLimit must be one of timeLimits');
  const spread = { a: 0, b: 0, c: 0, d: 0 };
  q.questions.forEach((x, i) => {
    const w = `${q.id} question ${i + 1}`;
    if (!(x.unit >= 1 && x.unit <= S.UNITS.length)) err(w, 'unit must be a unit number');
    if (!x.q || !x.why) err(w, 'needs q and why');
    if (Array.isArray(x.options)){ if (x.options.length !== 4 || new Set(x.options).size !== 4) err(w, 'needs four different options'); if (!(x.answer >= 0 && x.answer <= 3)) err(w, 'answer must be 0 to 3'); else spread['abcd'[x.answer]]++; }
    else checkTask(w, x);
  });
  console.log(`  ${q.id.padEnd(8)} answers a/b/c/d: ${Object.values(spread).join('/')}`);
}
for (const w of warnings) console.warn('Warning: ' + w);
if (errors.length){ console.error(`\n${errors.length} problem${errors.length > 1 ? 's' : ''} found:\n` + errors.map(e => '  ✗ ' + e).join('\n')); process.exit(1); }
const nEx = Object.values(S.LESSONS).reduce((n, L) => n + (L.ex || []).length, 0);
console.log(`✓ ${S.PATH.length} lessons, ${nEx} Excel exercises and ${S.QUIZZES.reduce((n, q) => n + q.questions.length, 0)} quiz questions checked.`);
