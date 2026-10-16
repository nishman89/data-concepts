/* ---------- Quizzes ----------
   Multiple choice (select ONE) and value questions, answered with a formula run here or a value from Excel.
   One question at a time, an optional timer, a 65% pass mark, a score by unit and a full review. */
let QZ=null;
const LETTERS='abcd';
const fmtTime = s => `${Math.floor(s/60)}:${String(Math.floor(s%60)).padStart(2,'0')}`;
const unitTitle = n => UNITS[n-1] ? UNITS[n-1].title : `Unit ${n}`;
const quizDef = id => QUIZZES.find(q=>q.id===id);
const quizRecord = id => (P.quizzes||{})[id];
const attemptText = a => `${a.score} of ${a.total} (${a.percent}%), ${a.passed?'passed':'not passed'}`;
const isMcq = q => Array.isArray(q.options);
const answered = k => isMcq(QZ.questions[k]) ? QZ.answers[k]!=null : !!String(QZ.answers[k]||'').trim();
const valueTask = q => ({ sheet:q.sheet, sol:q.sol });

function markValue(q, ans){
  const a=String(ans||'').trim(); if (!a) return {state:'blank', msg:'Not answered.'};
  const task=valueTask(q); if (!q._task) q._task=task;
  const want=expected(q._task);
  if (a.startsWith('=')){
    const sh=WB.sheet(q.sheet), r=runFormula(WB, a, {sheet:sh,row:2,col:answerCol(sh)});
    if (r.error) return {state:'wrong', msg:`Your formula gave ${r.error.code}: ${r.error.why}`};
    return sameValue(r.value, want) ? {state:'right', msg:''} : {state:'wrong', msg:`Your formula gives ${fmtVal(r.value)}. The answer is ${fmtVal(want)}.`};
  }
  const t=readTyped(a);
  return sameValue(t.value, want, {typed:true, dp:t.dp}) ? {state:'right', msg:''} : {state:'wrong', msg:`The answer is ${fmtVal(want)}.`};
}

function quizCard(q){
  const rec=quizRecord(q.id), nv=q.questions.filter(x=>!isMcq(x)).length;
  return h('article',{class:'quiz-card'+(q.unitQuiz?' unit-card':''),'data-testid':`quiz-card-${q.id}`},
    h('h2',{},q.title), h('p',{},q.blurb),
    h('dl',{}, h('dt',{},'Questions'), h('dd',{},String(q.questions.length)), h('dt',{},'Type'), h('dd',{}, nv?`Multiple choice and ${nv} Excel question${nv>1?'s':''}`:'Multiple choice'),
      h('dt',{},'Last attempt'), h('dd',{}, rec?attemptText(rec.last):'Not tried yet'), rec?h('dt',{},'Best'):null, rec?h('dd',{},`${rec.best.percent}%`):null),
    h('button',{class:'btn primary',type:'button','data-go':'quiz/'+q.id,'data-testid':`open-quiz-${q.id}`}, rec?'Try again':'Start'));
}
function renderQuizList(){
  const v=$('#v-quizzes'); v.innerHTML='';
  v.append(h('h1',{class:'dot',id:'h-quizzes'},'Quizzes'), h('p',{class:'lede'},`Test yourself without the hints. Choose whether to be marked at the end, like a real test, or straight away after each answer. The pass mark is ${QUIZ_SETTINGS.passMark}%.`));
  v.append(h('h2',{class:'dot quiz-group-h'},'End-of-unit quizzes'), h('div',{class:'quiz-cards unit-cards'}, QUIZZES.filter(q=>q.unitQuiz).map(quizCard)));
  v.append(h('h2',{class:'dot quiz-group-h'},'Final quiz'), h('div',{class:'quiz-cards'}, QUIZZES.filter(q=>!q.unitQuiz).map(quizCard)));
}
function openQuiz(id){ stopTimer(); const def=quizDef(id); if (!def){ goto('quiz'); return; } QZ={id, def, questions:def.questions, answers:[], i:0, phase:'start', onlyWrong:false}; renderQuiz(); }
function renderQuiz(){
  const v=$('#v-quiz'); v.innerHTML='';
  v.append(h('nav',{class:'crumbs','aria-label':'Breadcrumb'}, h('a',{href:'#/quiz'},'Quizzes'), h('span',{class:'sep','aria-hidden':'true'},'/'), h('span',{},QZ.def.title)));
  v.append(h('h1',{class:'dot',id:'h-quiz','data-testid':'quiz-title'}, QZ.def.title));
  if (QZ.phase==='start') return renderQuizStart(v);
  if (QZ.phase==='doing') return renderQuizDoing(v);
  renderQuizResults(v);
}
function renderQuizStart(v){
  const d=QZ.def, rec=quizRecord(QZ.id), nv=d.questions.filter(x=>!isMcq(x)).length;
  const sel=h('select',{id:'time-limit','data-testid':'time-limit'}, d.timeLimits.map(m=>h('option',{value:String(m),selected:m===d.defaultLimit}, m?`${m} minutes`:'No time limit')));
  v.append(h('p',{class:'lede'},d.blurb), h('section',{class:'quiz-start','data-testid':'quiz-start'},
    h('ul',{class:'facts'}, h('li',{}, h('b',{},String(d.questions.length)), ' questions'), h('li',{}, 'Pass mark ', h('b',{},`${QUIZ_SETTINGS.passMark}%`)),
      nv ? h('li',{}, `${nv} question${nv>1?'s ask':' asks'} for a number from the data. Type a formula (starting with =) to work it out here, or type the answer you got in Excel.`) : null,
      h('li',{}, 'Leaving this page ends the attempt without saving it.')),
    rec ? h('p',{}, `Last attempt: ${attemptText(rec.last)}. Best: ${rec.best.percent}%.`) : null,
    h('label',{class:'field'}, h('span',{},'Time limit'), sel),
    h('fieldset',{class:'mark-mode'}, h('legend',{},'When do you want to see if you\'re right?'),
      h('label',{}, h('input',{type:'radio',name:'mark-mode',value:'end',id:'mark-end',checked:true}), h('span',{}, h('b',{},'At the end'), ' – like a real test')),
      h('label',{}, h('input',{type:'radio',name:'mark-mode',value:'now',id:'mark-now','data-testid':'mark-straight-away'}), h('span',{}, h('b',{},'Straight away'), ' – answers lock once checked'))),
    h('button',{class:'btn primary',type:'button','data-testid':'start-quiz',onclick:()=>startQuiz(+sel.value, $('#mark-now').checked)},'Start the quiz')));
}
function startQuiz(limit, instant){
  QZ.instant=instant; QZ.checked=QZ.questions.map(()=>null); QZ.answers=QZ.questions.map(q=>isMcq(q)?null:'');
  QZ.phase='doing'; QZ.i=0; QZ.limit=limit*60; QZ.started=Date.now(); renderQuiz();
  if (QZ.limit) QZ.timer=setInterval(tick,1000); tick();
}
function stopTimer(){ if (QZ && QZ.timer){ clearInterval(QZ.timer); QZ.timer=null; } }
const elapsed = () => Math.round((Date.now()-QZ.started)/1000);
function tick(){ const t=$('#timer'); if (!t) return; if (!QZ.limit){ t.textContent=fmtTime(elapsed()); return; } const left=QZ.limit-elapsed(); t.textContent=fmtTime(Math.max(0,left)); t.classList.toggle('low',left<=120); if (left<=0){ stopTimer(); toast('Time\'s up. Your answers have been submitted.'); submitQuiz(true); } }
function feedbackBox(q, mark, N){
  const ok=mark.state==='right';
  const box=h('div',{class:'instant '+(ok?'right':'wrong'),role:'status','data-testid':`question-${N}-feedback`}, h('b',{},ok?'Correct.':'Not quite.'), mark.msg?' '+mark.msg:'');
  if (isMcq(q) && !ok) box.append(h('p',{}, 'The answer is ', h('b',{},`${LETTERS[q.answer]}) `), h('span',{html:q.options[q.answer]})));
  if (!isMcq(q)) box.append(h('div',{class:'rv-h'},'One possible formula'), h('pre',{},q.sol));
  box.append(h('p',{class:'rv-why',html:q.why})); return box;
}
function renderQuizDoing(v){
  const q=QZ.questions[QZ.i], n=QZ.questions.length, N=QZ.i+1;
  const navq=h('ol',{class:'qnav','aria-label':'Questions'});
  QZ.questions.forEach((_,k)=>navq.append(h('li',{}, h('button',{type:'button',class:(answered(k)?'ans ':'')+(k===QZ.i?'cur':''),'aria-current':k===QZ.i?'step':null,'aria-label':`Question ${k+1}`,onclick:()=>goQ(k)}, String(k+1)))));
  v.append(navq);
  const card=h('section',{class:'qcard','data-testid':`question-${N}`});
  card.append(h('div',{class:'ex-num'}, `Question ${N} of ${n}`, h('span',{class:'qunit'}, unitTitle(q.unit))));
  card.append(h('p',{class:'ex-prompt',html:q.q}));
  const locked = QZ.instant && QZ.checked[QZ.i];
  if (isMcq(q)){
    card.append(h('div',{class:'q-instr'},'Select ONE'));
    const opts=h('div',{class:'mcq-opts',role:'radiogroup'});
    q.options.forEach((text,k)=>{
      const inp=h('input',{type:'radio',name:`question-${N}`,id:`q${N}-${k}`,'data-testid':`question-${N}-option-${LETTERS[k]}-input`,checked:QZ.answers[QZ.i]===k,disabled:!!locked});
      inp.addEventListener('change',()=>{ QZ.answers[QZ.i]=k; refreshAnswered(); if (QZ.instant){ QZ.checked[QZ.i]={state:k===q.answer?'right':'wrong',msg:''}; renderQuiz(); } });
      const stc = locked ? (k===q.answer?' is-right':(QZ.answers[QZ.i]===k?' is-wrong':'')) : '';
      opts.append(h('label',{class:'mcq-opt'+stc,for:`q${N}-${k}`}, inp, h('span',{class:'key'},LETTERS[k]+')'), h('span',{html:text})));
    });
    card.append(opts);
  } else {
    const preview=h('div',{class:'xl-msg','data-testid':`question-${N}-preview`});
    const inp=h('input',{type:'text',class:'answer-input wide',id:`qv${N}`,'data-testid':`question-${N}-input`,spellcheck:'false',autocomplete:'off',placeholder:`A formula such as =COUNTIF(…), or the answer from Excel`,value:QZ.answers[QZ.i]||'',readonly:!!locked});
    const upd=()=>{ QZ.answers[QZ.i]=inp.value; refreshAnswered(); preview.innerHTML=''; const a=inp.value.trim();
      if (a.startsWith('=')){ const sh=WB.sheet(q.sheet), r=runFormula(WB, a, {sheet:sh,row:2,col:answerCol(sh)}); if (r.error) preview.append(h('span',{class:'muted'},`${r.error.code} ${r.error.why}`)); else preview.append('Your formula gives ', h('b',{},isErr(r.value)?r.value.code:fmtVal(r.value))); } };
    inp.addEventListener('input',upd);
    card.append(h('p',{class:'small'}, 'Uses the ', h('strong',{},q.sheet), ' sheet. ', h('button',{type:'button',class:'linkish','data-testid':`question-${N}-open-data`,onclick:()=>openData(q.sheet)},'Look at the data')),
      h('label',{for:`qv${N}`,class:'q-instr'},'Your answer'), inp, preview);
    setTimeout(upd);
    if (QZ.instant && !locked) card.append(h('p',{class:'check-row'}, h('button',{class:'btn primary',type:'button',onclick:()=>{ if (!answered(QZ.i)){ toast('Type an answer first.'); return; } QZ.checked[QZ.i]=markValue(q, QZ.answers[QZ.i]); renderQuiz(); }},'Check answer')));
  }
  if (locked) card.append(feedbackBox(q, QZ.checked[QZ.i], N));
  v.append(card);
  v.append(h('div',{class:'submit-bar'},
    h('span',{class:'answered'}, h('b',{id:'answered-count'},'0'), ` of ${n} answered`),
    h('div',{class:'mid'}, h('button',{class:'btn small',type:'button',disabled:QZ.i===0,onclick:()=>goQ(QZ.i-1),'data-testid':'previous-question'},'Previous'), h('span',{},`${N} / ${n}`), h('button',{class:'btn small',type:'button',disabled:QZ.i===n-1,onclick:()=>goQ(QZ.i+1),'data-testid':'next-question'},'Next')),
    h('span',{class:'timer-wrap'}, QZ.limit?'Time left ':'Time ', h('b',{id:'timer'},'')),
    h('button',{class:'btn primary',type:'button','data-testid':'submit-answers',onclick:()=>submitQuiz(false)},'Submit answers')));
  refreshAnswered(); tick();
}
function refreshAnswered(){ const c=QZ.answers.filter((_,k)=>answered(k)).length; const el=$('#answered-count'); if (el) el.textContent=String(c); $$('.qnav button').forEach((b,k)=>b.classList.toggle('ans',answered(k))); }
function goQ(k){ if (k<0||k>=QZ.questions.length) return; QZ.i=k; renderQuiz(); $('.qnav').scrollIntoView({block:'nearest'}); }
function submitQuiz(timeUp){
  const missing=QZ.answers.filter((_,k)=>!answered(k)).length;
  if (!timeUp && missing && !confirm(`You haven't answered ${missing} question${missing===1?'':'s'}. Submit anyway?`)) return;
  stopTimer();
  const marks=QZ.questions.map((q,k)=>{ if (QZ.instant && QZ.checked[k]) return QZ.checked[k]; if (isMcq(q)) return QZ.answers[k]==null?{state:'blank',msg:''}:{state:QZ.answers[k]===q.answer?'right':'wrong',msg:''}; return markValue(q, QZ.answers[k]); });
  const score=marks.filter(m=>m.state==='right').length, total=marks.length, percent=Math.round(score/total*100);
  const attempt={score,total,percent,passed:percent>=QUIZ_SETTINGS.passMark,seconds:elapsed(),date:new Date().toISOString(),units:{}};
  QZ.questions.forEach((q,k)=>{ const u=attempt.units[q.unit]||(attempt.units[q.unit]=[0,0]); u[1]++; if (marks[k].state==='right') u[0]++; });
  const rec=P.quizzes[QZ.id]||{attempts:0,best:attempt}; rec.attempts++; rec.last=attempt; if (attempt.percent>=rec.best.percent) rec.best=attempt; P.quizzes[QZ.id]=rec; save();
  QZ.results={marks, attempt}; QZ.phase='done'; renderQuiz(); window.scrollTo({top:0});
}
function renderQuizResults(v){
  const {marks, attempt}=QZ.results;
  v.append(h('section',{class:'results '+(attempt.passed?'pass':'fail'),'data-testid':'results'},
    h('div',{class:'verdict-big','data-testid':'quiz-verdict'}, attempt.passed?'Passed':'Not passed yet'),
    h('div',{class:'score-line'}, 'You scored ', h('b',{'data-testid':'score'},String(attempt.score)), ` out of ${attempt.total} (${attempt.percent}%). The pass mark is ${QUIZ_SETTINGS.passMark}%.`),
    h('div',{class:'small'}, 'Time taken: ', fmtTime(attempt.seconds))));
  const units=Object.keys(attempt.units);
  if (units.length>1){ const tb=h('tbody'); units.sort((a,b)=>a-b).forEach(u=>{ const [a,b]=attempt.units[u]; tb.append(h('tr',{}, h('td',{},`Unit ${u}: ${unitTitle(+u)}`), h('td',{},`${a} of ${b}`), h('td',{}, h('a',{href:'#/lesson/'+UNITS[u-1].lessons[0]},'Revise this unit')))); });
    v.append(h('h2',{class:'dot',style:'margin-top:28px'},'Score by unit'), h('div',{class:'tbl-wrap'}, h('table',{class:'grid breakdown'}, h('thead',{},h('tr',{},h('th',{},'Unit'),h('th',{},'Score'),h('th',{},''))), tb))); }
  else { const u=+units[0]; $('[data-testid="results"]',v).append(h('p',{class:'small',style:'margin-top:8px'}, h('a',{href:'#/lesson/'+UNITS[u-1].lessons[0]},`Revise unit ${u}: ${unitTitle(u)}`))); }
  const toggle=h('input',{type:'checkbox',id:'only-wrong',checked:QZ.onlyWrong}); toggle.addEventListener('change',()=>{ QZ.onlyWrong=toggle.checked; $$('.review',v).forEach(r=>r.hidden=QZ.onlyWrong&&r.dataset.state==='right'); });
  v.append(h('div',{class:'review-tools'}, h('h2',{class:'dot'},'Review'), h('label',{class:'toggle'}, toggle, ' Show only questions I got wrong'),
    h('div',{style:'margin-left:auto;display:flex;gap:8px;flex-wrap:wrap'}, h('button',{class:'btn primary',type:'button','data-testid':'retake',onclick:()=>openQuiz(QZ.id)},'Retake quiz'), h('button',{class:'btn',type:'button','data-go':'quiz'},'All quizzes'))));
  QZ.questions.forEach((q,k)=>{
    const m=marks[k], N=k+1, label={right:'Correct',wrong:'Incorrect',blank:'Not answered'}[m.state];
    const r=h('article',{class:'review '+m.state,'data-state':m.state,'data-testid':`question-${N}-review`,hidden:QZ.onlyWrong&&m.state==='right'},
      h('div',{class:'ex-num'}, `Question ${N}`, h('span',{class:'qunit'}, unitTitle(q.unit))), h('p',{class:'ex-prompt',html:q.q}), h('div',{class:'rv-verdict'}, h('b',{},label+'.'), m.msg?' '+m.msg:''));
    if (isMcq(q)){ const ul=h('ul',{class:'rv-opts'}); q.options.forEach((text,j)=>{ const mine=QZ.answers[k]===j, right=q.answer===j; ul.append(h('li',{class:(right?'is-right ':'')+(mine&&!right?'is-wrong':'')}, h('span',{class:'key'},LETTERS[j]+')'), h('span',{html:text}), h('span',{class:'flag'}, right?(mine?'Correct answer (yours)':'Correct answer'):mine?'Your answer':''))); }); r.append(ul); }
    else r.append(h('div',{class:'rv-cols'}, h('div',{}, h('div',{class:'rv-h'},'Your answer'), h('pre',{}, String(QZ.answers[k]||'(blank)'))), h('div',{}, h('div',{class:'rv-h'},'One possible formula'), h('pre',{}, q.sol))));
    r.append(h('p',{class:'rv-why',html:q.why})); v.append(r);
  });
}
