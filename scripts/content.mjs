// Loads the engine and content in Node, exactly as the page does, for the checks.
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';
export const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const CONTENT_FILES = ['helpers', 'lessons', 'intro', 'capstone', 'quizzes'].map(f => `src/content/${f}.js`);
export function load(){
  const read = p => fs.readFileSync(path.join(root, p), 'utf8');
  const src = `const WORKBOOK = ${read('data/workbook.json')};\nconst CAP_FACTS = ${read('data/capstone_facts.json')};\n`
    + read('src/engine.js') + '\n' + CONTENT_FILES.map(read).join('\n')
    + '\n;({ WORKBOOK, CAP_FACTS, UNITS, LESSONS, PATH, QUIZZES, QUIZ_SETTINGS, CAPSTONE_TASKS, makeWorkbook, runFormula, sameValue, isErr, lastRow, displayValue });';
  return vm.runInNewContext(src, { console }, { filename: 'site.js' });
}
