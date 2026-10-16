// Builds dist/index.html: one self-contained file with the styles, the data, the engine, the content and the app
// inlined. Run with `npm run build`, which runs the checks first.
import fs from 'node:fs';
import path from 'node:path';
import { root, CONTENT_FILES } from './content.mjs';
const read = p => fs.readFileSync(path.join(root, p), 'utf8');
const parts = [
  ['/*__CSS__*/', read('src/styles.css')],
  ['__WORKBOOK__', read('data/workbook.json')],
  ['__CAPFACTS__', read('data/capstone_facts.json')],
  ['/*__ENGINE__*/', read('src/engine.js')],
  ['/*__CONTENT__*/', CONTENT_FILES.map(read).join('\n')],
  ['/*__APP__*/', read('src/app.js') + '\n' + read('src/quiz.js')],
];
const dataUri = (p, type) => `data:${type};base64,${fs.readFileSync(path.join(root, p)).toString('base64')}`;
let html = read('src/index.html');
for (const [marker, value] of [['__SPARTA_LOGO__', dataUri('src/assets/sparta-logo.png','image/png')], ['__SPARTA_LOGO_WHITE__', dataUri('src/assets/sparta-logo-white.png','image/png')], ['__FAVICON__', dataUri('src/assets/favicon.svg','image/svg+xml')]]){
  if (!html.includes(marker)) throw new Error(`Placeholder ${marker} is missing from src/index.html`);
  html = html.split(marker).join(value);
}
for (const [marker, value] of parts){
  const at = html.indexOf(marker);
  if (at < 0) throw new Error(`Placeholder ${marker} is missing from src/index.html`);
  if (marker !== '/*__CSS__*/' && /<\/script/i.test(value)) throw new Error(`${marker}: the text "</script" would end the script early`);
  html = html.slice(0, at) + value + html.slice(at + marker.length);
}
fs.rmSync(path.join(root, 'dist'), { recursive: true, force: true });
fs.mkdirSync(path.join(root, 'dist'), { recursive: true });
fs.cpSync(path.join(root, 'public'), path.join(root, 'dist'), { recursive: true });   // the workbook and capstone downloads
fs.writeFileSync(path.join(root, 'dist/index.html'), html);
console.log(`Built dist/index.html (${(html.length / 1024 / 1024).toFixed(2)} MB)`);
