import { loadDocx } from '../lib/docx.mjs';
const d = loadDocx(process.argv[2]);
for (const e of d.entries) { if (!/\.xml$/.test(e.name)) continue; const x = d.get(e.name); const m = x.match(/w:ascii="Arial"/g); if (!m) continue; const i = x.indexOf('w:ascii="Arial"'); console.log(e.name, m.length, '\n', x.slice(Math.max(0, i - 400), i + 150).replace(/\u200B/g, '')); }
