import { loadDocx } from '../lib/docx.mjs';
const x = loadDocx(process.argv[2]).get('word/document.xml'); const k = {};
for (const m of x.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)) { const r0 = (m[0].match(/<w:tr\b[\s\S]*?<\/w:tr>/) || [''])[0]; const h = [...r0.matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(c => [...c[1].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').trim()).join('|'); const rows = (m[0].match(/<w:tr\b/g) || []).length; const img = (m[0].match(/<w:drawing>/g) || []).length; const key = h.slice(0, 120); k[key] ||= { n: 0, rows: 0, img: 0 }; k[key].n++; k[key].rows += rows; k[key].img += img; }
for (const [h, v] of Object.entries(k)) console.log(`${v.n}x rows=${v.rows} img=${v.img} :: ${h}`);
