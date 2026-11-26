// สร้างไฟล์ต้นแบบหัวกระดาษ (เอาเนื้อหาออก เหลือแค่ section สุดท้าย + หัว/ท้ายกระดาษ + รูปโลโก้)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const [src, dst] = process.argv.slice(2); const d = loadDocx(src); const x = d.get('word/document.xml');
const last = [...x.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].pop()[0];
const head = x.slice(0, x.indexOf('<w:body>') + 8);
d.set('word/document.xml', head + '<w:p/>' + last + '</w:body></w:document>');
const keep = new Set(); const rels = d.get('word/_rels/document.xml.rels');
for (const m of last.matchAll(/r:id="([^"]+)"/g)) { const t = (rels.match(new RegExp(`Id="${m[1]}"[^>]*Target="([^"]+)"`)) || [])[1]; if (t) keep.add('word/' + t); }
for (const p of [...keep]) { const r = d.get(p.replace(/^word\//, 'word/_rels/') + '.rels') || ''; for (const m of r.matchAll(/Target="(media\/[^"]+)"/g)) keep.add('word/' + m[1]); }
d.entries = d.entries.filter(e => !/^word\/media\//.test(e.name) || keep.has(e.name));
saveDocx(d, dst); console.log('header template saved, media kept:', [...keep].filter(k => /media/.test(k)).join(', '));
