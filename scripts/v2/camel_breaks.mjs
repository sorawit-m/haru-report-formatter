// จุดตัดที่มองไม่เห็นในคำอังกฤษยาว (>=13 ตัว) ตามตัวพิมพ์ใหญ่ เช่น encode|URI|Component และหลัง "(" — ไม่แตะกล่องโค้ด (ตาราง 1 คอลัมน์) และฟิลด์
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0;
const b = x.indexOf('<w:body>'); const ce = x.indexOf('</w:p>', x.indexOf('<w:sectPr', b)) + 6;
let body = x.slice(ce); const saved = [];
body = body.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => { const cols = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (cols !== 1) return t; saved.push(t); return `\u0000C${saved.length - 1}\u0000`; });
const fix = tok => tok
  .replace(/([a-z])(?=[A-Z])/g, (m, a) => { n++; return a + '\u200B'; })            // camelCase: encode|URI
  .replace(/([A-Z])(?=[A-Z][a-z])/g, (m, a) => { n++; return a + '\u200B'; })       // acronym end: URI|Component
  .replace(/^\((?=[A-Za-z])/, () => { n++; return '(\u200B'; });
body = body.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => /<w:instrText|<w:fldSimple/.test(p) ? p :
  p.replace(/(<w:t(?: [^>]*)?>)([^<]*)(<\/w:t>)/g, (a, o, t, c) => o + t.replace(/[(A-Za-z][A-Za-z0-9\u200B]{9,}\)?/g, tok => tok.replace(/\u200B/g, '').length >= 10 && !tok.includes('\u200B') ? fix(tok) : tok) + c));
body = body.replace(/\u0000C(\d+)\u0000/g, (m, i) => saved[+i]);
d.set('word/document.xml', x.slice(0, ce) + body); saveDocx(d, f); console.log('camelCase break points:', n);
