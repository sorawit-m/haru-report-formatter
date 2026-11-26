// จุดตัดที่มองไม่เห็น: หลัง _ ในคำอังกฤษ (ทุกความยาว >=8) และหลังเครื่องหมายคำพูดเปิดก่อนตัวอักษร — ไม่แตะกล่องโค้ด/ฟิลด์
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0;
const b = x.indexOf('<w:body>'); const ce = x.indexOf('</w:p>', x.indexOf('<w:sectPr', b)) + 6;
let body = x.slice(ce); const saved = [];
body = body.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => { const cols = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (cols !== 1) return t; saved.push(t); return `\u0000C${saved.length - 1}\u0000`; });
body = body.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => /<w:instrText|<w:fldSimple/.test(p) ? p :
  p.replace(/(<w:t(?: [^>]*)?>)([^<]*)(<\/w:t>)/g, (a, o, t, c) => o + t
    .replace(/[A-Za-z0-9]+(?:_[A-Za-z0-9]+)+/g, tok => tok.length >= 8 ? tok.replace(/_(?=[A-Za-z0-9])/g, () => { n++; return '_\u200B'; }) : tok)
    .replace(/(^|[\s(])(&quot;|["\u201C])(?=[A-Za-z])/g, (m, pre, q) => { n++; return pre + q + '\u200B'; }) + c));
body = body.replace(/\u0000C(\d+)\u0000/g, (m, i) => saved[+i]);
d.set('word/document.xml', x.slice(0, ce) + body); saveDocx(d, f); console.log('extra break points:', n);
