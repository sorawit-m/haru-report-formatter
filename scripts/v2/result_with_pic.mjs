// แถวข้อความ "ผลที่พบ" ที่แยกเป็นแถวของตัวเอง (1 ช่อง) และตามด้วยแถวรูป -> ติดกับแถวรูป
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0;
const addKeep = p => /<w:keepNext\/>/.test(p) ? p : (/<w:pPr>/.test(p) ? p.replace(/<w:pPr>(?:(<w:pStyle [^>]*\/>))?/, (m, s) => `<w:pPr>${s || ''}<w:keepNext/>`) : p.replace(/^(<w:p(?: [^>]*)?>)/, '$1<w:pPr><w:keepNext/></w:pPr>'));
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  const rows = [...t.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(m => m[0]);
  for (let i = 0; i < rows.length - 1; i++) {
    const one = (rows[i].match(/<w:tc>/g) || []).length === 1;
    if (one && !/<w:drawing/.test(rows[i]) && /<w:drawing/.test(rows[i + 1])) { const r = rows[i].replace(/<w:p[ >][\s\S]*?<\/w:p>/g, addKeep); if (r !== rows[i]) { t = t.replace(rows[i], r); n++; } }
  }
  return t;
});
d.set('word/document.xml', x); saveDocx(d, f); console.log('result rows tied to picture rows:', n);
