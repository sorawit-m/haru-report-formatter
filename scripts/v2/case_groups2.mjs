// ตารางทดสอบที่ "ผลที่พบ" กับรูปอยู่ในช่องเดียวกัน: แถวทดสอบติดกับแถวผลที่พบ; ในช่องผลที่พบ ข้อความติดกับรูปไม่ได้บังคับ (รูปใหญ่ขึ้นหน้าใหม่ได้)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const cnt = { testRows: 0, picRows: 0 };
const addKeep = p => /<w:keepNext\/>/.test(p) ? p : (/<w:pPr>/.test(p) ? p.replace(/<w:pPr>(?:(<w:pStyle [^>]*\/>))?/, (m, s) => `<w:pPr>${s || ''}<w:keepNext/>`) : p.replace(/^(<w:p(?: [^>]*)?>)/, '$1<w:pPr><w:keepNext/></w:pPr>'));
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g < 7 || !/<w:drawing/.test(t)) return t;
  return t.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, tr => {
    if (/<w:tblHeader\/>/.test(tr)) return tr;
    if (/<w:drawing/.test(tr)) { cnt.picRows++; return tr.replace(/<w:cantSplit\/>/g, '').replace(/<w:trPr><\/w:trPr>/g, '').replace(/<w:keepNext\/>/g, ''); }
    cnt.testRows++; return tr.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, addKeep);
  });
});
d.set('word/document.xml', x); saveDocx(d, f); console.log(JSON.stringify(cnt));
