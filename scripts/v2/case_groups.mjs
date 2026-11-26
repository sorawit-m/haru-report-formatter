// ตารางทดสอบแบบ (แถวทดสอบ -> แถวผลที่พบ -> แถวรูป): ให้ทั้งกลุ่มอยู่หน้าเดียวกัน
// แถวที่ไม่ใช่หัวตารางและไม่มีรูป -> keepNext ทุกย่อหน้า (ติดกับแถวถัดไป); แถวที่มีรูป -> ห้ามแยก, ไม่ keepNext (จบกลุ่ม)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const cnt = { kept: 0, pics: 0 };
const addKeep = p => /<w:keepNext\/>/.test(p) ? p : (/<w:pPr>/.test(p) ? p.replace(/<w:pPr>(?:(<w:pStyle [^>]*\/>))?/, (m, s) => `<w:pPr>${s || ''}<w:keepNext/>`) : p.replace(/^(<w:p(?: [^>]*)?>)/, '$1<w:pPr><w:keepNext/></w:pPr>'));
const cant = tr => /<w:cantSplit\/>/.test(tr) ? tr : (/<w:trPr>/.test(tr) ? tr.replace(/<w:trPr>/, '<w:trPr><w:cantSplit/>') : tr.replace(/^(<w:tr\b[^>]*>)/, '$1<w:trPr><w:cantSplit/></w:trPr>'));
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g < 7 || !/<w:drawing/.test(t)) return t;
  return t.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, tr => {
    if (/<w:tblHeader\/>/.test(tr)) return tr;
    if (/<w:drawing/.test(tr)) { cnt.pics++; return cant(tr).replace(/<w:keepNext\/>/g, ''); }
    cnt.kept++; return cant(tr).replace(/<w:p[ >][\s\S]*?<\/w:p>/g, addKeep);
  });
});
d.set('word/document.xml', x); saveDocx(d, f); console.log(JSON.stringify(cnt));
