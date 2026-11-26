// ย่อหน้าว่างที่ตามหลังตาราง แล้วตามด้วยข้อความ/หัวข้อ (ไม่ใช่ตาราง) -> ลบ (ไม่งั้นพอไปอยู่บนสุดของหน้าจะเป็นช่องว่าง) ; ไม่แตะตัวคั่นระหว่างตาราง/ตัวแบ่ง section
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0;
x = x.replace(/(<\/w:tbl>)(<w:p(?: [^>]*)?>(?:<w:pPr>(?:(?!<\/w:pPr>)[\s\S])*<\/w:pPr>)?(?:<w:r>(?:<w:rPr>(?:(?!<\/w:rPr>)[\s\S])*<\/w:rPr>)?<w:lastRenderedPageBreak\/><\/w:r>)?<\/w:p>)(?=<w:p[ >])/g, (m, a, p) => {
  if (/<w:sectPr|<w:br /.test(p)) return m; n++; return a;
});
d.set('word/document.xml', x); saveDocx(d, f); console.log('empty paragraphs after tables removed:', n);
