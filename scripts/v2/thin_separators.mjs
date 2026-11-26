// ย่อหน้าว่างที่คั่นระหว่างตารางสองตาราง -> สูงแค่ 1pt (ลบทิ้งไม่ได้ เพราะตารางจะรวมกัน) กันไม่ให้มันล้นไปอยู่หน้าใหม่คนเดียว
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0;
const THIN = '<w:pPr><w:spacing w:before="0" w:after="0" w:line="40" w:lineRule="exact"/><w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr></w:pPr>';
x = x.replace(/(<\/w:tbl>)(<w:p(?: [^>]*)?>)((?:<w:pPr>(?:(?!<\/w:pPr>)[\s\S])*<\/w:pPr>)?)(<\/w:p>)(?=<w:tbl>)/g, (m, a, open, ppr, close) => {
  if (/<w:sectPr/.test(ppr)) return m;
  if (/w:line="200" w:lineRule="exact"/.test(ppr)) return m;   // ตัวคั่นระหว่างกรณีทดสอบ (split_cases) ตั้งใจให้สูง 10pt เหมือนเล่ม 6 ที่ผ่าน — ห้ามย่อ
  n++; return a + open + THIN + close;
});
d.set('word/document.xml', x); saveDocx(d, f); console.log('separators thinned:', n);
