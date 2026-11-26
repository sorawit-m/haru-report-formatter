// ตารางทดสอบยาว (หัวตาราง + หลายกรณี แต่ละกรณี = แถวทดสอบ + แถวผลที่พบ/รูป) -> แยกเป็นตารางละ 1 กรณี มีหัวตารางของตัวเอง
// หน้าที่มีแต่รูปจึงไม่มีหัวตารางคั่น และทุกกรณีเริ่มด้วยหัวตาราง
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const cnt = { tables: 0, cases: 0 };
const shaded = tc => /<w:shd [^>]*w:fill="(?!auto|FFFFFF|ffffff)[0-9A-Fa-f]{6}"/.test(tc);
const SEP = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="200" w:lineRule="exact"/><w:rPr><w:sz w:val="8"/><w:szCs w:val="8"/></w:rPr></w:pPr></w:p>';
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  if (/<w:tbl>[\s\S]*<w:tbl>/.test(t) || !/<w:drawing/.test(t)) return t;
  const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g < 7) return t;
  const head = t.slice(0, t.search(/<w:tr\b/)); const rows = [...t.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(m => m[0]);
  const hdr = []; let i = 0;
  while (i < rows.length) { const tcs = [...rows[i].matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(m => m[1]); if (tcs.length > 1 && tcs.every(shaded)) { hdr.push(rows[i].replace(/<w:tblHeader\/>/g, '')); i++; } else break; }
  if (!hdr.length) return t;
  const cases = []; let cur = null;
  for (; i < rows.length; i++) {
    const r = rows[i]; const tcs = [...r.matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(m => m[1]);
    const isStart = tcs.length >= 5 && !/<w:drawing/.test(r);
    if (isStart || !cur) { cur = []; cases.push(cur); }
    cur.push(r);
  }
  if (cases.length < 2) return t;
  cnt.tables++; cnt.cases += cases.length;
  return cases.map(c => head + hdr.join('') + c.join('') + '</w:tbl>').join(SEP);
});
d.set('word/document.xml', x); saveDocx(d, f); console.log(JSON.stringify(cnt));
