// แถว "ผลที่พบ" (แถวก่อนแถวรูป): ไม่ต้องติดกับรูป — รูปขนาดใหญ่ขึ้นหน้าใหม่ได้ (แถวทดสอบยังติดกับแถวผลที่พบ)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0;
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  if (!/<w:drawing/.test(t)) return t;
  const rows = [...t.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(m => m[0]);
  for (let i = 0; i < rows.length - 1; i++) if (!/<w:drawing/.test(rows[i]) && /<w:drawing/.test(rows[i + 1]) && !/<w:tblHeader\/>/.test(rows[i])) { const nr = rows[i].replace(/<w:keepNext\/>/g, ''); if (nr !== rows[i]) { n++; t = t.replace(rows[i], nr); } }
  return t;
});
d.set('word/document.xml', x); saveDocx(d, f); console.log('result rows released from picture rows:', n);
