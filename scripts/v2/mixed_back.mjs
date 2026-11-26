// ย่อหน้าที่ตั้งเป็น both แต่มีภาษาไทยตั้งแต่ 30% -> กลับเป็น thaiDistribute (both ใช้ได้ดีเฉพาะข้อความอังกฤษเกือบล้วน)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0;
x = x.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
  if (!/<w:jc w:val="both"\/>/.test(p)) return p;
  const t = [...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').replace(/\u200B/g, '').replace(/\s/g, '');
  if (!t || (t.match(/[\u0E00-\u0E7F]/g) || []).length / t.length < 0.3) return p;
  n++; return p.replace(/<w:jc w:val="both"\/>/, '<w:jc w:val="thaiDistribute"/>');
});
d.set('word/document.xml', x); saveDocx(d, f); console.log('mixed paragraphs back to thaiDistribute:', n);
