// ตารางกว้าง: ย่อหน้าภาษาไทย (>=30%) ที่จัดแบบ both -> thaiDistribute (แบบเดียวกับเล่ม 3 ที่ผ่านแล้ว) ; ข้อความอังกฤษล้วนคง both
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0;
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => { const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g < 7) return t;
  return t.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => { if (!/<w:jc w:val="both"\/>/.test(p)) return p; const s = [...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').replace(/[\u200B\s]/g, ''); if (!s || (s.match(/[\u0E00-\u0E7F]/g) || []).length / s.length < 0.3) return p; n++; return p.replace(/<w:jc w:val="both"\/>/, '<w:jc w:val="thaiDistribute"/>'); }); });
d.set('word/document.xml', x); saveDocx(d, f); console.log('Thai cells back to Thai justify:', n);
