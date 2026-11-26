// ดูหัวกระดาษของเล่มก่อนตัดสินใจ (ผู้ใช้ 2026-11-26 "คงหัวกระดาษเดิม" + ตรงกับเล่ม 1–6 ที่ผ่าน):
//   none = ไม่มีหัวกระดาษ
//   old  = หัวกระดาษโครงการเดียวกับต้นแบบ แต่เป็นแบบเก่า (ย่อหน้าใช้สไตล์ชื่อ "หัวกระดาษ"/header ไม่ใช่ "หัวกระดาษ1") -> ใช้ต้นแบบแทน (เล่ม 4, 5, 6)
//   same = โครงการเดียวกันและเป็นแบบเดียวกับต้นแบบอยู่แล้ว -> คงไว้ (เล่ม 2, 3)
//   own  = หัวกระดาษของเล่มเอง (จัดรูปแบบเองไม่ใช้สไตล์ หรือข้อความคนละโครงการ) -> คงไว้ (เล่ม 1)
import { loadDocx } from '../lib/docx.mjs';
const [f, tpl] = process.argv.slice(2);
const parts = d => d.entries.filter(e => /^word\/header\d*\.xml$/.test(e.name)).map(e => d.get(e.name));
const txtOf = x => [...x.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').replace(/[\s​]+/g, '');
const D = loadDocx(f), T = loadDocx(tpl);
const key = parts(T).map(txtOf).join('').slice(0, 25);
const mine = parts(D).filter(x => txtOf(x));
let kind;
if (!mine.length) kind = 'none';
else if (!key || !mine.some(x => txtOf(x).includes(key))) kind = 'own';
else {
  const st = D.get('word/styles.xml') || '';
  const nameOf = id => { const i = st.indexOf(`w:styleId="${id}"`); if (i < 0) return ''; const e = st.indexOf('</w:style>', i); return ((st.slice(i, e).match(/<w:name w:val="([^"]+)"/) || [])[1] || '').toLowerCase(); };
  const names = new Set(mine.flatMap(x => [...x.matchAll(/<w:pStyle w:val="([^"]+)"\/>/g)].map(m => nameOf(m[1]))).filter(Boolean));
  if (!names.size) kind = 'own';                       // จัดรูปแบบเอง ไม่ใช้สไตล์ (เล่ม 1)
  else if (names.has('หัวกระดาษ1')) kind = 'same';      // แบบเดียวกับต้นแบบอยู่แล้ว (เล่ม 2, 3)
  else kind = 'old';                                   // แบบเก่า "หัวกระดาษ" / header (เล่ม 4, 5, 6)
}
console.log(kind);
