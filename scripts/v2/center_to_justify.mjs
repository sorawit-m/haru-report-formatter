// ข้อความในตารางที่จัดกึ่งกลาง แต่เป็นข้อความไทยยาว (> 12 ตัวอักษร) -> ชิดซ้าย-ขวา (ไม่แตะแถวหัวตาราง และข้อความสั้น เช่น รหัส / Pass)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0;
const shaded = tc => /<w:shd [^>]*w:fill="(?!auto|FFFFFF|ffffff)[0-9A-Fa-f]{6}"/.test(tc);
x = x.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, tr => {
  const tcs = [...tr.matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(m => m[1]); if (tcs.length > 1 && tcs.every(shaded)) return tr;
  return tr.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
    if (!/<w:jc w:val="center"\/>/.test(p) || /<w:drawing/.test(p)) return p;
    const t = [...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').replace(/\u200B/g, '');
    if (t.trim().length <= 12 || !/[\u0E00-\u0E7F]/.test(t)) return p;
    n++; return p.replace(/<w:jc w:val="center"\/>/, '<w:jc w:val="thaiDistribute"/>');
  });
});
d.set('word/document.xml', x); saveDocx(d, f); console.log('centered long Thai cells -> justified:', n);
