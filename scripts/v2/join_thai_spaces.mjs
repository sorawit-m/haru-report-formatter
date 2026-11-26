// ช่องว่างระหว่างข้อความไทยที่แยกเป็น run ของตัวเอง (ไม่มีป้ายภาษาไทย) -> รวมเข้า run ไทยข้างๆ ให้เป็น run เดียว
// Word วัดช่องว่างที่ไม่มีป้ายไทยด้วยตัวอักษรละติน ทำให้ตัดบรรทัดต่างจากเล่ม 6 ที่ผ่าน (ใช้กับเล่มผลทดสอบ)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0;
const TH = /[฀-๿]/;
const base = rpr => rpr.replace(/<w:cs\/>/g, '').replace(/<w:lang [^>]*\/>/g, '');
// run = <w:r><w:rPr>..</w:rPr><w:t ..>text</w:t></w:r>  (เฉพาะ run ที่มีแค่ข้อความชิ้นเดียว)
const RUN = '<w:r(?: [^>]*)?>(<w:rPr>(?:(?!</w:rPr>)[\\s\\S])*</w:rPr>)<w:t(?: [^>]*)?>([^<]*)</w:t></w:r>';
const re = new RegExp(RUN + RUN + RUN, 'g');
let changed = true;
while (changed) {
  changed = false;
  x = x.replace(re, (m, p1, t1, p2, t2, p3, t3) => {
    if (!/^ +$/.test(t2) || !TH.test(t1.slice(-1)) || !TH.test(t3[0])) return m;
    if (!/<w:cs\/>/.test(p1) || !/<w:cs\/>/.test(p3)) return m;
    if (base(p1) !== base(p2) || base(p1) !== base(p3)) return m;
    n++; changed = true;
    return `<w:r>${p1}<w:t xml:space="preserve">${t1}${t2}${t3}</w:t></w:r>`;
  });
}
d.set('word/document.xml', x); saveDocx(d, f); console.log('thai spaces joined:', n);
