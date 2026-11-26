// find-stretched.mjs — หาเซลล์ตารางที่ "กระจายแบบไทย" จะทำให้ตัวอักษรถ่างมาก (จำลองการตัดบรรทัดด้วยฟอนต์จริง)
// ใช้: node find-stretched.mjs <ไฟล์.docx> [เกณฑ์ em = 0.12] [out.tsv]
// ผลลัพธ์ TSV: ตาราง/แถว/คอลัมน์ | ความกว้าง(em) | gap | ข้อความ | การตัดบรรทัดที่คาด
// ใช้คู่กับขั้น "เติมคำเชื่อม" (ผู้ใช้อนุญาต): Claude เขียนข้อความใหม่ลง replacements แล้วรันซ้ำให้ gap ลด
import fs from 'node:fs';
import { loadDocx, parseXml, all, els, el, attr, textOf } from './lib/docx.mjs';
import * as T from './lib/thaiwrap.mjs';
const [, , f, thr = '0.12', out] = process.argv;
if (!f) { console.error('usage: node find-stretched.mjs file.docx [0.12] [out.tsv]'); process.exit(1); }
if (!T.available) { console.error('TH SarabunPSK font file not found'); process.exit(2); }
const root = parseXml(loadDocx(f).get('word/document.xml'));
const rows = [];
all(root, 'w:tbl').forEach((tbl, ti) => {
  const mar = el(el(tbl, 'w:tblPr') || { children: [] }, 'w:tblCellMar');
  const pad = mar ? ((+attr(el(mar, 'w:left') || { attrs: [] }, 'w:w') || 108) + (+attr(el(mar, 'w:right') || { attrs: [] }, 'w:w') || 108)) : 216;
  els(tbl, 'w:tr').forEach((tr, ri) => els(tr, 'w:tc').forEach((tc, ci) => {
    const w = el(el(tc, 'w:tcPr') || { children: [] }, 'w:tcW'); if (!w || attr(w, 'w:type') !== 'dxa') return;
    for (const p of els(tc, 'w:p')) {
      const jc = el(el(p, 'w:pPr') || { children: [] }, 'w:jc'); if (jc && /center|right/.test(attr(jc, 'w:val'))) continue;
      const s = textOf(p); if (!/[฀-๿]/.test(s) || s.trim().length < 6) continue;
      const sz = (all(p, 'w:sz')[0] && +attr(all(p, 'w:sz')[0], 'w:val')) || 32;
      const em = (+attr(w, 'w:w') - pad) / 20 / (sz / 2);
      const g = T.maxGap(s, em);
      if (g > +thr) rows.push([`T${ti + 1}/R${ri + 1}/C${ci + 1}`, em.toFixed(2), g.toFixed(3), s, T.wrap(s, em).map(l => l.text).join(' ⏎ ')]);
    }
  }));
});
rows.sort((a, b) => b[2] - a[2]);
const tsv = rows.map(r => r.join('\t')).join('\n');
if (out) fs.writeFileSync(out, '﻿' + tsv, 'utf8');
console.log(`stretched cells (gap > ${thr} em): ${rows.length}`);
console.log(rows.slice(0, 60).map(r => r.join(' | ')).join('\n'));
