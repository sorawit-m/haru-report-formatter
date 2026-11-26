// แทน box_strand_fix.ps1 บน Mac: กล่องโค้ด (ตาราง 1 คอลัมน์) ที่หัวกล่องค้างท้ายหน้า (บรรทัดแรกกับบรรทัดถัดไปอยู่คนละหน้า)
// -> ย่อหน้าก่อนกล่องขึ้นหน้าใหม่ (กล่องตามไปทั้งก้อน) ; ไม่ใช้ keepNext ในกล่อง (ทำให้ Word ดันทั้งกล่องจนเหลือหน้าว่างครึ่งหน้า)
// ตำแหน่งหน้าได้จาก PDF ที่ Word ส่งออก (scripts/mac/word.sh) วนสูงสุด 4 รอบ
// usage: node box_strand_fix.mjs <file.docx>
import fs from 'node:fs'; import os from 'node:os'; import path from 'node:path';
import { loadDocx, saveDocx } from '../lib/docx.mjs';
import { mapTables, gridCols, hasNestedTable, textOfXml, addPPr, getPPr } from '../lib/wordxml.mjs';
import { wordExport, pdfPages, pageIndex } from '../lib/macpdf.mjs';
const f = process.argv[2];
const tmpPdf = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'haru-box-')), 'b.pdf');
// 1) ลบ keepNext ในกล่องโค้ด
{ const d = loadDocx(f); let x = d.get('word/document.xml');
  x = mapTables(x, t => (hasNestedTable(t) || gridCols(t).length !== 1) ? t : t.replace(/<w:keepNext(?: w:val="(?:1|true|on)")?\/>/g, ''));
  d.set('word/document.xml', x); saveDocx(d, f); }
for (let round = 1; round <= 4; round++) {
  wordExport(f, tmpPdf); const idx = pageIndex(pdfPages(tmpPdf));
  const d = loadDocx(f); let x = d.get('word/document.xml'); let fixed = 0, cursor = 0;
  // กล่อง = ตาราง 1 คอลัมน์ที่มี >= 3 ย่อหน้า ; หาย่อหน้าก่อนกล่อง (นอกตาราง) แล้วใส่ pageBreakBefore
  let out = '', last = 0; const re = /<w:tbl>/g; let m;
  const boxes = []; x = mapTables(x, t => { boxes.push(t); return `\u0000BOX${boxes.length - 1}\u0000`; });
  const tokens = x.split(/\u0000BOX(\d+)\u0000/);
  // tokens: [text, idx, text, idx, ...]
  for (let i = 1; i < tokens.length; i += 2) {
    const t = boxes[+tokens[i]]; const before = tokens[i - 1];
    // เดินหาตำแหน่งข้อความในเล่ม (ข้อความก่อนกล่อง เพื่อเลื่อน cursor)
    const tailText = textOfXml(before.slice(-4000)).trim(); if (tailText) { const h = idx.find(tailText.slice(-20), cursor, 20); if (h) cursor = h.pos; }
    if (hasNestedTable(t) || gridCols(t).length !== 1) { const tt = textOfXml(t); const h = tt.trim() && idx.find(tt, cursor); if (h) cursor = h.pos + h.len; continue; }
    const paras = [...t.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map(a => textOfXml(a[0])).filter(s => s.trim());
    if (paras.length < 3) continue;
    const h1 = idx.find(paras[0], cursor); if (!h1) continue; const h2 = idx.find(paras[1], h1.pos + 1); if (!h2) continue;
    cursor = h2.pos;
    if (h1.page === h2.page) continue;
    // ย่อหน้าสุดท้ายก่อนกล่อง (ต้องอยู่นอกตาราง = ใน before โดยตรง)
    const pm = [...before.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].pop(); if (!pm) continue;
    const p = pm[0]; if (getPPr(p, 'pageBreakBefore') && !/w:val="(?:0|false|off)"/.test(getPPr(p, 'pageBreakBefore'))) continue;
    tokens[i - 1] = before.slice(0, pm.index) + addPPr(p, '<w:pageBreakBefore/>') + before.slice(pm.index + p.length); fixed++;
  }
  x = tokens.map((s, i) => i % 2 ? boxes[+s] : s).join('');
  d.set('word/document.xml', x); saveDocx(d, f);
  console.log(`round ${round}: stranded box headings moved: ${fixed}`);
  if (!fixed) break;
}
fs.rmSync(path.dirname(tmpPdf), { recursive: true, force: true });
