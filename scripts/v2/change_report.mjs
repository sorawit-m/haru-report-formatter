// เทียบต้นฉบับกับเล่มที่จัดแล้ว -> รายการสิ่งที่เปลี่ยนแบบละเอียด (หมวด + จำนวน + ตัวอย่างพร้อมข้อความก่อน/หลัง)
// usage: node change_report.mjs <original.docx> <result.docx> <out.json>
import { loadDocx } from '../lib/docx.mjs';
import fs from 'node:fs';
const [origF, resF, outF] = process.argv.slice(2);
const norm = s => s.replace(/[​⁠\s]+/g, '');
function paras(f) {
  const x = loadDocx(f).get('word/document.xml'); const out = []; let depth = 0, tbl = 0;
  for (const m of x.matchAll(/<w:tbl>|<\/w:tbl>|<w:p[ >][\s\S]*?<\/w:p>/g)) {
    const p = m[0]; if (p === '<w:tbl>') { depth++; tbl++; continue; } if (p === '</w:tbl>') { depth--; continue; }
    if (/PAGEREF|<w:instrText[^>]*>\s*TOC/.test(p)) continue;
    const raw = [...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('');
    if (!raw.trim()) continue;
    const jc = (p.match(/<w:pPr>[\s\S]*?<w:jc w:val="([^"]+)"/) || [])[1] || '';
    const sz = (p.match(/<w:r[ >](?:(?!<\/w:r>)[\s\S])*?<w:sz w:val="(\d+)"/) || [])[1] || '';
    const bold = /<w:r[ >](?:(?!<\/w:r>)[\s\S])*?<w:b\/>/.test(p);
    const font = (p.match(/w:ascii="([^"]+)"/) || [])[1] || '';
    const sp = /<w:spacing w:val="-?[1-9]/.test(p);
    out.push({ raw, n: norm(raw), jc, sz, bold, font, sp, inTable: depth > 0 });
  }
  return out;
}
const A = paras(origF), B = paras(resF);
const bIdx = new Map(); B.forEach((p, i) => { if (!bIdx.has(p.n)) bIdx.set(p.n, i); });
const aSet = new Set(A.map(p => p.n));
const cats = {};
const canon = s => s.replace(/InPut/g, 'Input').replace(/\d+/g, '#');
const bCanon = new Map(); B.forEach((p, i) => { const k = canon(p.n); if (!bCanon.has(k)) bCanon.set(k, i); });
const usedB = new Set();
const add = (key, label, ex) => { (cats[key] ||= { label, count: 0, examples: [] }).count++; if (cats[key].examples.length < 8) cats[key].examples.push(ex); };
const jcName = j => ({ thaiDistribute: 'ชิดซ้าย-ขวา (ไทย)', both: 'ชิดซ้าย-ขวา', left: 'ชิดซ้าย', center: 'กึ่งกลาง', right: 'ชิดขวา', '': 'ตามสไตล์' }[j] || j);
for (const a of A) {
  const i = bIdx.get(a.n);
  if (i === undefined) {
    // merged into a longer paragraph?
    const ci = bCanon.get(canon(a.n)); if (ci !== undefined) { usedB.add(ci); add('changedText', 'ข้อความที่ถูกแก้ (เลขบท / คำบรรยาย / คำสะกด)', { find: B[ci].raw, before: a.raw.replace(/\u200B/g, '').slice(0, 70), after: B[ci].raw.replace(/\u200B/g, '').slice(0, 70) }); continue; }
    const host = (a.n.length >= 6 || a.inTable) ? B.find(b => b.inTable === a.inTable && b.n.length > a.n.length && b.n.includes(a.n)) : null;
    if (host) add('merged', 'รวมข้อความที่ขาดเป็นท่อนให้ต่อกัน', { find: host.raw, before: a.raw.slice(0, 60), after: host.raw.replace(/​/g, '').slice(0, 90) });
    else if (/Test(ed)? ?by|Approve(d)? ?by/i.test(a.raw)) add('removedRows', 'ลบแถว Test by / Approve by', { find: '', before: a.raw.slice(0, 60), after: '(ลบ)' });
    else if (/บทนำ|ความเป็นมา|วัตถุประสงค์/.test(a.raw) || !a.inTable) add('removedText', 'ข้อความที่ถูกเอาออก (บทนำ ฯลฯ)', { find: '', before: a.raw.slice(0, 80), after: '(เอาออก)' });
    else add('removedText', 'ข้อความที่ถูกเอาออก (บทนำ ฯลฯ)', { find: '', before: a.raw.slice(0, 80), after: '(เอาออก)' });
    continue;
  }
  const b = B[i];
  if (a.jc !== b.jc && b.jc) add('align', 'การจัดข้อความ', { find: b.raw, before: jcName(a.jc), after: jcName(b.jc), text: b.raw.replace(/​/g, '').slice(0, 60) });
  if (a.sz && b.sz && a.sz !== b.sz) add('size', 'ขนาดตัวอักษร', { find: b.raw, before: a.sz / 2 + ' pt', after: b.sz / 2 + ' pt', text: b.raw.replace(/​/g, '').slice(0, 60) });
  if (!a.bold && b.bold && b.inTable) add('bold', 'หัวตารางเป็นตัวหนา', { find: b.raw, before: 'ปกติ', after: 'ตัวหนา', text: b.raw.replace(/​/g, '').slice(0, 40) });
  if (a.font && b.font && a.font !== b.font) add('font', 'เปลี่ยนฟอนต์ให้เป็นฟอนต์เดียว', { find: b.raw, before: a.font, after: b.font, text: b.raw.replace(/​/g, '').slice(0, 60) });
  if (a.sp && !b.sp) add('spacing', 'เอาการบีบ/ถ่างตัวอักษรที่ตั้งเองออก', { find: b.raw, before: 'บีบ/ถ่าง', after: 'ปกติ', text: b.raw.replace(/​/g, '').slice(0, 60) });
}
fs.writeFileSync(outF, JSON.stringify(Object.values(cats), null, 1), 'utf8');
console.log(Object.values(cats).map(c => `${c.label}: ${c.count}`).join(' | '));
