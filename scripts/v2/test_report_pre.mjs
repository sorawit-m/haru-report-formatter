// ก่อนจัดเล่ม (เล่มรายการทดสอบ): 1) ลบแถว "Test by" / "Approve by" ในตาราง  2) เว้นวรรคหลัง "รูปที่" / "ตารางที่" ก่อนเลข
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const cnt = { rowsRemoved: 0, captionSpaces: 0 };
const txt = s => [...s.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').replace(/\u200B/g, '').trim();
x = x.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, tr => {
  const first = (tr.match(/<w:tc>([\s\S]*?)<\/w:tc>/) || [, ''])[1];
  if (/^(Test(ed)?\s*by|Approve(d)?\s*by)\b/i.test(txt(first))) { cnt.rowsRemoved++; return ''; }
  return tr;
});
x = x.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
  if (/TOC \\|PAGEREF/.test(p)) return p;
  return p.replace(/<w:t((?: [^>]*)?)>([^<]*)<\/w:t>/g, (a, attrs, t) => {
    let n = t.replace(/(รูปที่|ตารางที่)(?=\d)/g, m => { cnt.captionSpaces++; return m + ' '; });
    if (/(รูปที่|ตารางที่)$/.test(n)) { n += ' '; cnt.captionSpaces++; }
    if (n === t) return a;
    return `<w:t xml:space="preserve">${n}</w:t>`;
  });
});
d.set('word/document.xml', x); saveDocx(d, f); console.log(JSON.stringify(cnt));
