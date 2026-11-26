// ช่อง Step ในตารางทดสอบ (หัว "Step"): ข้อความ "1.xxx 2.xxx 3.xxx" -> แยกเป็นข้อละย่อหน้า เยื้องแบบแขวน (บรรทัดต่อไปตรงกับข้อความ ไม่ใช่ตัวเลข)
// ข้อความทุกตัวอักษรเหมือนเดิม แค่ขึ้นบรรทัดใหม่ก่อนเลขข้อ
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const [f, colName = 'Step'] = process.argv.slice(2);
const d = loadDocx(f); let x = d.get('word/document.xml'); const cnt = { cells: 0, items: 0 };
const HANG = 284; // 0.5 cm
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  const rows = [...t.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(m => m[0]); if (!rows.length) return t;
  const hdr = [...rows[0].matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(c => [...c[1].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').replace(/​/g, '').trim());
  let col = hdr.indexOf(colName);
  // content-only tables (header in a separate table): same 9-column layout -> Step is column 6
  if (col < 0) { const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g === 9) col = 5; else return t; }
  return t.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, tr => { let ci = 0; return tr.replace(/<w:tc>([\s\S]*?)<\/w:tc>/g, (all, inner) => {
    const span = +((inner.match(/<w:gridSpan w:val="(\d+)"/) || [])[1] || 1); const here = ci; ci += span;
    if (here !== col) return all;
    const ps = [...inner.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map(m => m[0]); if (!ps.length) return all;
    const text = ps.map(p => [...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('')).join(' ');
    const plain = text.replace(/​/g, '');
    if (!/^\s*1\s*\./.test(plain) || !/\s2\s*\.|\S2\s*\.\D/.test(plain)) return all;
    // split before "N." (N = 2..20) when followed by a non-digit — keeps decimals like 1.5 intact
    const items = text.split(/\s*(?=(?<![\d.])(?:[2-9]|1\d|20)​?\.(?!\d))/).map(s => s.trim().replace(/^(\d{1,2})\u200B?\.\s*/, '$1. ')).filter(Boolean);
    if (items.length < 2) return all;
    const p0 = ps[0];
    let pPr = (p0.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || ['<w:pPr></w:pPr>'])[0].replace(/<w:ind [^>]*\/>/, '');
    const ind = `<w:ind w:left="${HANG}" w:hanging="${HANG}"/>`;
    const k = pPr.search(/<w:(?:jc|textDirection|textAlignment|outlineLvl|rPr|sectPr|pPrChange)\b/);
    pPr = k < 0 ? pPr.replace('</w:pPr>', ind + '</w:pPr>') : pPr.slice(0, k) + ind + pPr.slice(k);
    const rPr = (p0.match(/<w:r(?: [^>]*)?>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/) || [])[1] || '';
    const esc = s => s; // text is already XML-escaped
    const newPs = items.map(it => `<w:p>${pPr}<w:r>${rPr}<w:t xml:space="preserve">${esc(it)}</w:t></w:r></w:p>`).join('');
    const s = inner.indexOf(ps[0]), e = inner.lastIndexOf(ps[ps.length - 1]) + ps[ps.length - 1].length;
    cnt.cells++; cnt.items += items.length;
    return '<w:tc>' + inner.slice(0, s) + newPs + inner.slice(e) + '</w:tc>';
  }); });
});
d.set('word/document.xml', x); saveDocx(d, f); console.log(JSON.stringify(cnt));
