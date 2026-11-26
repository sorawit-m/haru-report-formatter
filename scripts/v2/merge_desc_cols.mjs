// ตารางทดสอบ 9 คอลัมน์: คอลัมน์ข้อความบรรยาย (Feature, Description, Expected Result) ที่ขึ้นบรรทัดใหม่เอง (หลายย่อหน้า หรือ <w:br/>)
// -> รวมเป็นย่อหน้าเดียวต่อกัน แล้วชิดซ้าย-ขวา (บรรทัดจะเต็มจริง ไม่ใช่บรรทัดสุดท้ายทุกบรรทัด)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const cnt = { cells: 0, brs: 0 };
const COLS = new Set(['feature', 'description', 'expectedresult']);
const norm = s => s.replace(/[\s​]+/g, '').toLowerCase();
const shaded = tc => /<w:shd [^>]*w:fill="(?!auto|FFFFFF|ffffff)[0-9A-Fa-f]{6}"/.test(tc);
const setJ = p => /<w:jc w:val="[^"]+"\/>/.test(p) ? p.replace(/<w:jc w:val="[^"]+"\/>/, '<w:jc w:val="thaiDistribute"/>') : (/<w:pPr>/.test(p) ? p.replace(/<\/w:pPr>/, '<w:jc w:val="thaiDistribute"/></w:pPr>') : p.replace(/^(<w:p(?: [^>]*)?>)/, '$1<w:pPr><w:jc w:val="thaiDistribute"/></w:pPr>'));
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g < 7) return t;
  const rows = [...t.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(m => m[0]); if (!rows.length) return t;
  const hdr = [...rows[0].matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(c => norm([...c[1].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('')));
  let want = hdr.map(h => COLS.has(h));
  if (!want.some(Boolean)) want = [false, false, true, true, false, false, false, true, false]; // content-only table: standard 9-column layout
  return t.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, tr => {
    const tcs = [...tr.matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(m => m[1]); if (tcs.length < 7 || tcs.every(shaded)) return tr;
    let ci = 0;
    return tr.replace(/<w:tc>([\s\S]*?)<\/w:tc>/g, (all, inner) => {
      const span = +((inner.match(/<w:gridSpan w:val="(\d+)"/) || [])[1] || 1); const here = ci; ci += span;
      if (!want[here] || /<w:drawing|<w:numPr>/.test(inner)) return all;
      let changed = false;
      inner = inner.replace(/<w:br\/>/g, () => { cnt.brs++; changed = true; return '<w:t xml:space="preserve"> </w:t>'; });
      const ps = [...inner.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map(m => m[0]).filter(p => /<w:t[ >]/.test(p));
      const allPs = [...inner.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map(m => m[0]);
      if (ps.length > 1) {
        const runs = p => (p.match(/^<w:p(?: [^>]*)?>(?:<w:pPr>[\s\S]*?<\/w:pPr>)?([\s\S]*)<\/w:p>$/) || [])[1] || '';
        const rPr = ((ps[0].match(/<w:r(?: [^>]*)?>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/) || [])[1]) || '';
        const merged = ps[0].replace(/<\/w:p>$/, '') + ps.slice(1).map(p => `<w:r>${rPr}<w:t xml:space="preserve"> </w:t></w:r>` + runs(p)).join('') + '</w:p>';
        const s = inner.indexOf(allPs[0]), e = inner.lastIndexOf(allPs[allPs.length - 1]) + allPs[allPs.length - 1].length;
        inner = inner.slice(0, s) + merged + inner.slice(e); changed = true;
      }
      inner = inner.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, setJ);
      if (changed) cnt.cells++;
      return '<w:tc>' + inner + '</w:tc>';
    });
  });
});
d.set('word/document.xml', x); saveDocx(d, f); console.log(JSON.stringify(cnt));
