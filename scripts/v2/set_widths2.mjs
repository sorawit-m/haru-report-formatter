// ตั้งความกว้างคอลัมน์ (cm) ให้ทุกตารางที่หัวตรงกัน — เฉพาะตารางที่ทุกแถวมีจำนวนช่อง (รวม gridSpan) เท่ากับจำนวนคอลัมน์ ไม่งั้นข้าม
// usage: node set_widths2.mjs file.docx "h1|h2|h3=w1,w2,w3" ["..."]
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const [f, ...specArgs] = process.argv.slice(2);
const norm = s => s.replace(/[\s\u200B]+/g, '').toLowerCase();
const specs = specArgs.map(a => { const [h, w] = a.split('='); return { hdr: h.split('|').map(norm), cm: w.split(',').map(Number), label: h, n: 0, skipped: 0 }; });
const d = loadDocx(f); let x = d.get('word/document.xml');
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  if (/<w:tbl>[\s\S]*<w:tbl>/.test(t)) return t;
  const rows = [...t.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(m => m[0]); if (!rows.length) return t;
  const cellsOf = tr => [...tr.matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(c => c[1]);
  const hdr = cellsOf(rows[0]).map(c => norm([...c.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('')));
  const sp = specs.find(s => s.hdr.length === hdr.length && s.hdr.every((h, i) => h === hdr[i])); if (!sp) return t;
  const cols = sp.cm.length;
  if (rows.some(tr => cellsOf(tr).reduce((a, c) => a + +((c.match(/<w:gridSpan w:val="(\d+)"/) || [])[1] || 1), 0) !== cols)) { sp.skipped++; return t; }
  const gridCols = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (gridCols !== cols) { sp.skipped++; return t; }
  const tw = sp.cm.map(c => Math.round(c * 567)); const total = tw.reduce((a, b) => a + b, 0);
  t = t.replace(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/, `<w:tblGrid>${tw.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`);
  t = /<w:tblW [^>]*\/>/.test(t) ? t.replace(/<w:tblW [^>]*\/>/, `<w:tblW w:w="${total}" w:type="dxa"/>`) : t.replace(/<w:tblPr>/, `<w:tblPr><w:tblW w:w="${total}" w:type="dxa"/>`);
  t = /<w:tblInd [^>]*\/>/.test(t) ? t.replace(/<w:tblInd [^>]*\/>/, '<w:tblInd w:w="0" w:type="dxa"/>') : t;
  t = /<w:tblLayout [^>]*\/>/.test(t) ? t.replace(/<w:tblLayout [^>]*\/>/, '<w:tblLayout w:type="fixed"/>') : t.replace(/(<w:tblW [^>]*\/>)/, '$1<w:tblLayout w:type="fixed"/>');
  t = t.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, tr => { let ci = 0; return tr.replace(/<w:tc>([\s\S]*?)<\/w:tc>/g, (all, inner) => {
    const span = +((inner.match(/<w:gridSpan w:val="(\d+)"/) || [])[1] || 1); const w = tw.slice(ci, ci + span).reduce((a, b) => a + b, 0); ci += span;
    const tcW = `<w:tcW w:w="${w}" w:type="dxa"/>`;
    inner = /<w:tcW [^>]*\/>/.test(inner) ? inner.replace(/<w:tcW [^>]*\/>/, tcW) : /<w:tcPr>/.test(inner) ? inner.replace(/<w:tcPr>/, `<w:tcPr>${tcW}`) : `<w:tcPr>${tcW}</w:tcPr>` + inner;
    return '<w:tc>' + inner + '</w:tc>'; }); });
  sp.n++; return t;
});
d.set('word/document.xml', x); saveDocx(d, f);
for (const s of specs) console.log(`${s.label}: ${s.n} tables set to ${s.cm.join(' / ')} cm${s.skipped ? `, skipped ${s.skipped} (irregular)` : ''}`);
