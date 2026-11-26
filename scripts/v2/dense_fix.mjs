// ตารางกว้าง (9 คอลัมน์): 1) ทุกตารางที่มี 9 คอลัมน์ ได้ความกว้างเท่ากัน (รวมตารางเนื้อหาที่แยกจากหัวตาราง)
// 2) ข้อความในช่องที่จัดแบบ thaiDistribute -> both (ไม่ถ่างตัวอักษร)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const [f, wcm] = process.argv.slice(2); const W = wcm.split(',').map(Number).map(c => Math.round(c * 567)); const COLS = W.length; const total = W.reduce((a, b) => a + b, 0);
const d = loadDocx(f); let x = d.get('word/document.xml'); const st = d.get('word/styles.xml'); const cnt = { widths: 0, both: 0, skipped: 0 };
const sty = {}; for (const m of st.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[\s\S]*?<\/w:style>/g)) sty[m[1]] = { jc: (m[0].match(/<w:jc w:val="([^"]+)"/) || [])[1], based: (m[0].match(/<w:basedOn w:val="([^"]+)"/) || [])[1], def: /w:default="1"/.test(m[0]) && /w:type="paragraph"/.test(m[0]) };
const defId = Object.keys(sty).find(k => sty[k].def) || 'Normal';
const styJc = id => { let s = id, k = 0; while (s && sty[s] && k++ < 8) { if (sty[s].jc) return sty[s].jc; s = sty[s].based; } return 'left'; };
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  if (/<w:tbl>[\s\S]*<w:tbl>/.test(t)) return t;
  const grid = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (grid !== COLS) return t;
  const rows = [...t.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(m => m[0]);
  const ok = rows.every(tr => [...tr.matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].reduce((a, c) => a + +((c[1].match(/<w:gridSpan w:val="(\d+)"/) || [])[1] || 1), 0) === COLS);
  if (!ok) { cnt.skipped++; return t; }
  t = t.replace(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/, `<w:tblGrid>${W.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`);
  t = t.replace(/<w:tblW [^>]*\/>/, `<w:tblW w:w="${total}" w:type="dxa"/>`).replace(/<w:tblInd [^>]*\/>/, '<w:tblInd w:w="0" w:type="dxa"/>');
  t = t.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, tr => { let ci = 0; return tr.replace(/<w:tc>([\s\S]*?)<\/w:tc>/g, (all, inner) => { const span = +((inner.match(/<w:gridSpan w:val="(\d+)"/) || [])[1] || 1); const w = W.slice(ci, ci + span).reduce((a, b) => a + b, 0); ci += span; return '<w:tc>' + inner.replace(/<w:tcW [^>]*\/>/, `<w:tcW w:w="${w}" w:type="dxa"/>`) + '</w:tc>'; }); });
  cnt.widths++;
  t = t.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
    const own = (p.match(/<w:pPr>[\s\S]*?<w:jc w:val="([^"]+)"\/>/) || [])[1];
    const eff = own || styJc((p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || defId);
    if (eff !== 'thaiDistribute') return p; cnt.both++;
    if (own) return p.replace(/(<w:pPr>[\s\S]*?)<w:jc w:val="thaiDistribute"\/>/, '$1<w:jc w:val="both"/>');
    return /<w:pPr>/.test(p) ? p.replace(/<w:pPr>([\s\S]*?)<\/w:pPr>/, (a, i) => { const k = i.search(/<w:(?:textDirection|textAlignment|outlineLvl|rPr|sectPr|pPrChange)\b/); return '<w:pPr>' + (k < 0 ? i + '<w:jc w:val="both"/>' : i.slice(0, k) + '<w:jc w:val="both"/>' + i.slice(k)) + '</w:pPr>'; }) : p.replace(/^(<w:p(?: [^>]*)?>)/, '$1<w:pPr><w:jc w:val="both"/></w:pPr>');
  });
  return t;
});
d.set('word/document.xml', x); saveDocx(d, f); console.log(JSON.stringify(cnt));
