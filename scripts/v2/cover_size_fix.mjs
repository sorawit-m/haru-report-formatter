// ปรับขนาดตัวอักษรบนปกให้เท่ากันทุกเล่ม: ชื่อเอกสาร 24 pt, บรรทัดชื่อโครงการ 18 pt — แก้แค่ขนาด ไม่แตะอย่างอื่น
import { loadDocx, saveDocx } from '../lib/docx.mjs';
import { coverEnd } from '../lib/wordxml.mjs';
const TITLE = 48, PROJECT = 36; // half-points
const setSz = (rPrInner, v) => rPrInner.replace(/<w:sz w:val="\d+"\/>|<w:szCs w:val="\d+"\/>/g, '') + `<w:sz w:val="${v}"/><w:szCs w:val="${v}"/>`;
// put sz/szCs in schema-friendly place: just before </w:rPr> is accepted by Word for these elements in practice, but keep order: insert before w:highlight/u/effect/... is not needed — Word tolerates; we instead rebuild ordering minimal: remove then append before any of the late elements
const LATE = /<w:(?:highlight|u|effect|bdr|shd|fitText|vertAlign|rtl|cs|em|lang|eastAsianLayout|specVanish|oMath)\b/;
const putSz = (inner, v) => { inner = inner.replace(/<w:sz w:val="\d+"\/>|<w:szCs w:val="\d+"\/>/g, ''); const tag = `<w:sz w:val="${v}"/><w:szCs w:val="${v}"/>`; const m = inner.match(LATE); return m ? inner.slice(0, m.index) + tag + inner.slice(m.index) : inner + tag; };
const fixRun = (r, v) => r.includes('<w:rPr>') ? r.replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/, (_, i) => `<w:rPr>${putSz(i, v)}</w:rPr>`) : r.replace(/^(<w:r(?: [^>]*)?>)/, `$1<w:rPr>${putSz('', v)}</w:rPr>`);
for (const f of process.argv.slice(2)) {
  const d = loadDocx(f); let x = d.get('word/document.xml');
  const b = x.indexOf('<w:body>'); const e = coverEnd(x);
  let n = 0;
  const cover = x.slice(b, e).replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
    const txt = [...p.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m => m[1]).join('');
    if (!txt.trim() || p.includes('<w:drawing') || p.includes('<mc:AlternateContent')) return p;
    const v = txt.includes('โครงการ') ? PROJECT : TITLE; n++;
    p = p.replace(/<w:r>[\s\S]*?<\/w:r>|<w:r [\s\S]*?<\/w:r>/g, r => r.includes('<w:t') ? fixRun(r, v) : r);
    // paragraph mark size too
    p = p.replace(/(<w:pPr>[\s\S]*?)(<\/w:pPr>)/, (m0, a, c) => /<w:rPr>[\s\S]*?<\/w:rPr>/.test(a) ? a.replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/, (_, i) => `<w:rPr>${putSz(i, v)}</w:rPr>`) + c : m0);
    return p;
  });
  x = x.slice(0, b) + cover + x.slice(e); d.set('word/document.xml', x); saveDocx(d, f);
  console.log(`${f.split('\\').pop()}: ${n} paragraphs`);
}
