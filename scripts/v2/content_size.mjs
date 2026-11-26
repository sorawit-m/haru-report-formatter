// ขนาดตัวหนังสือเนื้อหาทั้งเล่ม = SIZE pt (ในและนอกตาราง รวมโค้ด) — ไม่แตะปก หัวข้อ สารบัญ
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const [f, SIZE = '16'] = process.argv.slice(2); const HP = String(+SIZE * 2);
const d = loadDocx(f); let x = d.get('word/document.xml'); let st = d.get('word/styles.xml');
const sty = {}; for (const m of st.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[\s\S]*?<\/w:style>/g)) sty[m[1]] = { name: (m[0].match(/<w:name w:val="([^"]+)"/) || [])[1] || '', ol: /<w:outlineLvl w:val="[0-8]"/.test(m[0]), based: (m[0].match(/<w:basedOn w:val="([^"]+)"/) || [])[1], para: /w:type="paragraph"/.test(m[0]) };
const isHead = id => { let s = id, n = 0; while (s && sty[s] && n++ < 8) { if (sty[s].ol || /heading|toc|title|หัวเรื่อง|หัวข้อ/i.test(sty[s].name)) return true; s = sty[s].based; } return false; };
const setSz = rpr => { rpr = rpr.replace(/<w:sz w:val="\d+"\/>/, `<w:sz w:val="${HP}"/>`).replace(/<w:szCs w:val="\d+"\/>/, `<w:szCs w:val="${HP}"/>`); return rpr; };
const cnt = { runs: 0, styles: 0 };
// 1) styles: docDefaults + every non-heading paragraph/character/table style that sets a size
// header/footer styles keep their current size: pin it before the default changes
const oldDef = ((st.match(/<w:rPrDefault>[\s\S]*?<w:sz w:val="(\d+)"/) || [])[1]) || '20';
const pinned = s => /^(header|footer)$|หัวกระดาษ|ท้ายกระดาษ/i.test(s);
st = st.replace(/<w:style\b[^>]*w:styleId="([^"]+)"[\s\S]*?<\/w:style>/g, (m, id) => {
  if (!pinned((sty[id] || {}).name || '') || /<w:sz w:val=/.test(m)) return m;
  let cur = oldDef, s = sty[id].based, k = 0; while (s && k++ < 8) { const sm = st.match(new RegExp('<w:style\\b[^>]*w:styleId="' + s + '"[\\s\\S]*?</w:style>')); const v = sm && (sm[0].match(/<w:sz w:val="(\d+)"/) || [])[1]; if (v) { cur = v; break; } s = (sty[s] || {}).based; }
  const tag = `<w:sz w:val="${cur}"/><w:szCs w:val="${cur}"/>`;
  return /<w:rPr>/.test(m) ? m.replace(/<w:rPr>/, `<w:rPr>${tag}`) : m.replace(/<\/w:style>$/, `<w:rPr>${tag}</w:rPr></w:style>`);
});
st = st.replace(/(<w:rPrDefault>[\s\S]*?<\/w:rPrDefault>)/, m => setSz(m));
st = st.replace(/<w:style\b[^>]*w:styleId="([^"]+)"[\s\S]*?<\/w:style>/g, (m, id) => { if (isHead(id) || !/<w:sz w:val=/.test(m) || /^(header|footer)$|หัวกระดาษ|ท้ายกระดาษ/i.test((sty[id] || {}).name || '')) return m; cnt.styles++; return setSz(m); });
// 2) runs in body (after cover), except heading/TOC paragraphs
const b = x.indexOf('<w:body>'); const ce = x.indexOf('</w:p>', x.indexOf('<w:sectPr', b)) + 6;
let body = x.slice(ce).replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
  const sid = (p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || 'Normal';
  if (isHead(sid) || /<w:outlineLvl w:val="[0-8]"/.test(p)) return p;
  return p.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, r => { if (!/<w:sz w:val="(\d+)"/.test(r) || r.includes(`<w:sz w:val="${HP}"/>`)) return r; cnt.runs++; return setSz(r); });
});
d.set('word/document.xml', x.slice(0, ce) + body); d.set('word/styles.xml', st); saveDocx(d, f);
console.log(JSON.stringify(cnt));
