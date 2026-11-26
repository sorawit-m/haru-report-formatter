// ตารางกว้าง (>=7 คอลัมน์): ช่องที่จัดแบบไทยแล้วไม่ถ่าง -> thaiDistribute ; ช่องที่จะถ่าง -> both (ขยายเฉพาะช่องว่าง)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
import { chunkWidth } from '../lib/thaiwrap.mjs';
const [f, ptS = '14'] = process.argv.slice(2); const PT = +ptS;
const d = loadDocx(f); let x = d.get('word/document.xml'); const cnt = { thai: 0, both: 0 };
const SP = chunkWidth('a a') - chunkWidth('aa');
// break opportunities = spaces + invisible break points (what Word now breaks on)
function stretches(text, widthEm) {
  const toks = text.split(/(?<=[\u200B ])/); const lines = []; let cur = '', w = 0;
  for (const tk of toks) { const t = tk.replace(/\u200B/g, ''); const tw = chunkWidth(t.replace(/ $/, '')) + (t.endsWith(' ') ? SP : 0);
    if (w + chunkWidth(t.replace(/ $/, '')) > widthEm && cur) { lines.push([cur, w]); cur = t; w = tw; } else { cur += t; w += tw; } }
  for (const [s, used] of lines) { const n = s.replace(/[\u0E31\u0E34-\u0E3A\u0E47-\u0E4E\s]/g, '').length; if ((widthEm - used) / Math.max(1, n) > 0.035) return true; }
  return false;
}
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  const grid = [...((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].matchAll(/<w:gridCol w:w="(\d+)"\/>/g))].map(a => +a[1]); if (grid.length < 7) return t;
  const marL = 160, marR = 160;
  return t.replace(/<w:tc>([\s\S]*?)<\/w:tc>/g, (all, inner) => {
    const w = +((inner.match(/<w:tcW w:w="(\d+)"/) || [])[1] || 0); if (!w) return all;
    const widthEm = (w - marL - marR) / (PT * 20);
    return '<w:tc>' + inner.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
      const jc = (p.match(/<w:pPr>[\s\S]*?<w:jc w:val="([^"]+)"\/>/) || [])[1] || 'thaiDistribute'; if (jc !== 'thaiDistribute' && jc !== 'both') return p;
      const text = [...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').replace(/&[a-z]+;/g, 'x');
      const plain = text.replace(/[\u200B\s]/g, ''); if (!plain) return p;
      const thaiShare = (plain.match(/[\u0E00-\u0E7F]/g) || []).length / plain.length;
      const want = thaiShare >= 0.3 && !stretches(text, widthEm) ? 'thaiDistribute' : 'both';
      cnt[want === 'both' ? 'both' : 'thai']++;
      if (/<w:pPr>[\s\S]*?<w:jc w:val=/.test(p)) return p.replace(/<w:jc w:val="[^"]+"\/>/, `<w:jc w:val="${want}"/>`); return /<w:pPr>/.test(p) ? p.replace(/<w:pPr>([\s\S]*?)<\/w:pPr>/, (a, i) => { const k = i.search(/<w:(?:textDirection|textAlignment|outlineLvl|rPr|sectPr|pPrChange)\b/); const tag = `<w:jc w:val="${want}"/>`; return '<w:pPr>' + (k < 0 ? i + tag : i.slice(0, k) + tag + i.slice(k)) + '</w:pPr>'; }) : p.replace(/^(<w:p(?: [^>]*)?>)/, `<w:pPr><w:jc w:val="${want}"/></w:pPr>`);
    }) + '</w:tc>';
  });
});
d.set('word/document.xml', x); saveDocx(d, f); console.log(JSON.stringify(cnt));
