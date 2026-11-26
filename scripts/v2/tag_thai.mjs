// ติดป้ายข้อความไทยเป็นภาษาไทย (w:cs + lang bidi th-TH) ให้ Word ตัดบรรทัดตามพจนานุกรมไทย ไม่ตัดกลางคำ
// --no-compound: เอาจุดตัดคำประสมที่เคยใส่ (ภายในคำเดียวของพจนานุกรม) ออก
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const noCompound = process.argv.includes('--no-compound');
const d = loadDocx(f); let x = d.get('word/document.xml'); const cnt = { tagged: 0, unsplit: 0 };
const TH = /[฀-๿]/; const seg = new Intl.Segmenter('th', { granularity: 'word' });
const isDict = w => { const p = [...seg.segment(w + w)].map(s => s.segment); return p.length === 2 && p[0] === w && p[1] === w; };
const b = x.indexOf('<w:body>'); const ce = x.indexOf('</w:p>', x.indexOf('<w:sectPr', b)) + 6;
let body = x.slice(ce).replace(/⁠/g, '');
body = body.replace(/<w:r(?: [^>]*)?>[\s\S]*?<\/w:r>/g, r => {
  if (!/<w:t[ >]/.test(r)) return r;
  const txt = [...r.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join(''); if (!TH.test(txt)) return r;
  if (noCompound) r = r.replace(/(<w:t(?: [^>]*)?>)([^<]*)(<\/w:t>)/g, (a, o, t, c) => o + t.replace(/[฀-๿]+(?:​[฀-๿]+)+/g, m => {
    // merge neighbours whose join is itself one dictionary word (ครุ|ภัณฑ์ -> ครุภัณฑ์)
    const parts = m.split('​'); const out = [parts[0]];
    for (let i = 1; i < parts.length; i++) { const j = out[out.length - 1] + parts[i]; if (isDict(j)) { out[out.length - 1] = j; cnt.unsplit++; } else out.push(parts[i]); }
    return out.join('​'); }) + c);
  cnt.tagged++;
  const tags = '<w:cs/><w:lang w:bidi="th-TH"/>';
  if (/<w:rPr>/.test(r)) return r.replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/, (a, i) => { i = i.replace(/<w:cs\/>/g, '').replace(/<w:lang [^>]*\/>/g, ''); const k = i.search(/<w:(?:em|eastAsianLayout|specVanish|oMath)\b/); return '<w:rPr>' + (k < 0 ? i + tags : i.slice(0, k) + tags + i.slice(k)) + '</w:rPr>'; });
  return r.replace(/^(<w:r(?: [^>]*)?>)/, `$1<w:rPr>${tags}</w:rPr>`);
});
d.set('word/document.xml', x.slice(0, ce) + body); saveDocx(d, f); console.log(JSON.stringify(cnt));
