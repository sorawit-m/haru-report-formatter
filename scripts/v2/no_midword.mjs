// ไม่ให้ Word ตัดกลางคำไทย: เอาจุดตัดคำประสมออก (ครุ|ภัณฑ์ -> ครุภัณฑ์) และใส่ U+2060 (word joiner) ระหว่างตัวอักษรภายในคำ
// ตัดบรรทัดได้เฉพาะระหว่างคำ (U+200B / เว้นวรรค) — ไม่แตะกล่องโค้ด ฟิลด์ ปก
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const [f, scope = 'all'] = process.argv.slice(2);
const d = loadDocx(f); let x = d.get('word/document.xml'); const cnt = { words: 0, joined: 0 };
const seg = new Intl.Segmenter('th', { granularity: 'word' });
const TH = /[฀-๿]/;
const b = x.indexOf('<w:body>'); const ce = x.indexOf('</w:p>', x.indexOf('<w:sectPr', b)) + 6;
let body = x.slice(ce); const saved = [];
body = body.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => { const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g === 1 || (scope === 'dense' && g < 7)) { saved.push(t); return `\u0000C${saved.length - 1}\u0000`; } return t; });
const fix = s => {
  s = s.replace(/⁠/g, '');
  // rebuild: words separated by U+200B (only between whole words from the segmenter)
  const plain = s.replace(/​/g, '');
  return [...seg.segment(plain)].map(p => {
    const w = p.segment; if (!TH.test(w) || w.length < 2) return w;
    cnt.words++;
    // join every grapheme cluster inside the word
    const cl = [...new Intl.Segmenter('th', { granularity: 'grapheme' }).segment(w)].map(g => g.segment);
    cnt.joined += cl.length - 1;
    return cl.join('⁠');
  }).reduce((acc, w, i, arr) => acc + w + (i < arr.length - 1 && TH.test(w.slice(-1)) && TH.test(arr[i + 1][0]) ? '​' : ''), '');
};
body = body.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
  if (/<w:instrText|<w:fldSimple/.test(p)) return p;
  if (scope === 'dense' && !/\u0000/.test(p)) { /* paragraphs outside tables untouched in dense mode */ }
  return p.replace(/(<w:t(?: [^>]*)?>)([^<]*)(<\/w:t>)/g, (a, o, t, c) => TH.test(t) ? o + fix(t) + c : a);
});
body = body.replace(/\u0000C(\d+)\u0000/g, (m, i) => saved[+i]);
d.set('word/document.xml', x.slice(0, ce) + body); saveDocx(d, f); console.log(JSON.stringify(cnt));
