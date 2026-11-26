// จุดตัดที่มองไม่เห็นในคำไทยประสมยาว ตรงจุดที่ทั้งสองข้างเป็นคำในพจนานุกรม (เช่น ประชา|สัมพันธ์, หมวด|หมู่)
// --dense-back: ในตารางกว้าง (>=7 คอลัมน์) ย่อหน้าที่เป็นไทย >=30% และจัดแบบ both -> กลับเป็น thaiDistribute
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const denseBack = process.argv.includes('--dense-back');
const d = loadDocx(f); let x = d.get('word/document.xml'); const cnt = { splits: 0, back: 0 };
const seg = new Intl.Segmenter('th', { granularity: 'word' });
const MARK = /[\u0E31\u0E34-\u0E3A\u0E47-\u0E4E]/, LEAD = /[\u0E40-\u0E44]/, TH = /^[\u0E00-\u0E7F]+$/;
// a piece counts as a word only if the book itself uses it as a separate word at least twice
const freq = new Map(); for (const m of x.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)) for (const s of seg.segment(m[1].replace(/\u200B/g, ' '))) if (s.isWordLike && /^[\u0E00-\u0E7F]+$/.test(s.segment)) freq.set(s.segment, (freq.get(s.segment) || 0) + 1);
const isWord = s => { if (s.replace(/[\u0E31\u0E34-\u0E3A\u0E47-\u0E4E]/g, '').length < 2) return false; const p = [...seg.segment(s + s)].map(q => q.segment); return p.length === 2 && p[0] === s && p[1] === s; }; // dictionary word: 'ww' segments into [w, w]
const cache = new Map();
function split(w) {
  if (w.length < 6 || !TH.test(w)) return w;
  if (cache.has(w)) return cache.get(w);
  let best = -1, bestD = 1e9;
  for (let i = 2; i <= w.length - 2; i++) {
    if (MARK.test(w[i]) || LEAD.test(w[i - 1]) || w[i] === '\u0E46') continue;  // never split a cluster, after a leading vowel, or before ๆ
    if (isWord(w.slice(0, i)) && isWord(w.slice(i))) { const dd = Math.abs(i - w.length / 2); if (dd < bestD) { bestD = dd; best = i; } }
  }
  const r = best < 0 ? w : split(w.slice(0, best)) + '\u200B' + split(w.slice(best));
  if (best >= 0) cnt.splits++;
  cache.set(w, r); return r;
}
const b = x.indexOf('<w:body>'); const ce = x.indexOf('</w:p>', x.indexOf('<w:sectPr', b)) + 6;
let body = x.slice(ce); const saved = [];
body = body.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => { const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g !== 1) return t; saved.push(t); return `\u0000C${saved.length - 1}\u0000`; });
body = body.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => /<w:instrText|<w:fldSimple/.test(p) ? p :
  p.replace(/(<w:t(?: [^>]*)?>)([^<]*)(<\/w:t>)/g, (a, o, t, c) => o + t.split('\u200B').map(part => [...seg.segment(part)].map(s => s.isWordLike ? split(s.segment) : s.segment).join('')).join('\u200B') + c));
if (denseBack) body = body.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g < 7) return t;
  return t.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => { if (!/<w:jc w:val="both"\/>/.test(p)) return p; const s = [...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').replace(/[\u200B\s]/g, ''); if (!s || (s.match(/[\u0E00-\u0E7F]/g) || []).length / s.length < 0.3) return p; cnt.back++; return p.replace(/<w:jc w:val="both"\/>/, '<w:jc w:val="thaiDistribute"/>'); });
});
body = body.replace(/\u0000C(\d+)\u0000/g, (m, i) => saved[+i]);
d.set('word/document.xml', x.slice(0, ce) + body); saveDocx(d, f); console.log(JSON.stringify(cnt));
