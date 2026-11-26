// จัดข้อความแบบเล่ม (1) ในขั้นเดียว โดย "กันโค้ด" ไว้ทั้งหมด:
//   โค้ด = ตาราง 1 คอลัมน์ (กล่องตัวอย่าง Request/Response) + run ที่ใช้ฟอนต์โค้ด
// 1) รวมบรรทัดที่ขาดในช่องตาราง (ช่องที่ไม่ใช่โค้ด/ลำดับเลข)  2) จุดตัดคำไทย  3) จุดตัดในคำอังกฤษยาว/URL  4) จุดตัดหลัง / ที่ติดคำไทย
// 5) ทุกย่อหน้าเนื้อหา = thaiDistribute (ยกเว้นหัวข้อ ฟิลด์ รูป โค้ด ย่อหน้ากลาง/ขวา แถวหัวตารางพื้นสี)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const st = d.get('word/styles.xml');
const CODEFONT = /w:ascii="(?:Consolas|Courier New|Courier|Menlo|Monaco|Lucida Console|Source Code Pro)"/;
const cnt = { codeBoxes: 0, merged: 0, thaiBreaks: 0, latinBreaks: 0, slashBreaks: 0, justified: 0, skip: {} };
const sk = k => cnt.skip[k] = (cnt.skip[k] || 0) + 1;
// styles
const sty = {}; for (const m of st.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[\s\S]*?<\/w:style>/g)) sty[m[1]] = { name: (m[0].match(/<w:name w:val="([^"]+)"/) || [])[1] || '', ol: /<w:outlineLvl w:val="[0-8]"/.test(m[0]), jc: (m[0].match(/<w:jc w:val="([^"]+)"/) || [])[1], based: (m[0].match(/<w:basedOn w:val="([^"]+)"/) || [])[1] };
const isHead = id => { let s = id, n = 0; while (s && sty[s] && n++ < 8) { if (sty[s].ol || /heading|toc|caption|title|หัวเรื่อง|หัวข้อ/i.test(sty[s].name)) return true; s = sty[s].based; } return false; };
const styJc = id => { let s = id, n = 0; while (s && sty[s] && n++ < 8) { if (sty[s].jc) return sty[s].jc; s = sty[s].based; } return null; };
// split off cover
const b = x.indexOf('<w:body>'); const ce = x.indexOf('</w:p>', x.indexOf('<w:sectPr', b)) + 6;
let body = x.slice(ce);
// protect code boxes (1-column tables)
const saved = [];
body = body.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  if (/<w:tbl>[\s\S]*<w:tbl>/.test(t)) return t;
  const cols = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length;
  if (cols !== 1) return t;
  cnt.codeBoxes++; saved.push(t); return `\u0000CODE${saved.length - 1}\u0000`;
});
// 1) merge split lines inside cells
body = body.replace(/<w:tc>([\s\S]*?)<\/w:tc>/g, (all, inner) => {
  const ps = [...inner.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map(a => a[0]); if (ps.length < 2) return all;
  if (ps.some(p => CODEFONT.test(p) || /<w:numPr>/.test(p) || !/<w:t[ >]/.test(p) || /<w:drawing|<w:pict|<w:instrText/.test(p))) return all;
  const runs = p => (p.match(/^<w:p(?: [^>]*)?>(?:<w:pPr>[\s\S]*?<\/w:pPr>)?([\s\S]*)<\/w:p>$/) || [])[1] || '';
  const rPr = ((ps[0].match(/<w:r(?: [^>]*)?>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/) || [])[1]) || '';
  const merged = ps[0].replace(/<\/w:p>$/, '') + ps.slice(1).map(p => `<w:r>${rPr}<w:t xml:space="preserve"> </w:t></w:r>` + runs(p)).join('') + '</w:p>';
  const s = inner.indexOf(ps[0]), e = inner.lastIndexOf(ps[ps.length - 1]) + ps[ps.length - 1].length;
  cnt.merged++; return '<w:tc>' + inner.slice(0, s) + merged + inner.slice(e) + '</w:tc>';
});
// 2-4) invisible break points (text runs only; not fields, not code-font runs)
const seg = new Intl.Segmenter('th', { granularity: 'word' }); const TH = /[\u0E00-\u0E7F]/;
const thaiBrk = s => { if (!TH.test(s)) return s; const parts = [...seg.segment(s)].map(p => p.segment); let out = ''; for (let i = 0; i < parts.length; i++) { out += parts[i]; const a = parts[i], n = parts[i + 1]; if (n && TH.test(a.slice(-1)) && TH.test(n[0])) { out += '\u200B'; cnt.thaiBreaks++; } } return out; };
const latinFix = tok => tok.replace(/([_\/.])(?=[A-Za-z{\[]|[0-9](?![0-9.]*$))/g, (m, c, off) => { if (c === '/' && (tok[off - 1] === '/' || tok[off - 1] === ':')) return m; cnt.latinBreaks++; return c + '\u200B'; });
body = body.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
  if (/<w:instrText|<w:fldSimple/.test(p)) return p;
  return p.replace(/<w:r(?: [^>]*)?>[\s\S]*?<\/w:r>/g, r => CODEFONT.test(r) ? r : r.replace(/(<w:t(?: [^>]*)?>)([^<]*)(<\/w:t>)/g, (a, o, t, c) => {
    t = t.replace(/\u200B/g, '');
    if (!/https?:|\{/.test(t)) t = thaiBrk(t);
    t = t.replace(/[A-Za-z0-9_.\/:\-{}*]{13,}/g, latinFix);
    t = t.replace(/([\u0E00-\u0E7F)])\/(?=[\u0E00-\u0E7F(])/g, (m, ch) => { cnt.slashBreaks++; return ch + '/\u200B'; });
    return o + t + c;
  }));
});
// 5) justify every content paragraph
const setJc = p => {
  const ppr = (p.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [''])[0];
  if (/<w:jc w:val="[^"]+"\/>/.test(ppr)) return p.replace(/(<w:pPr>[\s\S]*?)<w:jc w:val="[^"]+"\/>/, '$1<w:jc w:val="thaiDistribute"/>');
  if (ppr) return p.replace(/<w:pPr>([\s\S]*?)<\/w:pPr>/, (all, inner) => { const k = inner.search(/<w:(?:textDirection|textAlignment|textboxTightWrap|outlineLvl|divId|cnfStyle|rPr|sectPr|pPrChange)\b/); const tag = '<w:jc w:val="thaiDistribute"/>'; return '<w:pPr>' + (k < 0 ? inner + tag : inner.slice(0, k) + tag + inner.slice(k)) + '</w:pPr>'; });
  return p.replace(/^(<w:p(?: [^>]*)?>)/, '$1<w:pPr><w:jc w:val="thaiDistribute"/></w:pPr>');
};
const para = p => {
  const t = [...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join(''); if (!t.trim()) return p;
  const sid = (p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || 'Normal';
  if (isHead(sid) || /<w:outlineLvl w:val="[0-8]"/.test(p)) { sk('heading'); return p; }
  if (/<w:instrText|<w:fldSimple/.test(p)) { sk('field'); return p; }
  if (p.includes('<w:drawing') || p.includes('<w:pict')) { sk('picture'); return p; }
  if (CODEFONT.test(p)) { sk('code font'); return p; }
  { const plain = t.replace(/\u200B/g, '').replace(/&[a-z]+;/g, 'x').trim(); if (!/[\u0E00-\u0E7F]/.test(plain) && !/\s/.test(plain)) { sk('identifier'); return p; } }
  const jc = (p.match(/<w:pPr>[\s\S]*?<w:jc w:val="([^"]+)"/) || [])[1] || styJc(sid) || 'left';
  if (['center', 'right', 'end'].includes(jc)) { sk('center/right'); return p; }
  if (jc === 'thaiDistribute') return p;
  cnt.justified++; return setJc(p);
};
const isHdrRow = tr => /<w:trPr>[\s\S]*?<w:tblHeader\/>[\s\S]*?<\/w:trPr>/.test(tr) && [...tr.matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].every(tc => /<w:shd [^>]*w:fill="(?!auto|FFFFFF)[0-9A-Fa-f]{6}"/.test(tc[1]));
body = body.replace(/<w:tr\b[\s\S]*?<\/w:tr>|<w:p[ >][\s\S]*?<\/w:p>/g, m => {
  if (m.startsWith('<w:tr')) { if (isHdrRow(m)) { sk('header row'); return m; } return m.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, para); }
  return para(m);
});
// restore code boxes: only let them split across pages (no other change)
body = body.replace(/\u0000CODE(\d+)\u0000/g, (m, i) => { let k = 0; return saved[+i].replace(/<w:cantSplit\/>/g, '').replace(/<w:tblHeader\/>/g, '').replace(/<w:trPr><\/w:trPr>/g, '').replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => (k++ === 0 ? p : p.replace(/<w:keepNext\/>/g, '').replace(/<w:keepLines\/>/g, ''))); });
// stacked empty paragraphs between blocks -> keep one (never touches section breaks)
body = body.replace(/(?:<w:p [^>]*>(?:<w:pPr>(?:(?!<\/w:pPr>)[\s\S])*<\/w:pPr>)?<\/w:p>){2,}/g, m => { if (/<w:sectPr/.test(m)) return m; cnt.emptyCollapsed = (cnt.emptyCollapsed || 0) + 1; return m.match(/^<w:p [^>]*>(?:<w:pPr>(?:(?!<\/w:pPr>)[\s\S])*<\/w:pPr>)?<\/w:p>/)[0]; });
d.set('word/document.xml', x.slice(0, ce) + body); saveDocx(d, f);
console.log(JSON.stringify(cnt));
