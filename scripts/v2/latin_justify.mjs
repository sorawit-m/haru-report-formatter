// ย่อหน้าที่จัดแบบ thaiDistribute (ตั้งเองหรือได้จากสไตล์) และเป็นภาษาอังกฤษเกินครึ่ง -> both (ขยายช่องว่างระหว่างคำ ไม่ถ่างตัวอักษร)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const st = d.get('word/styles.xml'); let n = 0;
const sty = {}; for (const m of st.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[\s\S]*?<\/w:style>/g)) sty[m[1]] = { jc: (m[0].match(/<w:jc w:val="([^"]+)"/) || [])[1], based: (m[0].match(/<w:basedOn w:val="([^"]+)"/) || [])[1], def: /w:default="1"/.test(m[0]) && /w:type="paragraph"/.test(m[0]) };
const defId = Object.keys(sty).find(k => sty[k].def) || 'Normal';
const styJc = id => { let s = id, k = 0; while (s && sty[s] && k++ < 8) { if (sty[s].jc) return sty[s].jc; s = sty[s].based; } return 'left'; };
const b = x.indexOf('<w:body>'); const ce = x.indexOf('</w:p>', x.indexOf('<w:sectPr', b)) + 6;
const body = x.slice(ce).replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
  const ppr = (p.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [''])[0];
  const own = (ppr.match(/<w:jc w:val="([^"]+)"\/>/) || [])[1];
  const eff = own || styJc((p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || defId);
  if (eff !== 'thaiDistribute') return p;
  const t = [...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').replace(/\u200B/g, '').replace(/\s/g, '');
  if (!t || (t.match(/[\u0E00-\u0E7F]/g) || []).length / t.length >= 0.5) return p;
  n++;
  if (own) return p.replace(/<w:jc w:val="thaiDistribute"\/>/, '<w:jc w:val="both"/>');
  if (ppr) return p.replace(/<w:pPr>([\s\S]*?)<\/w:pPr>/, (a, i) => { const k = i.search(/<w:(?:textDirection|textAlignment|textboxTightWrap|outlineLvl|divId|cnfStyle|rPr|sectPr|pPrChange)\b/); const tag = '<w:jc w:val="both"/>'; return '<w:pPr>' + (k < 0 ? i + tag : i.slice(0, k) + tag + i.slice(k)) + '</w:pPr>'; });
  return p.replace(/^(<w:p(?: [^>]*)?>)/, '$1<w:pPr><w:jc w:val="both"/></w:pPr>');
});
d.set('word/document.xml', x.slice(0, ce) + body); saveDocx(d, f); console.log('english-heavy paragraphs -> both:', n);
