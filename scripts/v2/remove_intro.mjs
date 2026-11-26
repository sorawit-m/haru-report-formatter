// แทน remove_intro.ps1 (ไม่ต้องเปิด Word): ลบบท "บทนำ" (หัวข้อระดับ 1 หลังสารบัญ) ถึงก่อนบทระดับ 1 ถัดไป
// คงการเริ่มเลขหน้าใหม่ของ section: ถ้า section ของบทนำเริ่มเลขใหม่ (pgNumType start) ย้ายค่าไปให้ section ที่ตามมา
// usage: node remove_intro.mjs <file.docx>
import { loadDocx, saveDocx } from '../lib/docx.mjs';
import { textOfXml } from '../lib/wordxml.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const st = d.get('word/styles.xml');
// สไตล์ระดับ 1: outlineLvl 0 หรือชื่อ heading 1 / หัวเรื่อง 1 (รวมสไตล์ที่ basedOn)
const styles = {}; for (const m of st.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g)) {
  styles[m[1]] = { name: ((m[2].match(/<w:name w:val="([^"]+)"/) || [])[1] || '').toLowerCase(), lvl: (m[2].match(/<w:outlineLvl w:val="(\d)"/) || [])[1], based: (m[2].match(/<w:basedOn w:val="([^"]+)"/) || [])[1] };
}
const lvlOf = id => { for (let i = 0, s = styles[id]; s && i < 10; i++, s = styles[s.based]) { if (s.lvl !== undefined) return +s.lvl; if (/^(heading 1|หัวเรื่อง 1)$/.test(s.name)) return 0; } return 9; };
// ลูกระดับบนของ body
const bs = x.indexOf('<w:body>') + 8, be = x.lastIndexOf('</w:body>');
const kids = []; { let i = bs; const re = /<w:(p|tbl|sdt|sectPr|bookmarkStart|bookmarkEnd|customXml|moveFrom|moveTo|ins|del)\b/g;
  while (i < be) { re.lastIndex = i; const m = re.exec(x); if (!m || m.index >= be) break; const name = m[1]; const s = m.index;
    const gt = x.indexOf('>', s); if (x[gt - 1] === '/') { kids.push({ s, e: gt + 1, name }); i = gt + 1; continue; }
    let depth = 0; const r2 = new RegExp(`<w:${name}(?=[\\s>/])[^>]*?(/?)>|</w:${name}>`, 'g'); r2.lastIndex = s; let k, e = be;
    while ((k = r2.exec(x))) { if (k[0][1] === '/') { if (--depth === 0) { e = k.index + k[0].length; break; } } else if (!k[1]) depth++; }
    kids.push({ s, e, name }); i = e; } }
const xmlOf = k => x.slice(k.s, k.e);
const isLvl1 = k => { if (k.name !== 'p') return false; const p = xmlOf(k); if (!textOfXml(p).trim()) return false;
  const ol = p.match(/<w:pPr>[\s\S]*?<w:outlineLvl w:val="(\d)"/); if (ol) return +ol[1] === 0;
  const sid = (p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1]; return !!sid && lvlOf(sid) === 0; };
// เริ่มหาหลังสารบัญ (sdt สารบัญ หรือย่อหน้าที่มีฟิลด์ TOC/PAGEREF ตัวสุดท้าย)
let from = 0; kids.forEach((k, i) => { const t = xmlOf(k); if ((k.name === 'sdt' && /Table of Contents|docPartGallery w:val="Table of Contents"/.test(t)) || /<w:instrText[^>]*>\s*TOC\b|PAGEREF/.test(t)) from = i + 1; });
const ii = kids.findIndex((k, i) => i >= from && isLvl1(k) && textOfXml(xmlOf(k)).includes('บทนำ'));
if (ii < 0) { console.log('no intro chapter'); process.exit(0); }
let ni = kids.findIndex((k, i) => i > ii && isLvl1(k)); const lastIdx = (ni < 0 ? kids.length : ni) - 1;
// sectPr ที่อยู่ในช่วงที่ลบ (ย่อหน้าที่จบ section) — ถ้ามี pgNumType start ส่งต่อให้ sectPr ถัดไป
const gone = kids.slice(ii, lastIdx + 1).filter(k => k.name !== 'sectPr');
const startNo = (() => { for (const k of gone) { const m = xmlOf(k).match(/<w:sectPr\b[\s\S]*?<w:pgNumType [^>]*w:start="(\d+)"/); if (m) return m[1]; } return null; })();
const s = gone[0].s, e = gone[gone.length - 1].e;
let y = x.slice(0, s) + x.slice(e);
if (startNo) {
  const at = y.indexOf('<w:sectPr', s); if (at >= 0) { const end = y.indexOf('</w:sectPr>', at); let sp = y.slice(at, end);
    if (!/<w:pgNumType [^>]*w:start=/.test(sp)) {
      sp = /<w:pgNumType\b/.test(sp) ? sp.replace(/<w:pgNumType\b/, `<w:pgNumType w:start="${startNo}"`)
        : (() => { const a = sp.search(/<w:(?:cols|formProt|vAlign|noEndnote|titlePg|textDirection|bidi|rtlGutter|docGrid|printerSettings)\b/); const pos = a >= 0 ? a : sp.length; return sp.slice(0, pos) + `<w:pgNumType w:start="${startNo}"/>` + sp.slice(pos); })();
      y = y.slice(0, at) + sp + y.slice(end);
    } }
}
d.set('word/document.xml', y); saveDocx(d, f);
console.log(`intro chapter removed (${textOfXml(x.slice(s, e)).length} characters)`);
