// ช่องในตารางที่ใช้สไตล์หัวข้อ (heading) -> ใช้รูปแบบเดียวกับช่องข้างๆ ในแถวเดียวกัน (จะได้ไม่เข้าไปในสารบัญ และหน้าตาเหมือนเพื่อนในแถว)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const st = d.get('word/styles.xml'); let n = 0, fallback = 0;
const head = new Set(); for (const m of st.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[\s\S]*?<\/w:style>/g)) { const nm = (m[0].match(/<w:name w:val="([^"]+)"/) || [])[1] || ''; if (/<w:outlineLvl w:val="[0-8]"/.test(m[0]) || /^(?:heading|หัวเรื่อง|หัวข้อ)\s*\d$/i.test(nm)) head.add(m[1]); }
const isHeadP = p => { const s = (p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1]; return (s && head.has(s)) || /<w:outlineLvl w:val="[0-8]"/.test(p); };
x = x.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, tr => {
  const ps = [...tr.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map(m => m[0]);
  if (!ps.some(isHeadP)) return tr;
  const model = ps.find(p => !isHeadP(p) && /<w:t[ >]/.test(p));
  const mPPr = model ? ((model.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [''])[0]) : '<w:pPr><w:jc w:val="center"/></w:pPr>';
  const mRPr = model ? ((model.match(/<w:r(?: [^>]*)?>\s*(<w:rPr>[\s\S]*?<\/w:rPr>)/) || [])[1] || '') : '';
  if (!model) fallback++;
  return tr.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
    if (!isHeadP(p)) return p; n++;
    p = p.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, '').replace(/^(<w:p(?: [^>]*)?>)/, `$1${mPPr}`);
    if (mRPr) p = p.replace(/(<w:r(?: [^>]*)?>)(<w:rPr>[\s\S]*?<\/w:rPr>)?/g, `$1${mRPr}`);
    return p;
  });
});
d.set('word/document.xml', x); saveDocx(d, f); console.log(`heading-styled table cells fixed: ${n}${fallback ? `, rows without a model cell: ${fallback}` : ''}`);
