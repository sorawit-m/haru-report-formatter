// ตารางที่ไม่มีรูป: แถวหัวตาราง (แถวต้นที่ทุกช่องมีพื้นสี) ซ้ำทุกหน้าที่ตารางต่อไป
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const cnt = { tables: 0, rows: 0 };
const shaded = tc => /<w:shd [^>]*w:fill="(?!auto|FFFFFF|ffffff)[0-9A-Fa-f]{6}"/.test(tc);
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  if (/<w:tbl>[\s\S]*<w:tbl>/.test(t) || /<w:drawing/.test(t)) return t;
  const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g < 2) return t;
  let leading = true, did = false;
  t = t.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, tr => {
    if (!leading) return tr;
    const tcs = [...tr.matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(m => m[1]);
    if (!(tcs.length && tcs.every(shaded))) { leading = false; return tr; }
    if (/<w:tblHeader\/>/.test(tr)) return tr;
    did = true; cnt.rows++;
    return /<w:trPr>/.test(tr) ? tr.replace(/<w:trPr>/, '<w:trPr><w:tblHeader/>') : tr.replace(/^(<w:tr\b[^>]*>)/, '$1<w:trPr><w:tblHeader/></w:trPr>');
  });
  if (did) cnt.tables++; return t;
});
d.set('word/document.xml', x); saveDocx(d, f); console.log(JSON.stringify(cnt));
