// หัวข้อ + หัวตาราง + แถวหัวกลุ่ม (พื้นสีทั้งแถว) ต้องอยู่หน้าเดียวกับแถวข้อมูลแรกที่ตามมา
// - ย่อหน้านอกตารางที่ตามด้วยตาราง และสั้น (หัวข้อ/คำบรรยาย) -> keepNext
// - ในตาราง (>1 คอลัมน์): แถวแรก และทุกแถวที่ทุกช่องมีพื้นสี -> keepNext ทุกย่อหน้า + ห้ามแยกแถว
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const cnt = { rows: 0, captions: 0 };
const shaded = tc => /<w:shd [^>]*w:fill="(?!auto|FFFFFF|ffffff)[0-9A-Fa-f]{6}"/.test(tc);
const addKeep = p => /<w:keepNext\/>/.test(p) ? p : (/<w:pPr>/.test(p) ? p.replace(/<w:pPr>(?:(<w:pStyle [^>]*\/>))?/, (m, s) => `<w:pPr>${s || ''}<w:keepNext/>`) : p.replace(/^(<w:p(?: [^>]*)?>)/, '$1<w:pPr><w:keepNext/></w:pPr>'));
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  if (/<w:tbl>[\s\S]*<w:tbl>/.test(t)) return t;
  const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g < 2) return t;
  let i = 0;
  return t.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, tr => {
    const tcs = [...tr.matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(m => m[1]);
    const keep = !/<w:drawing/.test(tr) && (i++ === 0 || (tcs.length && tcs.every(shaded)));
    if (!keep) return tr; cnt.rows++;
    tr = /<w:cantSplit\/>/.test(tr) ? tr : (/<w:trPr>/.test(tr) ? tr.replace(/<w:trPr>/, '<w:trPr><w:cantSplit/>') : tr.replace(/^(<w:tr\b[^>]*>)/, '$1<w:trPr><w:cantSplit/></w:trPr>'));
    return tr.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, addKeep);
  });
});
// short paragraph right before a table -> keep with the table
x = x.replace(/(<w:p[ >](?:(?!<w:p[ >])[\s\S])*?<\/w:p>)(?=<w:tbl>)/g, p => {
  const t = [...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('');
  if (!t.trim() || t.length > 120 || /<w:sectPr/.test(p)) return p; cnt.captions++; return addKeep(p);
});
d.set('word/document.xml', x); saveDocx(d, f); console.log(JSON.stringify(cnt));
