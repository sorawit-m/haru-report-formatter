// เล่มรายงานผลทดสอบ (ตารางกรณีทดสอบ + ช่องผลที่พบ/รูป)
// 1) แถวข้อมูล: ข้อความชิดบนทุกช่อง (ไม่จัดกึ่งกลางแนวตั้ง); ข้อความไทยยาวที่จัดกึ่งกลางแนวนอน -> ชิดซ้าย-ขวา
// 2) "ผลที่พบ (Actual Result)" อยู่หน้าเดียวกับรูป: แถวรูปห้ามแยก, ข้อความก่อนรูป keepNext; แถวทดสอบไม่ต้องลากแถวรูป
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const cnt = { vTop: 0, justify: 0, picRows: 0 };
const shaded = tc => /<w:shd [^>]*w:fill="(?!auto|FFFFFF|ffffff)[0-9A-Fa-f]{6}"/.test(tc);
const addKeep = p => /<w:keepNext\/>/.test(p) ? p : (/<w:pPr>/.test(p) ? p.replace(/<w:pPr>(?:(<w:pStyle [^>]*\/>))?/, (m, s) => `<w:pPr>${s || ''}<w:keepNext/>`) : p.replace(/^(<w:p(?: [^>]*)?>)/, '$1<w:pPr><w:keepNext/></w:pPr>'));
const cant = tr => /<w:cantSplit\/>/.test(tr) ? tr : (/<w:trPr>/.test(tr) ? tr.replace(/<w:trPr>/, '<w:trPr><w:cantSplit/>') : tr.replace(/^(<w:tr\b[^>]*>)/, '$1<w:trPr><w:cantSplit/></w:trPr>'));
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g < 7) return t;
  return t.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, tr => {
    const tcs = [...tr.matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(m => m[1]);
    if (tcs.length > 1 && tcs.every(shaded)) return tr;                       // header row untouched
    if (/<w:drawing/.test(tr)) {                                              // result text + picture: one page
      cnt.picRows++;
      tr = cant(tr);
      return tr.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => /<w:drawing/.test(p) ? p.replace(/<w:keepNext\/>/g, '') : addKeep(p));
    }
    // test row: top-aligned cells, long Thai centered text justified, no keepNext into the picture row
    tr = tr.replace(/<w:vAlign w:val="(?:center|bottom)"\/>/g, () => { cnt.vTop++; return '<w:vAlign w:val="top"/>'; });
    tr = tr.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
      p = p.replace(/<w:keepNext\/>/g, '');
      if (!/<w:jc w:val="center"\/>/.test(p)) return p;
      const s = [...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').replace(/​/g, '');
      if (s.trim().length <= 6 || !/[฀-๿]/.test(s)) return p;
      cnt.justify++; return p.replace(/<w:jc w:val="center"\/>/, '<w:jc w:val="thaiDistribute"/>');
    });
    return tr;
  });
});
d.set('word/document.xml', x); saveDocx(d, f); console.log(JSON.stringify(cnt));
