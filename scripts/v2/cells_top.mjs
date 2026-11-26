// ตารางกว้าง (>=7 คอลัมน์): ทุกช่องที่ไม่ใช่หัวตาราง -> ชิดบน (vAlign top ชัดเจน) + ไม่มีระยะก่อนย่อหน้า + ลบบรรทัดว่างหัวช่อง
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const cnt = { cells: 0, emptyLead: 0, spaceBefore: 0 };
const shaded = tc => /<w:shd [^>]*w:fill="(?!auto|FFFFFF|ffffff)[0-9A-Fa-f]{6}"/.test(tc);
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g < 7) return t;
  return t.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, tr => {
    const tcs = [...tr.matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(m => m[1]); if (tcs.length > 1 && tcs.every(shaded)) return tr;
    return tr.replace(/<w:tc>([\s\S]*?)<\/w:tc>/g, (all, inner) => {
      if (/<w:drawing/.test(inner)) return all;
      cnt.cells++;
      inner = inner.replace(/<w:vAlign [^>]*\/>/g, '');
      inner = /<w:tcPr>/.test(inner) ? inner.replace(/<\/w:tcPr>/, '<w:vAlign w:val="top"/></w:tcPr>') : '<w:tcPr><w:vAlign w:val="top"/></w:tcPr>' + inner;
      // leading empty paragraphs (keep at least one paragraph)
      let ps = [...inner.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map(m => m[0]);
      while (ps.length > 1 && !/<w:t[ >]/.test(ps[0])) { inner = inner.replace(ps[0], ''); ps.shift(); cnt.emptyLead++; }
      inner = inner.replace(/<w:spacing ([^>]*)w:before="[1-9]\d*"/g, (m, a) => { cnt.spaceBefore++; return `<w:spacing ${a}w:before="0"`; });
      return '<w:tc>' + inner + '</w:tc>';
    });
  });
});
d.set('word/document.xml', x); saveDocx(d, f); console.log(JSON.stringify(cnt));
