import { loadDocx } from '../lib/docx.mjs';
const x = loadDocx(process.argv[2]).get('word/document.xml'); const k = {};
for (const m of x.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)) { const g = ((m[0].match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g === 1) continue;
  for (const p of m[0].matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)) { const ind = ((p[0].match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [''])[0].match(/<w:ind [^>]*\/>/) || ['(none)'])[0]; k[ind] = (k[ind] || 0) + 1; } }
console.log(Object.entries(k).sort((a, b) => b[1] - a[1]).slice(0, 8));
