// ตารางที่มีคอลัมน์ล่องหนเกินมา (grid N+1 แต่แถวหัวมี N ช่อง): ตัดคอลัมน์สุดท้ายทิ้ง ช่องที่คลุมเต็มแถว (span N+1) -> span N
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0;
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  if (/<w:tbl>[\s\S]*<w:tbl>/.test(t)) return t;
  const g = [...((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].matchAll(/<w:gridCol w:w="(\d+)"\/>/g))].map(a => +a[1]);
  const rows = [...t.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(r => [...r[0].matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(c => +((c[1].match(/<w:gridSpan w:val="(\d+)"/) || [])[1] || 1)));
  if (!rows.length) return t;
  const N = rows[0].reduce((a, b) => a + b, 0);
  if (g.length !== N + 1 || g[N] > 300) return t;              // only a thin extra column
  if (!rows.every(r => { const s = r.reduce((a, b) => a + b, 0); return s === N || s === N + 1; })) return t;
  n++;
  const ng = g.slice(0, N); ng[N - 1] += g[N];
  t = t.replace(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/, `<w:tblGrid>${ng.map(w => `<w:gridCol w:w="${w}"/>`).join('')}</w:tblGrid>`);
  t = t.replace(/<w:gridSpan w:val="(\d+)"\/>/g, (m, v) => +v === N + 1 ? `<w:gridSpan w:val="${N}"/>` : m);
  t = t.replace(/<w:gridAfter w:val="1"\/>/g, '').replace(/<w:wAfter [^>]*\/>/g, '');
  return t;
});
d.set('word/document.xml', x); saveDocx(d, f); console.log('tables with extra hidden column fixed:', n);
