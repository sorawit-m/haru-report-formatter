import { loadDocx } from '../lib/docx.mjs';
const x = loadDocx(process.argv[2]).get('word/document.xml'); const kinds = {}; let ex = 0;
for (const m of x.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)) { const t = m[0]; const g = [...((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].matchAll(/w:w="(\d+)"/g))].map(a => +a[1]);
  const rows = [...t.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(r => [...r[0].matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(c => +((c[1].match(/<w:gridSpan w:val="(\d+)"/) || [])[1] || 1)));
  if (g.length < 7) continue; const sig = `grid${g.length} rows:` + rows.map(r => r.join('+')).join(' / '); kinds[sig] = (kinds[sig] || 0) + 1;
  if (g.length !== 9 && ex++ < 1) console.log('grid', g.map(v => (v / 567).toFixed(2)).join(','), '\n', sig); }
for (const [k, v] of Object.entries(kinds).sort((a, b) => b[1] - a[1]).slice(0, 8)) console.log(v + 'x', k.slice(0, 160));
