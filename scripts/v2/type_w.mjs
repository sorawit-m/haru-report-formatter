import { loadDocx } from '../lib/docx.mjs';
import { chunkWidth } from '../lib/thaiwrap.mjs';
const x = loadDocx(process.argv[2]).get('word/document.xml'); let best = 0, ex = '';
for (const m of x.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)) { const rows = [...m[0].matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(r => [...r[0].matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(c => [...c[1].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('')));
  if (!/^Field/.test(rows[0]?.[0] || '')) continue; for (const r of rows.slice(1)) for (const w of (r[1] || '').split(/\s+/)) { const v = chunkWidth(w); if (v > best) { best = v; ex = r[1]; } } }
console.log('longest Type word:', ex, (best * 16 / 28.3465 + 0.4).toFixed(2), 'cm');
