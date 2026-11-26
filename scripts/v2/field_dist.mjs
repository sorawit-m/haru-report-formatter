import { loadDocx } from '../lib/docx.mjs';
import { chunkWidth } from '../lib/thaiwrap.mjs';
const x = loadDocx(process.argv[2]).get('word/document.xml'); const whole = [], segs = [];
for (const m of x.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)) {
  const rows = [...m[0].matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(r => [...r[0].matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(c => [...c[1].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('')));
  if (!/^Field/.test(rows[0]?.[0] || '')) continue;
  for (const r of rows.slice(1)) { const t = (r[0] || '').trim(); if (!t) continue; whole.push([chunkWidth(t) * 16 / 28.3465, t]); for (const s of t.split(/(?<=[_.\/])/)) segs.push([chunkWidth(s) * 16 / 28.3465, s]); }
}
whole.sort((a, b) => a[0] - b[0]); segs.sort((a, b) => a[0] - b[0]);
const pct = p => whole[Math.floor(whole.length * p) - 1];
console.log('names', whole.length, 'p90', pct(0.9)[0].toFixed(2), pct(0.9)[1], 'p95', pct(0.95)[0].toFixed(2), 'p98', pct(0.98)[0].toFixed(2), 'max', whole.at(-1)[0].toFixed(2));
console.log('longest piece after break points:', segs.at(-1)[0].toFixed(2), segs.at(-1)[1]);
console.log('names wider than 5.0cm text:', whole.filter(w => w[0] > 4.9).length, whole.filter(w => w[0] > 4.9).slice(-8).map(w => w[1]).join(' , '));
