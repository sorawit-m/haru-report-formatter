import { loadDocx } from '../lib/docx.mjs';
import { chunkWidth, longestUnit, wrap } from '../lib/thaiwrap.mjs';
const x = loadDocx(process.argv[2]).get('word/document.xml'); const tot = Array(9).fill(0), longest = Array(9).fill(0), ex = Array(9).fill(''); let rows = 0;
for (const m of x.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)) {
  const trs = [...m[0].matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(r => [...r[0].matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(c => [...c[1].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').replace(/\u200B/g, '')));
  if (!trs.length || trs[0][0] !== 'Test ID') continue;
  for (const r of trs.slice(1)) { if (r.length !== 9) continue; rows++; r.forEach((t, i) => { tot[i] += chunkWidth(t); const u = longestUnit(t); if (u > longest[i]) { longest[i] = u; ex[i] = t.slice(0, 50); } }); }
}
const cm = em => em * 14 / 28.3465;
console.log('rows', rows);
['TestID','CaseID','Feature','Desc','Category','Step','Input','Expected','Result'].forEach((n, i) => console.log(n.padEnd(9), 'avg text', cm(tot[i] / rows).toFixed(1), 'cm  longest unbreakable', cm(longest[i]).toFixed(2), 'cm', ex[i]));
