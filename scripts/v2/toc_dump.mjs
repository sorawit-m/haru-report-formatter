import { loadDocx } from '../lib/docx.mjs';
for (const f of process.argv.slice(2)) {
  const d = loadDocx(f); const x = d.get('word/document.xml'); const st = d.get('word/styles.xml');
  const names = Object.fromEntries([...st.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[\s\S]*?<w:name w:val="([^"]+)"/g)].map(m => [m[1], m[2]]));
  console.log('### ' + f.split('\\').pop());
  for (const i of x.matchAll(/<w:instrText[^>]*>([^<]*TOC[^<]*)<\/w:instrText>/g)) console.log('  INSTR', i[1]);
  const toc = [...x.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map(a => a[0]).filter(p => { const s = (p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1]; return s && /^toc \d$/i.test(names[s] || ''); });
  const cnt = {}; for (const p of toc) { const t = [...p.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').replace(/\d+$/, '').trim(); cnt[t] = (cnt[t] || 0) + 1; }
  console.log('  toc lines', toc.length); for (const [t, n] of Object.entries(cnt).slice(0, 60)) console.log(`   ${n > 1 ? n + 'x ' : ''}${t.slice(0, 80)}`);
}
