import { loadDocx } from '../lib/docx.mjs';
for (const f of process.argv.slice(2)) { const d = loadDocx(f); console.log('###', f.split('\\').pop());
  for (const e of d.entries.filter(e => /word\/(footer|header)\d*\.xml$/.test(e.name))) { const x = d.get(e.name); console.log(' ', e.name, 'PAGE:', /PAGE/.test(x), 'text:', [...x.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').slice(0, 60)); }
  const doc = d.get('word/document.xml'); for (const s of doc.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)) console.log('  sect refs:', [...s[0].matchAll(/<w:(header|footer)Reference [^>]*w:type="(\w+)"[^>]*r:id="(\w+)"/g)].map(m => m[1][0] + m[2] + ':' + m[3]).join(' '), (s[0].match(/<w:pgNumType[^>]*\/>/) || [''])[0], /<w:titlePg\/>/.test(s[0]) ? 'titlePg' : ''); }
