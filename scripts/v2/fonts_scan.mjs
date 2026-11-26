import { loadDocx } from '../lib/docx.mjs';
const d = loadDocx(process.argv[2]);

const names = d.entries.map(e => e.name); console.log(names.filter(n => /footer|header|styles|theme/.test(n)).join(', '));
for (const n of names.filter(n => /footer\d*\.xml$/.test(n))) { const x = d.get(n); console.log(n, (x.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || []).map(p => `[jc=${(p.match(/<w:jc w:val="([^"]+)"/) || [])[1]} style=${(p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1]} PAGE=${/PAGE/.test(p)} tab=${/<w:tab\/>/.test(p)}]`).join(' ')); }
const all = names.filter(n => /^word\/(document|styles|header\d*|footer\d*|numbering)\.xml$/.test(n)).map(n => d.get(n)).join('');
const fonts = {}; for (const m of all.matchAll(/<w:rFonts ([^>]*)\/>/g)) for (const a of m[1].matchAll(/w:(\w+)="([^"]+)"/g)) { const k = a[1] + '=' + a[2]; fonts[k] = (fonts[k] || 0) + 1; }
console.log(fonts);
const th = d.get('word/theme/theme1.xml') || ''; console.log('theme latin:', (th.match(/<a:majorFont><a:latin typeface="([^"]*)"/) || [])[1], '/', (th.match(/<a:minorFont><a:latin typeface="([^"]*)"/) || [])[1]);
