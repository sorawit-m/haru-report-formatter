import { loadDocx } from '../lib/docx.mjs';
const x = loadDocx(process.argv[2]).get('word/document.xml');
const i = x.search(/DASH_\u200B?001/); const t0 = x.lastIndexOf('<w:tbl>', i); const t = x.slice(t0, x.indexOf('</w:tbl>', i));
for (const r of t.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)) { const tcs = [...r[0].matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(c => c[1]);
  console.log('row cells', tcs.length, (r[0].match(/<w:trPr>[\s\S]*?<\/w:trPr>/) || [''])[0], 'img', /<w:drawing/.test(r[0]));
  if (tcs.length === 1) for (const p of tcs[0].matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)) console.log('   para keepNext', /<w:keepNext\/>/.test(p[0]), 'img', /<w:drawing/.test(p[0]), [...p[0].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').slice(0, 30)); }
console.log('tables after:', x.slice(x.indexOf('</w:tbl>', i), x.indexOf('</w:tbl>', i) + 300).replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, ''));
