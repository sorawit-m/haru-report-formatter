import { loadDocx } from '../lib/docx.mjs';
const x = loadDocx(process.argv[2]).get('word/document.xml');
const i = x.search(/LOGIN_\u200B?002/); const t0 = x.lastIndexOf('<w:tbl>', i); const t = x.slice(t0, x.indexOf('</w:tbl>', i));
for (const r of t.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)) console.log((r[0].match(/<w:trPr>[\s\S]*?<\/w:trPr>/) || ['(no trPr)'])[0], 'keepNext paras:', (r[0].match(/<w:keepNext\/>/g) || []).length, 'img:', /<w:drawing/.test(r[0]));
