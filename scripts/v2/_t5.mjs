import { loadDocx } from '../lib/docx.mjs';
const x = loadDocx(process.argv[2]).get('word/document.xml');
const i = x.indexOf('Test by'); const t0 = x.lastIndexOf('<w:tbl>', i); const t = x.slice(t0, x.indexOf('</w:tbl>', i));
const rows = [...t.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(r => [...r[0].matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(c => [...c[1].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').slice(0, 25) + (c[1].includes('<w:drawing') ? '[IMG]' : '') + ((c[1].match(/<w:gridSpan w:val="(\d+)"/) || [])[1] ? '{span' + c[1].match(/<w:gridSpan w:val="(\d+)"/)[1] + '}' : '')));
rows.forEach((r, k) => console.log(k, r.join(' | ')));
const caps = [...x.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map(m => m[0]).filter(p => /รูปที่/.test(p)).slice(0, 2); caps.forEach(p => console.log('CAP:', p.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g, '').slice(0, 500)));
