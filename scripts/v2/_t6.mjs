import { loadDocx } from '../lib/docx.mjs';
const x = loadDocx(process.argv[2]).get('word/document.xml');
const t = x.match(/<w:tbl>[\s\S]*?<\/w:tbl>/)[0]; const rows = [...t.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(r => r[0]);
rows.slice(0, 4).forEach((r, k) => console.log(k, [...r.matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(c => ((c[1].match(/<w:gridSpan w:val="(\d+)"/) || [])[1] ? '{span' + c[1].match(/<w:gridSpan w:val="(\d+)"/)[1] + '}' : '') + [...c[1].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').slice(0, 20) + (c[1].includes('<w:drawing') ? '[IMG]' : '')).join(' | '), /tblHeader/.test(r) ? 'HDR' : ''));
console.log('grid', [...(t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].matchAll(/w:w="(\d+)"/g)].length);
