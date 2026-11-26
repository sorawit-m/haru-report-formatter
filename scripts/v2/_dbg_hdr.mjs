import { loadDocx } from '../lib/docx.mjs';
const x = loadDocx(process.argv[2]).get('word/document.xml');
const show = (needle) => { const i = x.search(needle); const p = x.slice(x.lastIndexOf('<w:p ', i), x.indexOf('</w:p>', i) + 6); console.log('---', needle, '\n', p.slice(0, 900)); const tc = x.slice(x.lastIndexOf('<w:tc>', i), i); console.log(' tcPr:', (tc.match(/<w:tcPr>[\s\S]*?<\/w:tcPr>/) || [''])[0].slice(0, 400)); };
show(/>Test\u200B? ID</); show(/>Test Case ID</); show(/>Functional</);
const t = x.lastIndexOf('<w:tbl>', x.search(/>Test Case ID</)); console.log('summary tblPr:', x.slice(t, t + 600));
