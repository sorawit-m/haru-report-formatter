import { loadDocx } from '../lib/docx.mjs';
const x = loadDocx(process.argv[2]).get('word/document.xml');
const i = x.indexOf('>Test ID<'); const p = x.slice(x.lastIndexOf('<w:p ', i), x.indexOf('</w:p>', i) + 6); console.log('TestID hdr:', p.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g,'[rPr]').slice(0, 400));
const j = x.search(/รูปประกอบการทดสอบ/); console.log('img caption:', x.slice(x.lastIndexOf('<w:p ', j), x.indexOf('</w:p>', j) + 6).replace(/<w:rPr>[\s\S]*?<\/w:rPr>/g,'').slice(0, 900));
