import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0;
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => { const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g < 7 || !/<w:tblHeader\/>/.test(t)) return t; n++; return t.replace(/<w:tblHeader\/>/g, '').replace(/<w:trPr><\/w:trPr>/g, ''); });
d.set('word/document.xml', x); saveDocx(d, f); console.log('wide tables: header row no longer repeats on the next page:', n);
