// ชื่อ "สารบัญ" (สไตล์ TOC Heading) ไม่ให้มีเลขลำดับต่อหน้า
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const st = d.get('word/styles.xml'); let n = 0;
const ids = new Set([...st.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[\s\S]*?<\/w:style>/g)].filter(m => /<w:name w:val="(TOC Heading|table of figures)"/i.test(m[0])).map(m => m[1]));
x = x.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => { const s = (p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1]; if (!s || !ids.has(s) || /<w:numPr>/.test(p)) return p; n++; return p.replace(/(<w:pStyle w:val="[^"]+"\/>)/, '$1<w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr>'); });
d.set('word/document.xml', x); saveDocx(d, f); console.log('TOC titles without numbering:', n);
