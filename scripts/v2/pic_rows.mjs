// แถวที่มีรูป (ป้าย "รูปประกอบ..." + รูป) ห้ามแยกข้ามหน้า — ป้ายกับรูปอยู่หน้าเดียวกัน
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0;
x = x.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, tr => { if (!/<w:drawing/.test(tr) || /<w:cantSplit\/>/.test(tr)) return tr; n++; return /<w:trPr>/.test(tr) ? tr.replace(/<w:trPr>/, '<w:trPr><w:cantSplit/>') : tr.replace(/^(<w:tr\b[^>]*>)/, '$1<w:trPr><w:cantSplit/></w:trPr>'); });
d.set('word/document.xml', x); saveDocx(d, f); console.log('picture rows kept on one page:', n);
