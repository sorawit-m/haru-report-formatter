// กล่องโค้ด: 3 ย่อหน้าแรก (หัวกล่อง + บรรทัดแรก) ติดกัน ไม่ให้หัว "ตัวอย่าง ..." ค้างท้ายหน้าคนเดียว
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0;
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => { const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g !== 1) return t; let k = 0; n++;
  return t.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => { if (k++ >= 2 || /<w:keepNext\/>/.test(p)) return p; return /<w:pPr>/.test(p) ? p.replace(/<w:pPr>/, '<w:pPr><w:keepNext/>') : p.replace(/^(<w:p(?: [^>]*)?>)/, '$1<w:pPr><w:keepNext/></w:pPr>'); }); });
d.set('word/document.xml', x); saveDocx(d, f); console.log('code boxes: heading kept with first lines:', n);
