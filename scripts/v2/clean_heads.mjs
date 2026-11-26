// หัวข้อและสารบัญไม่ต้องมีจุดตัดที่มองไม่เห็น (Word ใส่ฟอนต์ Arial ให้ตัวอักษรล่องหนในสารบัญ) -> ลบออก แล้วฟอนต์ที่เหลือเป็นแบบเดียว
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const st = d.get('word/styles.xml'); let n = 0;
const head = new Set(); for (const m of st.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[\s\S]*?<\/w:style>/g)) { const nm = (m[0].match(/<w:name w:val="([^"]+)"/) || [])[1] || ''; if (/<w:outlineLvl w:val="[0-8]"/.test(m[0]) || /^(?:heading|หัวเรื่อง|หัวข้อ|toc)\s*\d$/i.test(nm)) head.add(m[1]); }
x = x.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => { const s = (p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1]; if (!(s && head.has(s)) && !/<w:outlineLvl/.test(p) && !/PAGEREF/.test(p)) return p; if (!/\u200B/.test(p)) return p; n++; return p.replace(/\u200B/g, '').replace(/<w:r(?: [^>]*)?>(?:<w:rPr>(?:(?!<\/w:rPr>)[\s\S])*<\/w:rPr>)?<w:t(?: [^>]*)?><\/w:t><\/w:r>/g, ''); });
d.set('word/document.xml', x); saveDocx(d, f); console.log('headings/TOC lines cleaned:', n);
