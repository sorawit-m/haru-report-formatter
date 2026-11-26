// ข้อความบนปกเป็นสีดำล้วน (ไม่แตะรูป ขนาด ฟอนต์)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
import { coverEnd } from '../lib/wordxml.mjs';
for (const f of process.argv.slice(2)) {
  const d = loadDocx(f); let x = d.get('word/document.xml');
  const b = x.indexOf('<w:body>'); const e = coverEnd(x); let n = 0;
  const cover = x.slice(b, e).replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
    if (!/<w:t[ >]/.test(p)) return p;
    return p.replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/g, (all, i) => { n++; i = i.replace(/<w:color [^>]*\/>/g, ''); const k = i.search(/<w:(?:spacing|w|kern|position|sz|szCs|highlight|u|effect|bdr|shd|fitText|vertAlign|rtl|cs|em|lang|eastAsianLayout|specVanish|oMath)\b/); const tag = '<w:color w:val="000000"/>'; return '<w:rPr>' + (k < 0 ? i + tag : i.slice(0, k) + tag + i.slice(k)) + '</w:rPr>'; })
            .replace(/<w:r>(?!<w:rPr>)/g, () => { n++; return '<w:r><w:rPr><w:color w:val="000000"/></w:rPr>'; });
  });
  x = x.slice(0, b) + cover + x.slice(e); d.set('word/document.xml', x); saveDocx(d, f);
  console.log(`${f.split('\\').pop().slice(0, 3)}: ${n} runs -> black`);
}
