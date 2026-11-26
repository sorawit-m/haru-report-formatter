// ข้อความไทยทั้งเล่ม (นอกกล่องโค้ด/ฟิลด์): จุดตัดที่มองไม่เห็นเฉพาะระหว่าง "คำ" (ไม่ตัดกลางคำ) + ติดป้ายภาษาไทย (cs + lang th-TH)
// + เอาระยะตัวอักษรที่ตั้งเอง (w:spacing ใน run) ออก — ตัวอักษรต้องไม่บีบ/ไม่ถ่าง
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const cnt = { runs: 0, breaks: 0, spacing: 0 };
const seg = new Intl.Segmenter('th', { granularity: 'word' }); const TH = /[฀-๿]/;
const b = x.indexOf('<w:body>'); const ce = x.indexOf('</w:p>', x.indexOf('<w:sectPr', b)) + 6;
let body = x.slice(ce); const saved = [];
body = body.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => { const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g !== 1) return t; saved.push(t); return `\u0000C${saved.length - 1}\u0000`; });
body = body.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
  if (/<w:instrText|<w:fldSimple/.test(p)) return p;
  return p.replace(/<w:r(?: [^>]*)?>[\s\S]*?<\/w:r>/g, r => {
    if (!/<w:t[ >]/.test(r)) return r;
    r = r.replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/, (a, i) => { const j = i.replace(/<w:spacing w:val="-?\d+"\/>/g, () => { cnt.spacing++; return ''; }); return '<w:rPr>' + j + '</w:rPr>'; });
    const txt = [...r.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join(''); if (!TH.test(txt)) return r;
    cnt.runs++;
    r = r.replace(/(<w:t(?: [^>]*)?>)([^<]*)(<\/w:t>)/g, (a, o, t, c) => {
      const plain = t.replace(/​/g, ''); const parts = [...seg.segment(plain)].map(s => s.segment); let out = '';
      for (let i = 0; i < parts.length; i++) { out += parts[i]; const n = parts[i + 1]; if (n && TH.test(parts[i].slice(-1)) && TH.test(n[0])) { out += '​'; cnt.breaks++; } }
      // keep the latin break points (after _ / .) that were there before
      out = out.replace(/([_\/.])(?=[A-Za-z0-9])/g, '$1');
      return o + out + c;
    });
    const tags = '<w:cs/><w:lang w:bidi="th-TH"/>';
    if (/<w:rPr>/.test(r)) return r.replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/, (a, i) => { i = i.replace(/<w:cs\/>/g, '').replace(/<w:lang [^>]*\/>/g, ''); const k = i.search(/<w:(?:em|eastAsianLayout|specVanish|oMath)\b/); return '<w:rPr>' + (k < 0 ? i + tags : i.slice(0, k) + tags + i.slice(k)) + '</w:rPr>'; });
    return r.replace(/^(<w:r(?: [^>]*)?>)/, `$1<w:rPr>${tags}</w:rPr>`);
  });
});
body = body.replace(/\u0000C(\d+)\u0000/g, (m, i) => saved[+i]);
d.set('word/document.xml', x.slice(0, ce) + body); saveDocx(d, f); console.log(JSON.stringify(cnt));
