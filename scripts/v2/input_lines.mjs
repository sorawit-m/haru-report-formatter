// คอลัมน์ Input Data: แยกบรรทัดก่อนป้ายกำกับแต่ละค่า (เช่น "Password:" "เลขบัตร:") เหมือนต้นฉบับ ไม่ให้บรรทัดแรกถูกยืดจนห่าง
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0;
const norm = s => s.replace(/[\s​]+/g, '').toLowerCase();
const shaded = tc => /<w:shd [^>]*w:fill="(?!auto|FFFFFF|ffffff)[0-9A-Fa-f]{6}"/.test(tc);
const LABEL = /(\S)( +)((?:[A-Za-z][A-Za-z ]{0,18}|[฀-๿​]{2,20}) ?:)(?!\/)/g;
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g < 7) return t;
  const rows = [...t.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(m => m[0]);
  const hdr = [...rows[0].matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(c => norm([...c[1].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('')));
  let col = hdr.indexOf('inputdata'); if (col < 0) col = 6;
  return t.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, tr => {
    const tcs = [...tr.matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(m => m[1]); if (tcs.length < 7 || tcs.every(shaded)) return tr;
    let ci = 0;
    return tr.replace(/<w:tc>([\s\S]*?)<\/w:tc>/g, (all, inner) => {
      const span = +((inner.match(/<w:gridSpan w:val="(\d+)"/) || [])[1] || 1); const here = ci; ci += span; if (here !== col) return all;
      return '<w:tc>' + inner.replace(/(<w:r(?: [^>]*)?>)((?:<w:rPr>[\s\S]*?<\/w:rPr>)?)<w:t(?: [^>]*)?>([^<]*)<\/w:t><\/w:r>/g, (m, ro, rpr, txt) => {
        if (!LABEL.test(txt)) { LABEL.lastIndex = 0; return m; } LABEL.lastIndex = 0;
        const parts = txt.replace(LABEL, (a, pre, sp, lab) => { n++; return pre + '\u0001' + lab; }).split('\u0001');
        return parts.map((p, i) => `${ro}${rpr}${i ? '<w:br/>' : ''}<w:t xml:space="preserve">${p}</w:t></w:r>`).join('');
      }) + '</w:tc>';
    });
  });
});
d.set('word/document.xml', x); saveDocx(d, f); console.log('line breaks before labels in Input Data:', n);
