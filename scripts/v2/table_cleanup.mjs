// ทุกตาราง: 1) แถวหัวตาราง (ทุกช่องมีพื้นสี หรือทำเครื่องหมายหัวตาราง) -> ตัวหนาทั้งแถว
//           2) ข้อความในตารางไม่มีย่อหน้า (firstLine = 0) — ยกเว้นย่อหน้าที่เยื้องแบบแขวน (รายการข้อ 1. 2. 3.)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const cnt = { headerRows: 0, boldRuns: 0, indentFixed: 0 };
const shaded = tc => /<w:shd [^>]*w:fill="(?!auto|FFFFFF|ffffff)[0-9A-Fa-f]{6}"/.test(tc);
const boldRun = r => {
  if (/<w:b\/>/.test(r) && /<w:bCs\/>/.test(r)) return r;
  cnt.boldRuns++;
  if (/<w:rPr>/.test(r)) return r.replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/, (a, i) => { i = i.replace(/<w:b w:val="(?:0|false)"\/>|<w:bCs w:val="(?:0|false)"\/>/g, ''); const k = i.search(/<w:(?:i|iCs|caps|smallCaps|strike|dstrike|outline|shadow|emboss|imprint|noProof|snapToGrid|vanish|webHidden|color|spacing|w|kern|position|sz|szCs|highlight|u|effect|bdr|shd|fitText|vertAlign|rtl|cs|em|lang|eastAsianLayout|specVanish|oMath)\b/); const tag = (/<w:b\/>/.test(i) ? '' : '<w:b/>') + (/<w:bCs\/>/.test(i) ? '' : '<w:bCs/>'); return '<w:rPr>' + (k < 0 ? i + tag : i.slice(0, k) + tag + i.slice(k)) + '</w:rPr>'; });
  return r.replace(/^(<w:r(?: [^>]*)?>)/, '$1<w:rPr><w:b/><w:bCs/></w:rPr>');
};
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  if (/<w:tbl>[\s\S]*<w:tbl>/.test(t)) return t;
  let first = true;
  t = t.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, tr => {
    const tcs = [...tr.matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(m => m[1]);
    const isHdr = (first && tcs.length > 1 && tcs.every(shaded)) || (/<w:tblHeader\/>/.test(tr) && tcs.length > 1 && tcs.every(shaded));
    first = false;
    if (!isHdr) return tr;
    cnt.headerRows++;
    return tr.replace(/<w:r(?: [^>]*)?>[\s\S]*?<\/w:r>/g, r => /<w:t[ >]/.test(r) ? boldRun(r) : r);
  });
  // no first-line indent inside tables (never in code boxes = 1-column tables)
  if ((((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0]).match(/<w:gridCol\b/g) || []).length === 1) return t;
  t = t.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
    if (/<w:drawing|<w:pict/.test(p)) return p;
    const ppr = (p.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [''])[0];
    const ind = (ppr.match(/<w:ind [^>]*\/>/) || [''])[0];
    if (/w:hanging=/.test(ind)) return p;
    if (/w:firstLine="0"/.test(ind) && !/w:left="[1-9]/.test(ind)) return p;
    cnt.indentFixed++;
    const newInd = '<w:ind w:left="0" w:firstLine="0"/>';
    if (ind) return p.replace(ind, newInd);
    if (ppr) return p.replace(/<w:pPr>([\s\S]*?)<\/w:pPr>/, (a, i) => { const k = i.search(/<w:(?:contextualSpacing|mirrorIndents|suppressOverlap|jc|textDirection|textAlignment|textboxTightWrap|outlineLvl|divId|cnfStyle|rPr|sectPr|pPrChange)\b/); return '<w:pPr>' + (k < 0 ? i + newInd : i.slice(0, k) + newInd + i.slice(k)) + '</w:pPr>'; });
    return p.replace(/^(<w:p(?: [^>]*)?>)/, `$1<w:pPr>${newInd}</w:pPr>`);
  });
  return t;
});
d.set('word/document.xml', x); saveDocx(d, f); console.log(JSON.stringify(cnt));
