// การขึ้นบรรทัดด้วย Shift+Enter (<w:br/>) ในช่องตาราง: Word จะยืดบรรทัดก่อนหน้าให้เต็มเมื่อชิดซ้าย-ขวา (User:      admin_uat1)
// -> แยกเป็นย่อหน้าใหม่ (รูปแบบย่อหน้าเดิม) บรรทัดนั้นจึงไม่ถูกยืด
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0;
x = x.replace(/<w:tc>([\s\S]*?)<\/w:tc>/g, (all, inner) => {
  if (!/<w:br\/>/.test(inner)) return all;
  inner = inner.replace(/<w:p( [^>]*)?>((?:<w:pPr>[\s\S]*?<\/w:pPr>)?)([\s\S]*?)<\/w:p>/g, (p, attrs, ppr, body) => {
    if (!/<w:br\/>/.test(body) || /<w:drawing/.test(body)) return p;
    const out = []; let cur = '';
    for (const r of body.match(/<w:r(?: [^>]*)?>[\s\S]*?<\/w:r>|<w:(?:hyperlink|bookmarkStart|bookmarkEnd|proofErr)[\s\S]*?(?:<\/w:hyperlink>|\/>)/g) || []) {
      if (!/^<w:r[ >]/.test(r) || !/<w:br\/>/.test(r)) { cur += r; continue; }
      const open = r.match(/^<w:r(?: [^>]*)?>/)[0]; const rpr = (r.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [''])[0];
      const pieces = r.slice(open.length + rpr.length, -6).split('<w:br/>');
      pieces.forEach((pc, i) => { if (i) { out.push(cur); cur = ''; n++; } if (pc) cur += open + rpr + pc + '</w:r>'; });
    }
    out.push(cur);
    return out.map(b => `<w:p${attrs || ''}>${ppr}${b}</w:p>`).join('');
  });
  return '<w:tc>' + inner + '</w:tc>';
});
d.set('word/document.xml', x); saveDocx(d, f); console.log('line breaks in cells turned into paragraphs:', n);
