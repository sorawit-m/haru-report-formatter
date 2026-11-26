// ทุกช่องในตารางกว้าง: ป้ายกำกับค่า (Password:, User:, Email: ...) ที่ไม่ได้อยู่ต้นย่อหน้า -> ขึ้นย่อหน้าใหม่
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0;
const LAB = /(?:Password|Username|User|Email|E-mail|OTP|PIN|รหัสผ่าน|ชื่อผู้ใช้|อีเมล)\u200B?\s?:/;
x = x.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g < 7) return t;
  return t.replace(/<w:p( [^>]*)?>((?:<w:pPr>[\s\S]*?<\/w:pPr>)?)([\s\S]*?)<\/w:p>/g, (p, a, ppr, body) => {
    const runs = body.match(/<w:r(?: [^>]*)?>[\s\S]*?<\/w:r>/g) || []; if (runs.length === 0) return p;
    const plain = runs.map(r => [...r.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(m => m[1]).join(''));
    const full = plain.join(''); if (!LAB.test(full)) return p;
    // split each run text at label positions (not at paragraph start)
    let seen = 0; const out = [[]];
    runs.forEach((r, i) => {
      const open = r.match(/^<w:r(?: [^>]*)?>/)[0], rpr = (r.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [''])[0];
      if (/<w:drawing|<w:fldChar|<w:instrText|<w:tab\/>/.test(r)) { out[out.length - 1].push(r); return; }
      const re = new RegExp(LAB.source, 'g'); let last = 0, m, txt = plain[i];
      while ((m = re.exec(txt))) { if (seen + m.index > 0 && txt.slice(0, m.index).trim() + (seen ? 'x' : '') !== '') { const before = txt.slice(last, m.index).replace(/\s+$/, ''); if (before) out[out.length - 1].push(`${open}${rpr}<w:t xml:space="preserve">${before}</w:t></w:r>`); out.push([]); last = m.index; n++; } }
      const rest = txt.slice(last); if (rest) out[out.length - 1].push(`${open}${rpr}<w:t xml:space="preserve">${rest}</w:t></w:r>`);
      seen += txt.length;
    });
    if (out.length === 1) return p;
    return out.filter(o => o.length).map(o => `<w:p${a || ''}>${ppr}${o.join('')}</w:p>`).join('');
  });
});
d.set('word/document.xml', x); saveDocx(d, f); console.log('label lines split:', n);
