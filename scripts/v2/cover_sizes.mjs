// แสดงข้อความแต่ละย่อหน้าบนปก + ขนาดตัวอักษร (half-points -> pt)
import { loadDocx } from '../lib/docx.mjs';
for (const f of process.argv.slice(2)) {
  const x = loadDocx(f).get('word/document.xml');
  const b = x.indexOf('<w:body>'); const s = x.indexOf('<w:sectPr', b); const e = x.indexOf('</w:p>', s) + 6;
  const cover = x.slice(b, e);
  console.log('### ' + f.split('\\').pop());
  for (const p of cover.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || []) {
    const runs = [...p.matchAll(/<w:r[ >][\s\S]*?<\/w:r>/g)].map(m => m[0]);
    const parts = runs.map(r => { const t = [...r.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(m => m[1]).join(''); const sz = (r.match(/<w:sz w:val="(\d+)"/) || [])[1]; const szc = (r.match(/<w:szCs w:val="(\d+)"/) || [])[1]; return t ? `${t}[${sz ? sz / 2 : '-'}/${szc ? szc / 2 : '-'}]` : ''; }).filter(Boolean);
    const psz = (p.match(/<w:pPr>[\s\S]*?<w:sz w:val="(\d+)"/) || [])[1];
    const tb = p.includes('txbxContent') ? ' (textbox)' : '';
    if (parts.length) console.log(`  ${tb} ${parts.join(' | ').slice(0, 300)}`);
  }
}
