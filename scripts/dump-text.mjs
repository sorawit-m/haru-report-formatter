// dump-text.mjs — ดึงข้อความทั้งเล่มออกมาเป็นไฟล์ .txt สำหรับตรวจคำผิด (ทุก OS)
// ใช้: node dump-text.mjs <in.docx> <out.txt>
// รูปแบบ: 1 บรรทัด = 1 ย่อหน้า, ข้อความในตารางขึ้นต้นด้วย "[T] ", ย่อหน้าหัวข้อขึ้นต้นด้วย "## "
import fs from 'node:fs';
import { loadDocx, parseXml, all, el, attr, textOf } from './lib/docx.mjs';
const [, , inF, outF] = process.argv;
if (!inF || !outF) { console.error('usage: node dump-text.mjs in.docx out.txt'); process.exit(1); }
const doc = loadDocx(inF);
const styles = doc.get('word/styles.xml') || '';
const heads = new Set([...styles.matchAll(/<w:style\b[^>]*w:styleId="([^"]*)"[^>]*>(?:(?!<\/w:style>)[\s\S])*?<w:outlineLvl w:val="\d"/g)].map(m => m[1]));
const root = parseXml(doc.get('word/document.xml'));
const out = [];
for (const p of all(root, 'w:p')) {
  let inTable = false; for (let n = p.parent; n; n = n.parent) if (n.name === 'w:tc') { inTable = true; break; }
  const pPr = el(p, 'w:pPr'); const ps = pPr && el(pPr, 'w:pStyle');
  const t = all(p, 'w:r', [], 'w:p').map(r => r.children.map(c => c.name === 'w:t' ? c.children.map(x => x.text).join('') : c.name === 'w:tab' ? '\t' : (c.name === 'w:br' ? ' ⏎ ' : '')).join('')).join('');
  if (!t.trim()) continue;
  out.push((inTable ? '[T] ' : '') + (ps && heads.has(attr(ps, 'w:val')) ? '## ' : '') + t);
}
fs.writeFileSync(outF, '﻿' + out.join('\n'), 'utf8');
console.log(`paragraphs: ${out.length} -> ${outF}`);
