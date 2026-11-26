import { loadDocx } from '../lib/docx.mjs';
const x = loadDocx(process.argv[2]).get('word/document.xml'); const k = {}; let shown = 0;
for (const m of x.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)) { const p = m[0]; if (!/รูปที่/.test(p) || /TableofFigures|PAGEREF|TOC \\/.test(p)) continue;
  const flat = p.replace(/<w:instrText[^>]*>([^<]*)<\/w:instrText>/g, '{$1}').replace(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g, '«$1»').replace(/<[^>]+>/g, '');
  const key = flat.replace(/\d+/g, 'N').slice(0, 80); k[key] = (k[key] || 0) + 1; if (shown++ < 3) console.log(flat.slice(0, 200)); }
console.log(k);
