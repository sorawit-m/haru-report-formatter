// แสดงหัวข้อระดับ 1-2 (นอกตาราง) + นับย่อหน้าหัวข้อที่อยู่ในตาราง
import { loadDocx } from '../lib/docx.mjs';
for (const f of process.argv.slice(2)) {
  const d = loadDocx(f); const x = d.get('word/document.xml'); const st = d.get('word/styles.xml');
  const lvl = {}; for (const m of st.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[\s\S]*?<\/w:style>/g)) { const o = m[0].match(/<w:outlineLvl w:val="(\d)"/); const nm = (m[0].match(/<w:name w:val="([^"]+)"/) || [])[1] || ""; const h = nm.match(/^(?:heading|หัวเรื่อง|หัวข้อ)\s*(\d)$/i); if (o) lvl[m[1]] = +o[1]; else if (h) lvl[m[1]] = +h[1] - 1; else { const b = (m[0].match(/<w:basedOn w:val="([^"]+)"/) || [])[1]; if (b) lvl["?" + m[1]] = b; } }
  for (const k of Object.keys(lvl)) if (k[0] === "?") { const id = k.slice(1); let b = lvl[k], n = 0; while (b && lvl[b] == null && n++ < 9) b = lvl["?" + b]; if (lvl[id] == null && typeof lvl[b] === "number") lvl[id] = lvl[b]; }
  console.log('### ' + f.split('\\').pop());
  let depth = 0, inTbl = {}, i = 0;
  const re = /<w:tbl>|<\/w:tbl>|<w:p[ >][\s\S]*?<\/w:p>/g; let m;
  while ((m = re.exec(x))) {
    if (m[0] === '<w:tbl>') { depth++; continue; } if (m[0] === '</w:tbl>') { depth--; continue; }
    const p = m[0]; const sid = (p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1]; const ol = (p.match(/<w:outlineLvl w:val="(\d)"/) || [])[1];
    const L = ol != null ? +ol : lvl[sid]; if (L == null || L > 8) continue;
    const t = [...p.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').trim(); if (!t) continue;
    if (depth) { const k = `L${L + 1}:${t.slice(0, 20)}`; inTbl[k] = (inTbl[k] || 0) + 1; continue; }
    if (L <= 1) console.log(`  ${'  '.repeat(L)}[${sid}] ${t.slice(0, 90)}`);
  }
  if (Object.keys(inTbl).length) console.log('  ** headings inside tables:', JSON.stringify(inTbl).slice(0, 400));
}
