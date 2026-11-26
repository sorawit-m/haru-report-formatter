import { loadDocx } from '../lib/docx.mjs';
const d = loadDocx(process.argv[2]); const x = d.get('word/document.xml'); const st = d.get('word/styles.xml'); const c = {};
const sty = {}; for (const m of st.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[\s\S]*?<\/w:style>/g)) sty[m[1]] = { jc: (m[0].match(/<w:jc w:val="([^"]+)"/) || [])[1], based: (m[0].match(/<w:basedOn w:val="([^"]+)"/) || [])[1], def: /w:default="1"/.test(m[0]) && /w:type="paragraph"/.test(m[0]) };
const defId = Object.keys(sty).find(k => sty[k].def); const sj = id => { let s = id, k = 0; while (s && sty[s] && k++ < 8) { if (sty[s].jc) return sty[s].jc; s = sty[s].based; } return 'left'; };
for (const m of x.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)) for (const p of m[0].matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)) { const t = [...p[0].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join(''); if (!t.trim()) continue; const j = (p[0].match(/<w:pPr>[\s\S]*?<w:jc w:val="([^"]+)"\/>/) || [])[1] || sj((p[0].match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || defId); c[j] = (c[j] || 0) + 1; }
console.log('table paragraph alignment:', c);
