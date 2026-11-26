// ทุกตาราง: ไม่ล้ำขอบซ้าย (tblInd < 0 -> 0) และไม่กว้างเกินความกว้างข้อความของ section ที่อยู่ (ย่อตามสัดส่วน)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); let x = d.get('word/document.xml'); const cnt = { indFixed: 0, scaled: 0 };
const tw = s => +s.match(/<w:pgSz [^>]*w:w="(\d+)"/)[1] - +s.match(/<w:pgMar [^>]*w:left="(-?\d+)"/)[1] - +s.match(/<w:pgMar [^>]*w:right="(-?\d+)"/)[1];
// split body into sections (each ends at a sectPr)
const out = []; let rest = x; const re = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/g; let m, last = 0;
const parts = []; while ((m = re.exec(x))) { parts.push([x.slice(last, m.index + m[0].length), tw(m[0])]); last = m.index + m[0].length; } parts.push([x.slice(last), parts.length ? parts[parts.length - 1][1] : 9000]);
const fixed = parts.map(([chunk, TW]) => chunk.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/g, t => {
  if (/<w:tbl>[\s\S]*<w:tbl>/.test(t)) return t;
  const ind = +((t.match(/<w:tblInd w:w="(-?\d+)"/) || [])[1] || 0);
  if (ind < 0) { t = t.replace(/<w:tblInd [^>]*\/>/, '<w:tblInd w:w="0" w:type="dxa"/>'); cnt.indFixed++; }
  const grid = [...((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].matchAll(/<w:gridCol w:w="(\d+)"\/>/g))].map(a => +a[1]);
  const total = grid.reduce((a, b) => a + b, 0);
  if (!total || total <= TW + 20) return t;
  const k = TW / total; cnt.scaled++;
  t = t.replace(/<w:gridCol w:w="(\d+)"\/>/g, (a, w) => `<w:gridCol w:w="${Math.floor(+w * k)}"/>`);
  t = t.replace(/<w:tcW w:w="(\d+)" w:type="dxa"\/>/g, (a, w) => `<w:tcW w:w="${Math.floor(+w * k)}" w:type="dxa"/>`);
  t = t.replace(/<w:tblW w:w="\d+" w:type="dxa"\/>/, `<w:tblW w:w="${TW}" w:type="dxa"/>`);
  return t;
}));
d.set('word/document.xml', fixed.join('')); saveDocx(d, f); console.log(JSON.stringify(cnt));
