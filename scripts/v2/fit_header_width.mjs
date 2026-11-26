// หัวกระดาษที่เป็นตาราง: ปรับความกว้างตารางให้เท่าความกว้างข้อความของ section ที่ใช้หัวกระดาษนั้น (แนวนอนจะยาวเต็มหน้า)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const d = loadDocx(f); const x = d.get('word/document.xml'); const rels = d.get('word/_rels/document.xml.rels');
const target = id => (rels.match(new RegExp(`<Relationship [^>]*Id="${id}"[^>]*Target="([^"]+)"`)) || [])[1];
const done = {};
for (const s of x.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)) {
  const sec = s[0]; const W = +sec.match(/<w:pgSz [^>]*w:w="(\d+)"/)[1] - +sec.match(/w:left="(-?\d+)"/)[1] - +sec.match(/w:right="(-?\d+)"/)[1];
  for (const m of sec.matchAll(/<w:headerReference [^>]*r:id="([^"]+)"/g)) {
    const part = 'word/' + target(m[1]); if (done[part]) continue; let h = d.get(part); if (!h || !/<w:tbl>/.test(h)) continue;
    h = h.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/, t => {
      const g = [...t.matchAll(/<w:gridCol w:w="(\d+)"\/>/g)].map(a => +a[1]); const tot = g.reduce((a, b) => a + b, 0); if (!tot) return t;
      const first = g[0]; const ng = [first, W - first].concat(g.slice(2).map(() => 0)).slice(0, g.length);
      let i = 0; t = t.replace(/<w:gridCol w:w="\d+"\/>/g, () => `<w:gridCol w:w="${ng[i++]}"/>`);
      t = t.replace(/<w:tblW w:w="\d+" w:type="dxa"\/>/, `<w:tblW w:w="${W}" w:type="dxa"/>`);
      let c = 0; t = t.replace(/<w:tcW w:w="\d+" w:type="dxa"\/>/g, () => `<w:tcW w:w="${ng[c++] ?? 0}" w:type="dxa"/>`);
      return t;
    });
    d.set(part, h); done[part] = W; console.log(`${part}: header width -> ${(W / 567).toFixed(1)} cm`);
  }
}
saveDocx(d, f);
