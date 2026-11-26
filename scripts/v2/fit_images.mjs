// แทน fit_images.ps1 / enlarge_images.ps1 / fit_multi_images.ps1 (ไม่ต้องเปิด Word) — รูป inline ในตาราง
//   fit                : รูปที่กว้างกว่าช่อง ย่อตามสัดส่วนให้พอดีช่อง (เว้นขอบช่อง 2pt)                    [fit_images.ps1]
//   enlarge [share]    : ขยาย/ย่อให้เต็มความกว้างช่อง แต่สูงไม่เกิน share ของพื้นที่ข้อความในหน้า (ค่าเริ่ม 0.62) [enlarge_images.ps1]
//   multi [share]      : ย่อหน้าที่มีหลายรูป ย่อให้วางข้างกันในบรรทัดเดียว สูงไม่เกิน share (ค่าเริ่ม 0.6)          [fit_multi_images.ps1]
// usage: node fit_images.mjs <file.docx> fit|enlarge|multi [share]
import { loadDocx, saveDocx } from '../lib/docx.mjs';
import { sections, mapTables, hasNestedTable, mapCells, inlineDrawings, drawingSize, resizeDrawing, EMU_PER_TWIP } from '../lib/wordxml.mjs';
const [f, mode = 'fit', shareArg] = process.argv.slice(2);
const share = +(shareArg || (mode === 'multi' ? 0.6 : 0.62));
const PT = 20 * EMU_PER_TWIP; // 1 pt ใน EMU
const d = loadDocx(f); let n = 0, grow = 0;
const parts = sections(d.get('word/document.xml'));
for (const part of parts) {
  const maxH = part.sect.textH * EMU_PER_TWIP * share;
  part.chunk = mapTables(part.chunk, t => hasNestedTable(t) ? t : mapCells(t, ({ tc, width, pads }) => {
    if (!width || !/<wp:inline/.test(tc)) return tc;
    const inner = (width - pads) * EMU_PER_TWIP;
    return tc.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
      const ds = inlineDrawings(p); if (!ds.length) return p;
      if (mode === 'multi') {
        if (ds.length < 2) return p;
        const sz = ds.map(drawingSize); if (sz.some(s => !s || !s.cx || !s.cy)) return p;
        const avail = inner - 6 * PT * ds.length; const ratio = sz.reduce((a, s) => a + s.cx / s.cy, 0);
        const h = Math.min(maxH, avail / ratio); n++;
        ds.forEach((dr, i) => { p = p.replace(dr, resizeDrawing(dr, h * sz[i].cx / sz[i].cy, h)); });
        return p;
      }
      for (const dr of ds) {
        const s = drawingSize(dr); if (!s || !s.cx || !s.cy) continue;
        let k;
        if (mode === 'enlarge') { k = Math.min((inner - 4 * PT) / s.cx, maxH / s.cy); if (Math.abs(k - 1) < 0.02) continue; }
        else { const max = inner - 2 * PT; if (s.cx <= max) continue; k = max / s.cx; }
        n++; if (k > 1) grow++;
        p = p.replace(dr, resizeDrawing(dr, s.cx * k, s.cy * k));
      }
      return p;
    });
  }));
}
d.set('word/document.xml', parts.map(p => p.chunk).join('')); saveDocx(d, f);
console.log(mode === 'multi' ? `multi-picture paragraphs fitted on one line: ${n}` : mode === 'enlarge' ? `images resized: ${n} (enlarged: ${grow})` : `images scaled to fit their cell: ${n}`);
