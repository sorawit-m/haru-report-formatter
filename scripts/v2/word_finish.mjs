// แทน word_finish.ps1 (ไม่ต้องเปิด Word): สารบัญ 16pt (บรรทัดบทหนา ไม่เอียง), ปก (ระยะชื่อเรื่อง 12pt, บรรทัดโครงการ 1.3,
// รูปพื้นหลังเต็มหน้า), ตัวอักษรตารางทดสอบ / ตารางอื่น 16pt — สารบัญและฟิลด์อัปเดตตอนส่งออก PDF (scripts/mac/word.sh)
// usage: node word_finish.mjs <file.docx> [TestFont=0] [TestHead=0] [PicText=0]
import { loadDocx, saveDocx } from '../lib/docx.mjs';
import { sections, mapTables, gridCols, hasNestedTable, mapCells, setRunSize, textOfXml, EMU_PER_TWIP, setPPr, getPPr, setRPrIn, removeRPrIn, coverEnd } from '../lib/wordxml.mjs';
const [f, TF = '0', TH = '0', PT = '0'] = process.argv.slice(2);
const TestFont = +TF, TestHead = +TH, PicText = +PT;
const d = loadDocx(f); const cnt = { tocStyles: 0, coverPara: 0, coverBg: 0, tables16: 0, testTables: 0, tocLines: 0 };

// ---- สไตล์สารบัญ (toc 1-9) + สารบัญรูป/ตาราง: 16pt ไม่เอียง, toc 1 หนา
let st = d.get('word/styles.xml');
st = st.replace(/<w:style\b[^>]*>[\s\S]*?<\/w:style>/g, s => {
  const nm = ((s.match(/<w:name w:val="([^"]+)"/) || [])[1] || '').toLowerCase();
  const m = nm.match(/^toc ([1-9])$/); if (!m && nm !== 'table of figures') return s;
  cnt.tocStyles++; const bold = m && m[1] === '1';
  const rPr = (s.match(/<w:rPr>[\s\S]*?<\/w:rPr>|<w:rPr\/>/) || [''])[0];
  let r = removeRPrIn(rPr || '<w:rPr/>', ['b', 'bCs', 'i', 'iCs']);
  for (const t of bold ? ['<w:b/>', '<w:bCs/>'] : []) r = setRPrIn(r, t);
  for (const t of ['<w:sz w:val="32"/>', '<w:szCs w:val="32"/>']) r = setRPrIn(r, t);
  s = s.replace(/<w:autoRedefine(?: [^>]*)?\/>/g, ''); // กัน Word เขียนทับสไตล์ด้วยรูปแบบในสารบัญตอนอัปเดต
  return rPr ? s.replace(rPr, r) : s.replace('</w:style>', r + '</w:style>');
});
d.set('word/styles.xml', st);
const tocIds = new Set([...st.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>[\s\S]*?<w:name w:val="(toc [1-9]|table of figures)"/gi)].map(m => m[1]));

let x = d.get('word/document.xml');
// ---- บรรทัดสารบัญที่มีอยู่แล้ว: ตัดตัวเอียง/ขนาดที่ตั้งเอง
x = x.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
  const s = (p.match(/<w:pStyle w:val="([^"]+)"/) || [])[1]; if (!s || !tocIds.has(s)) return p;
  cnt.tocLines++; return setRunSize(p.replace(/<w:r(?: [^>]*)?><w:rPr>[\s\S]*?<\/w:rPr>/g, r => r.replace(/<w:rPr>[\s\S]*<\/w:rPr>/, x => removeRPrIn(x, ['i', 'iCs']))), 32);
});

// ---- ปก (หน้าแรก: ถึง page break แรก หรือจบ section แรก)
{
  const ce = coverEnd(x); const sect = sections(x)[0].sect; let first = true; const cnt0 = cnt.coverPara;
  let cover = x.slice(0, ce).replace(/<w:p[ >][\s\S]*?<\/w:p>/g, p => {
    const t = textOfXml(p).trim(); if (t.length <= 1) return p;
    if (first) { first = false; cnt.coverPara++; p = setSpacing(p, { before: 240 }); }
    if (t.includes('โครงการ')) { cnt.coverPara++; p = setSpacing(p, { line: 312, lineRule: 'auto' }); }
    return p;
  });
  // รูปพื้นหลังปก (anchor กว้าง >= 90% หน้า) -> เต็มหน้าพอดี
  const pw = sect.w * EMU_PER_TWIP, ph = sect.h * EMU_PER_TWIP;
  cover = cover.replace(/<wp:anchor\b[\s\S]*?<\/wp:anchor>/g, a => {
    const m = a.match(/<wp:extent cx="(\d+)" cy="(\d+)"/); if (!m || +m[1] < pw * 0.9) return a;
    cnt.coverBg++;
    return a.replace(/<wp:positionH\b[\s\S]*?<\/wp:positionH>/, '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>')
      .replace(/<wp:positionV\b[\s\S]*?<\/wp:positionV>/, '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>')
      .replace(/<wp:extent cx="\d+" cy="\d+"/, `<wp:extent cx="${pw}" cy="${ph}"`).replace(/<a:ext cx="\d+" cy="\d+"/g, `<a:ext cx="${pw}" cy="${ph}"`);
  });
  // ย่อหน้าที่มีแต่ตัวขึ้นหน้าใหม่ท้ายปก: ถ้าปกเต็มหน้า บรรทัดนี้ล้นไปเป็นหน้าว่าง -> ลบ แล้วให้ย่อหน้าถัดไปขึ้นหน้าใหม่แทน (ผลเหมือนเดิม)
  const lastP = [...cover.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].pop();
  let rest = x.slice(ce);
  if (lastP && lastP.index + lastP[0].length === cover.length && !textOfXml(lastP[0]).trim() && /<w:br [^>]*w:type="page"/.test(lastP[0]) && !/<w:drawing|<w:sectPr|<w:pict/.test(lastP[0])) {
    const nx = rest.match(/^\s*(<w:p[ >][\s\S]*?<\/w:p>)/);
    if (nx) { cover = cover.slice(0, lastP.index); rest = rest.replace(nx[1], setPPr(nx[1], '<w:pageBreakBefore/>')); cnt.coverBreak = 1; }
  }
  x = cover + rest;
}

// ---- ตัวอักษรในตาราง
x = mapTables(x, t => {
  if (hasNestedTable(t)) return t;
  const cols = gridCols(t).length;
  if (cols >= 7 && TestFont > 0) {
    cnt.testTables++;
    return mapCells(t, ({ tc, row }) => {
      const pic = /<w:drawing/.test(tc);
      if (row > 0 && !pic) return setRunSize(tc, TestFont * 2);
      if (row > 0 && pic && PicText > 0) return setRunSize(tc, PicText * 2);
      if (row === 0 && (TestHead > 0 || TestFont === 12)) return setRunSize(tc, (TestHead > 0 ? TestHead : 14) * 2);
      return tc;
    });
  }
  if (cols > 1) {
    const sizes = new Set([...t.matchAll(/<w:r(?: [^>]*)?><w:rPr>(?:(?!<\/w:rPr>)[\s\S])*?<w:sz w:val="(\d+)"/g)].map(m => m[1]));
    if (sizes.size === 1 && sizes.has('32')) return t;
    cnt.tables16++; return setRunSize(t, 32);
  }
  return t;
});
d.set('word/document.xml', x); saveDocx(d, f);
console.log(`tables set to 16pt: ${cnt.tables16}; test tables at ${TestFont}pt: ${cnt.testTables}; toc styles: ${cnt.tocStyles}; toc lines: ${cnt.tocLines}; cover: ${cnt.coverPara} para, ${cnt.coverBg} background${cnt.coverBreak ? ', page break moved to next page' : ''}`);

function setSpacing(p, { before, line, lineRule }) {
  let n = getPPr(p, 'spacing') || '<w:spacing/>';
  const setA = (s, k, v) => new RegExp(` w:${k}="`).test(s) ? s.replace(new RegExp(` w:${k}="[^"]*"`), ` w:${k}="${v}"`) : s.replace('<w:spacing', `<w:spacing w:${k}="${v}"`);
  if (before !== undefined) { n = setA(n, 'before', before); n = n.replace(/ w:beforeLines="[^"]*"| w:beforeAutospacing="[^"]*"/g, ''); }
  if (line !== undefined) { n = setA(n, 'line', line); n = setA(n, 'lineRule', lineRule); }
  return setPPr(p, n);
}
