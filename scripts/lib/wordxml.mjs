// ตัวช่วยแก้ document.xml แบบไม่ต้องเปิด Word (ใช้แทนขั้น Word COM บน Mac และใช้ได้บน Windows ด้วย)
// หน่วย: twips (1/20 pt) ; EMU = twips * 635

export const EMU_PER_TWIP = 635;

// ตัดเนื้อหาเป็นช่วงตาม section: [{ chunk, sect: { w, h, left, right, top, bottom, textW, textH } }]
export function sections(x) {
  const parts = []; const re = /<w:sectPr\b[\s\S]*?<\/w:sectPr>/g; let m, last = 0;
  while ((m = re.exec(x))) { parts.push({ chunk: x.slice(last, m.index + m[0].length), sect: sectInfo(m[0]) }); last = m.index + m[0].length; }
  const tailSect = parts.length ? parts[parts.length - 1].sect : sectInfo('');
  parts.push({ chunk: x.slice(last), sect: tailSect });
  // ย่อหน้าก่อน sectPr ตัวแรก ใช้ sectPr ตัวนั้น (ถูกแล้ว: sectPr จบ section ที่อยู่ก่อนหน้า)
  return parts;
}
export function sectInfo(s) {
  const g = (re, d) => { const m = s.match(re); return m ? +m[1] : d; };
  const w = g(/<w:pgSz [^>]*w:w="(\d+)"/, 11906), h = g(/<w:pgSz [^>]*w:h="(\d+)"/, 16838);
  const left = g(/<w:pgMar [^>]*w:left="(-?\d+)"/, 1440), right = g(/<w:pgMar [^>]*w:right="(-?\d+)"/, 1440);
  const top = Math.abs(g(/<w:pgMar [^>]*w:top="(-?\d+)"/, 1440)), bottom = Math.abs(g(/<w:pgMar [^>]*w:bottom="(-?\d+)"/, 1440));
  return { w, h, left, right, top, bottom, textW: w - left - right, textH: h - top - bottom };
}

// ตารางระดับบนสุด (ไม่ซ้อน) ในข้อความ -> เรียก fn(tableXml) แล้วแทนที่
export function mapTables(x, fn) {
  let out = '', i = 0;
  while (true) {
    const s = x.indexOf('<w:tbl>', i); if (s < 0) { out += x.slice(i); break; }
    let depth = 0, j = s, e = -1; const re = /<w:tbl>|<\/w:tbl>/g; re.lastIndex = s; let m;
    while ((m = re.exec(x))) { depth += m[0] === '<w:tbl>' ? 1 : -1; if (depth === 0) { e = m.index + m[0].length; break; } }
    if (e < 0) { out += x.slice(i); break; }
    out += x.slice(i, s) + fn(x.slice(s, e)); i = e; j = e;
  }
  return out;
}
export const gridCols = t => [...((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].matchAll(/<w:gridCol w:w="(\d+)"/g))].map(a => +a[1]);
export const hasNestedTable = t => t.indexOf('<w:tbl>', 7) >= 0;

// ขอบซ้าย-ขวาของช่อง (twips): tcMar > tblCellMar > ค่าเริ่มต้นของ Word 108
export function cellPads(tbl, tc) {
  const pick = (src, side) => { const m = src.match(new RegExp(`<w:(?:${side})\\s+w:w="(\\d+)"`)); return m ? +m[1] : null; };
  const tcMar = (tc.match(/<w:tcMar>[\s\S]*?<\/w:tcMar>/) || [''])[0];
  const tblMar = (tbl.match(/<w:tblCellMar>[\s\S]*?<\/w:tblCellMar>/) || [''])[0];
  const l = pick(tcMar, 'left|start') ?? pick(tblMar, 'left|start') ?? 108;
  const r = pick(tcMar, 'right|end') ?? pick(tblMar, 'right|end') ?? 108;
  return l + r;
}

// เดินทุกช่องของตาราง (ระดับเดียว) พร้อมความกว้างช่องจาก grid: fn({ tc, row, col, width, pads }) -> tc ใหม่
export function mapCells(tbl, fn) {
  const grid = gridCols(tbl); let row = 0;
  return tbl.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, tr => {
    let col = +((tr.match(/<w:gridBefore w:val="(\d+)"/) || [])[1] || 0); const r = row++;
    return tr.replace(/<w:tc>[\s\S]*?<\/w:tc>/g, tc => {
      const span = +((tc.match(/<w:gridSpan w:val="(\d+)"/) || [])[1] || 1);
      let width = grid.slice(col, col + span).reduce((a, b) => a + b, 0);
      const tcW = tc.match(/<w:tcW w:w="(\d+)" w:type="dxa"/); if (!width && tcW) width = +tcW[1];
      const res = fn({ tc, row: r, col, span, width, pads: cellPads(tbl, tc) }); col += span; return res ?? tc;
    });
  });
}

// รูปแบบ inline ในย่อหน้า: ขนาด (EMU) และตั้งขนาดใหม่ทั้ง wp:extent และ a:ext ให้ตรงกัน
export const inlineDrawings = p => [...p.matchAll(/<wp:inline\b[\s\S]*?<\/wp:inline>/g)].map(m => m[0]);
export function drawingSize(d) { const m = d.match(/<wp:extent cx="(\d+)" cy="(\d+)"/); return m ? { cx: +m[1], cy: +m[2] } : null; }
export function resizeDrawing(d, cx, cy) {
  cx = Math.max(1, Math.round(cx)); cy = Math.max(1, Math.round(cy));
  return d.replace(/<wp:extent cx="\d+" cy="\d+"/, `<wp:extent cx="${cx}" cy="${cy}"`).replace(/<a:ext cx="\d+" cy="\d+"/g, `<a:ext cx="${cx}" cy="${cy}"`);
}

// ตั้งขนาดตัวอักษร (half-points) ให้ทุก run ในข้อความ xml (ใส่ตามลำดับ schema)
export function setRunSize(x, hp) {
  return x.replace(/(<w:r(?: [^>]*)?>)(<w:rPr>[\s\S]*?<\/w:rPr>|<w:rPr\/>)?/g, (m, open, rPr) => {
    let r = rPr || '<w:rPr/>';
    r = setRPrIn(r, `<w:sz w:val="${hp}"/>`); r = setRPrIn(r, `<w:szCs w:val="${hp}"/>`);
    return open + r;
  });
}

export const textOfXml = s => [...s.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('');
export const norm = s => s.replace(/[\s​⁠ ]+/g, '');

// ใส่ property ลงใน pPr ตามลำดับ schema — ไม่ทับถ้ามีแล้ว
export const addPPr = (p, tag) => setPPr(p, tag, false);

// ---- ใส่ลูกใน pPr / rPr ตามลำดับที่ schema กำหนด (Word ฟ้องไฟล์เสียถ้าลำดับผิด)
const P_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
const R_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange'];
// แยกลูกระดับแรกของ <w:pPr>…</w:pPr> หรือ <w:rPr>…</w:rPr>
function children(inner) {
  const out = []; let i = 0;
  while (i < inner.length) {
    const s = inner.indexOf('<w:', i); if (s < 0) break;
    const name = inner.slice(s + 3).match(/^[A-Za-z]+/)[0];
    const gt = inner.indexOf('>', s);
    if (inner[gt - 1] === '/') { out.push({ name, xml: inner.slice(s, gt + 1) }); i = gt + 1; continue; }
    // หา tag ปิดคู่ (ซ้อนชื่อเดียวกันได้ เช่น rPr ใน pPr)
    let depth = 0; const re = new RegExp(`<w:${name}(?=[\\s/>])[^>]*?(/?)>|</w:${name}>`, 'g'); re.lastIndex = s; let m, e = inner.length;
    while ((m = re.exec(inner))) { if (m[0].startsWith('</')) { depth--; if (depth === 0) { e = m.index + m[0].length; break; } } else if (!m[1]) depth++; }
    out.push({ name, xml: inner.slice(s, e) }); i = e;
  }
  return out;
}
function putOrdered(inner, tagXml, order, replace = true) {
  const name = tagXml.match(/^<w:([A-Za-z]+)/)[1]; const kids = children(inner);
  const at = kids.findIndex(k => k.name === name);
  if (at >= 0) { if (replace) kids[at].xml = tagXml; return kids.map(k => k.xml).join(''); }
  const rank = n => { const r = order.indexOf(n); return r < 0 ? order.length : r; };
  const pos = kids.findIndex(k => rank(k.name) > rank(name));
  if (pos < 0) kids.push({ name, xml: tagXml }); else kids.splice(pos, 0, { name, xml: tagXml });
  return kids.map(k => k.xml).join('');
}
// ตั้ง/แทน property ของย่อหน้า (p = <w:p …>…</w:p>)
export function setPPr(p, tagXml, replace = true) {
  const m = p.match(/^(<w:p(?: [^>]*)?>)(<w:pPr>([\s\S]*?)<\/w:pPr>|<w:pPr\/>)?/);
  const inner = m[3] || ''; const ni = putOrdered(inner, tagXml, P_ORDER, replace);
  return m[1] + `<w:pPr>${ni}</w:pPr>` + p.slice(m[0].length);
}
export const getPPr = (p, name) => { const m = p.match(/^<w:p(?: [^>]*)?>(?:<w:pPr>([\s\S]*?)<\/w:pPr>)?/); return (children(m && m[1] || '').find(k => k.name === name) || {}).xml; };
// ตั้ง/แทน property ใน rPr (ข้อความ <w:rPr>…</w:rPr>)
export function setRPrIn(rPr, tagXml, replace = true) {
  const inner = rPr === '<w:rPr/>' ? '' : rPr.replace(/^<w:rPr>|<\/w:rPr>$/g, '');
  return `<w:rPr>${putOrdered(inner, tagXml, R_ORDER, replace)}</w:rPr>`;
}
export function removeRPrIn(rPr, names) {
  const inner = rPr === '<w:rPr/>' ? '' : rPr.replace(/^<w:rPr>|<\/w:rPr>$/g, '');
  return `<w:rPr>${children(inner).filter(k => !names.includes(k.name)).map(k => k.xml).join('')}</w:rPr>`;
}

// ปลายหน้าปก (index ใน document.xml): ปกจบที่ตัวขึ้นหน้าใหม่ตัวแรก หรือจบ section แรก แล้วแต่อันไหนมาก่อน
// (ปก สารบัญ ทะเบียนเวอร์ชัน มักอยู่ section เดียวกันคั่นด้วย page break — ห้ามนับทั้ง section เป็นปก)
// เอกสาร section เดียวที่ไม่มี page break -> ไม่มีปก (คืนตำแหน่งต้น body)
export function coverEnd(x) {
  const b = x.indexOf('<w:body>') + 8; const bodyEnd = x.lastIndexOf('</w:body>');
  const cands = [];
  const s = x.indexOf('<w:sectPr', b); if (s >= 0) { const pe = x.indexOf('</w:p>', s); const be = x.lastIndexOf('<w:sectPr', bodyEnd); if (s !== be && pe >= 0) cands.push(pe + 6); }
  const br = x.slice(b).search(/<w:br [^>]*w:type="page"/); if (br >= 0) { const pe = x.indexOf('</w:p>', b + br); if (pe >= 0) cands.push(pe + 6); }
  const pb = x.slice(b).search(/<w:pageBreakBefore(?: w:val="(?:1|true|on)")?\/>/); if (pb >= 0) { const ps = x.lastIndexOf('<w:p', b + pb); if (ps > b) cands.push(ps); }
  return cands.length ? Math.min(...cands) : b;
}
