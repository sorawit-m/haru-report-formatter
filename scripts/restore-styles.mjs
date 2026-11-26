// restore-styles.mjs — หลัง Word บันทึกไฟล์ บางครั้ง Style หัวเรื่องถูกผูกกับเลขรายการอัตโนมัติ (เช่น "ส่วน 2.03")
// ที่ต้นฉบับไม่มี -> เอาเลขรายการออกจาก Style ที่ต้นฉบับไม่มี (เทียบตามชื่อ Style)
// ใช้: node restore-styles.mjs <ต้นฉบับ.docx> <ไฟล์ที่จัดแล้ว.docx>
import { loadDocx, saveDocx, parseXml, serialize, all, els, remove } from './lib/docx.mjs';
const [, , origF, outF] = process.argv;
if (!origF || !outF) { console.error('usage: node restore-styles.mjs original.docx formatted.docx'); process.exit(1); }
const styleMap = xml => { const m = {}; for (const s of xml.matchAll(/<w:style\b[^>]*w:type="paragraph"[^>]*>([\s\S]*?)<\/w:style>/g)) { const n = (s[1].match(/<w:name w:val="([^"]*)"/) || [])[1]; if (n) m[n] = s[0]; } return m; };
const origDoc = loadDocx(origF);
const orig = styleMap(origDoc.get('word/styles.xml') || '');
const origNumbering = origDoc.get('word/numbering.xml') || '';
// Word renumbers style IDs when it saves (e.g. "heading 1" 10 -> 1) and can attach automatic numbering to a built-in style
// ("บทความ III." in front of "เรื่อง … หน้า" on the TOC pages). Compare by style NAME and strip the numbering only when it is
// certain the original had none: the original style has no numPr AND no numbering level in the original links to it.
// (Real heading numbers like "1.1" have numPr or a numbering.xml link in the original, so they are left alone.)
const doc = loadDocx(outF); let n = 0;
if (process.env.RF_STRIP_STYLE_NUMBERING === '1') {   // OFF: tested on a real book it removed real "1.1" numbers too
  let st = doc.get('word/styles.xml') || '';
  st = st.replace(/<w:style\b[^>]*w:type="paragraph"[^>]*>[\s\S]*?<\/w:style>/g, s => {
    if (!/<w:numPr>/.test(s)) return s;
    const name = (s.match(/<w:name w:val="([^"]*)"/) || [])[1]; const o = name && orig[name]; if (!o) return s;
    if (/<w:numPr>/.test(o)) return s;
    const oid = (o.match(/w:styleId="([^"]*)"/) || [])[1];
    if (oid && new RegExp(`<w:pStyle w:val="${oid.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`).test(origNumbering)) return s;
    n++; return s.replace(/<w:numPr>[\s\S]*?<\/w:numPr>/, '');
  });
  if (n) doc.set('word/styles.xml', st);
}
// a table row must not "keep with next" into a row that holds a picture (screenshot): together they do not fit one page,
// so Word pushes the whole block and leaves the heading alone on a page
const root = parseXml(doc.get('word/document.xml')); let k = 0;
// Word sometimes attaches automatic numbering to a built-in style while saving ("บทความ I." in front of "เรื่อง … หน้า").
// Per PARAGRAPH (never per style): a paragraph without its own numPr whose style had no numbering in the original
// (by style name, incl. basedOn chain and numbering.xml links) but has numbering now -> switch numbering off for it only.
let numOff = 0;
{
  const parse = xml => { const byId = {}; for (const m of xml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) { const id = (m[1].match(/w:styleId="([^"]*)"/) || [])[1]; if (!id) continue; byId[id] = { name: (m[2].match(/<w:name w:val="([^"]*)"/) || [])[1], base: (m[2].match(/<w:basedOn w:val="([^"]*)"/) || [])[1], num: /<w:numPr>[\s\S]*?<w:numId w:val="(?!0")\d+"/.test(m[2]), ind: (m[2].match(/<w:ind [^>]*\/>/) || [])[0] }; } return byId; };
  const O = parse(origDoc.get('word/styles.xml') || ''), N = parse(doc.get('word/styles.xml') || '');
  const oByName = {}; for (const [id, s] of Object.entries(O)) if (s.name) oByName[s.name] = id;
  const chainHas = (map, id, numXml) => { for (let i = 0; id && i < 12; i++, id = map[id] && map[id].base) { if (!map[id]) break; if (map[id].num) return true; if (numXml && new RegExp(`<w:pStyle w:val="${id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"`).test(numXml)) return true; } return false; };
  for (const p of all(root, 'w:p')) {
    const pPr = els(p, 'w:pPr')[0]; if (!pPr) continue;
    if (els(pPr, 'w:numPr').length) continue;
    const ps = els(pPr, 'w:pStyle')[0]; if (!ps) continue;
    const sid = ps.attrs.find(a => a[0] === 'w:val')[1]; const s = N[sid]; if (!s || !s.name) continue;
    const oid = oByName[s.name]; if (!oid) continue;
    if (!chainHas(N, sid, null) || chainHas(O, oid, origNumbering)) continue;
    const np = { type: 'el', name: 'w:numPr', attrs: [], children: [], parent: pPr };
    np.children.push({ type: 'el', name: 'w:ilvl', attrs: [['w:val', '0']], children: [], parent: np }, { type: 'el', name: 'w:numId', attrs: [['w:val', '0']], children: [], parent: np });
    const at = pPr.children.findIndex(c => c.type === 'el' && !['w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr', 'w:widowControl'].includes(c.name));
    if (at < 0) pPr.children.push(np); else pPr.children.splice(at, 0, np);
    // the style also lost its own indent in the save ("เรื่อง" pushed right, "หน้า" wraps) -> put the original indent back
    if (O[oid].ind && !els(pPr, 'w:ind').length && O[oid].ind !== s.ind) {
      const m = O[oid].ind.match(/<w:ind ([^/]*)\/>/); const attrs = [...m[1].matchAll(/([\w:]+)="([^"]*)"/g)].map(x => [x[1], x[2]]);
      const ind = { type: 'el', name: 'w:ind', attrs, children: [], parent: pPr };
      const j = pPr.children.findIndex(c => c.type === 'el' && ['w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc', 'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange'].includes(c.name));
      if (j < 0) pPr.children.push(ind); else pPr.children.splice(j, 0, ind);
    }
    numOff++;
  }
}
console.log(`numbers Word attached to un-numbered paragraphs switched off: ${numOff}`);

// long numbered body paragraphs that wear a heading style ("2.4.1 กระบวนการ…: …" 3–4 lines) inherit "keep with next" and
// "keep lines together" from the heading style -> 2.4.1…2.4.4 + the following picture move as ONE block and leave half a
// page empty. They are body text: switch both off on the paragraph (the style and real short headings are unchanged).
let unchained = 0;
function unchainLongNumberedBody() {
  const sMap = {}; for (const m of (doc.get('word/styles.xml') || '').matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) { const id = (m[1].match(/w:styleId="([^"]*)"/) || [])[1]; if (!id) continue; sMap[id] = { base: (m[2].match(/<w:basedOn w:val="([^"]*)"/) || [])[1], kn: /<w:keepNext\/>|<w:keepNext w:val="(1|true|on)"/.test(m[2]), kl: /<w:keepLines\/>|<w:keepLines w:val="(1|true|on)"/.test(m[2]) }; }
  const inh = (id, k) => { for (let i = 0; id && i < 12; i++, id = sMap[id] && sMap[id].base) if (sMap[id] && sMap[id][k]) return true; return false; };
  const ORDER = ['w:pStyle', 'w:keepNext', 'w:keepLines'];
  const setOff = (pPr, name) => {
    let e = els(pPr, name)[0];
    if (e) { e.attrs = [['w:val', '0']]; return; }
    e = { type: 'el', name, attrs: [['w:val', '0']], children: [], parent: pPr };
    const after = pPr.children.findIndex(c => c.type === 'el' && !ORDER.slice(0, ORDER.indexOf(name) + 1).includes(c.name));
    if (after < 0) pPr.children.push(e); else pPr.children.splice(after, 0, e);
  };
  for (const p of all(root, 'w:p', [], 'w:tbl')) {
    const pPr = els(p, 'w:pPr')[0]; if (!pPr) continue;
    const ps = els(pPr, 'w:pStyle')[0]; if (!ps) continue; const sid = ps.attrs.find(a => a[0] === 'w:val')[1];
    if (resolveOut(sid) >= 9) continue;
    const t = all(p, 'w:t').map(x => x.children.map(c => c.text || '').join('')).join('');
    if (t.length < 100) continue;
    if (all(p, 'w:drawing').length) continue;
    let did = false;
    if (inh(sid, 'kn') || els(pPr, 'w:keepNext').length) { setOff(pPr, 'w:keepNext'); did = true; }
    if (inh(sid, 'kl') || els(pPr, 'w:keepLines').length) { setOff(pPr, 'w:keepLines'); did = true; }
    if (did) unchained++;
  }
}
// deepest level the table of contents lists: TOC \o "1-2" -> 2 (outline index 0..1). No \o switch -> treat all levels as listed.
const tocM = (doc.get('word/document.xml') || '').match(/TOC [^<]*\\o\s*&quot;(\d)-(\d)&quot;|TOC [^<]*\\o\s*"(\d)-(\d)"/);
const TOC_MAX = tocM ? +(tocM[2] || tocM[4]) : 9;
for (const tbl of all(root, 'w:tbl')) {
  const trs = els(tbl, 'w:tr');
  for (let i = 0; i < trs.length - 1; i++) {
    if (!(all(trs[i + 1], 'w:drawing').length || all(trs[i + 1], 'w:pict').length)) continue;
    const kn = all(trs[i], 'w:keepNext'); if (kn.length) { kn.forEach(remove); k++; }
  }
}
// long body paragraphs that carry a heading outline level (they show up as entries in the table of contents)
const styles = doc.get('word/styles.xml') || ''; const sOut = {}, sBase = {}; let defStyle = null;
for (const m of styles.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) { const id = (m[1].match(/w:styleId="([^"]*)"/) || [])[1]; if (!id) continue; if (/w:default="1"/.test(m[1]) && /w:type="paragraph"/.test(m[1])) defStyle = id; const o = (m[2].match(/<w:outlineLvl w:val="(\d)"/) || [])[1]; if (o !== undefined) sOut[id] = +o; const b = (m[2].match(/<w:basedOn w:val="([^"]*)"/) || [])[1]; if (b) sBase[id] = b; }
const resolveOut = id => { for (let i = 0; id && i < 10; i++, id = sBase[id]) if (sOut[id] !== undefined) return sOut[id]; return 9; };
let demoted = 0, typedOff = 0;
for (const p of all(root, 'w:p')) {
  let inT = false; for (let q = p.parent; q; q = q.parent) if (q.name === 'w:tc') { inT = true; break; } if (inT) continue;
  const t = all(p, 'w:t').map(x => x.children.map(c => c.text).join('')).join('').trim();
  const pPr = els(p, 'w:pPr')[0];
  // heading that already starts with its own typed number ("บทที่ 1 …", "2.3 …") -> switch off the automatic list number
  if (pPr && t.length < 150 && /^(บทที่\s*\d+|\d+(\.\d+)+\.?|[A-Z]{2,6}[-_]\d{2,4})\s|^(สารบัญ|สารบัญรูปภาพ|สารบัญตาราง)$/.test(t)) {
    const hasStyle = els(pPr, 'w:pStyle')[0];
    if (hasStyle && !els(pPr, 'w:numPr').length) {
      const np = { type: 'el', name: 'w:numPr', attrs: [], children: [], parent: pPr };
      np.children = [{ type: 'el', name: 'w:ilvl', attrs: [['w:val', '0']], children: [], parent: np }, { type: 'el', name: 'w:numId', attrs: [['w:val', '0']], children: [], parent: np }];
      const idx = pPr.children.findIndex(c => c.type === 'el' && !['w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr', 'w:widowControl'].includes(c.name));
      if (idx < 0) pPr.children.push(np); else pPr.children.splice(idx, 0, np);
      typedOff++;
    }
    continue;
  }
  if (t.length < 100) continue;
  // body text that wears a heading STYLE is picked up by the TOC even with outlineLvl 9 -> give it the previous paragraph's style
  {
    const psEl = pPr && els(pPr, 'w:pStyle')[0]; const sid = psEl && psEl.attrs.find(a => a[0] === 'w:val')[1];
    // only levels the TOC actually lists (TOC \o "1-2" -> outline 0..1); deeper numbered items such as "1.2.1 …"
    // never appear in it and must keep their style (it carries their automatic number)
    if (sid && resolveOut(sid) < TOC_MAX) {
      const sib = p.parent.children; let q = sib.indexOf(p) - 1; let prevStyle = null;
      for (; q >= 0; q--) { const c = sib[q]; if (c.type !== 'el' || c.name !== 'w:p') continue; const pp = els(c, 'w:pPr')[0]; const s2 = pp && els(pp, 'w:pStyle')[0]; const v = s2 ? s2.attrs.find(a => a[0] === 'w:val')[1] : defStyle; if (resolveOut(v) >= 9) { prevStyle = v; break; } }
      if (prevStyle) { psEl.attrs = [['w:val', prevStyle]]; demoted++; continue; }
    }
  } const ps = pPr && els(pPr, 'w:pStyle')[0]; const ol = pPr && els(pPr, 'w:outlineLvl')[0];
  const lvl = ol ? +ol.attrs.find(a => a[0] === 'w:val')[1] : resolveOut(ps ? ps.attrs.find(a => a[0] === 'w:val')[1] : defStyle);
  if (lvl >= TOC_MAX) continue;
  if (ol) ol.attrs = [['w:val', '9']]; else { const n = { type: 'el', name: 'w:outlineLvl', attrs: [['w:val', '9']], children: [], parent: pPr }; const idx = pPr.children.findIndex(c => c.type === 'el' && ['w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange'].includes(c.name)); if (idx < 0) pPr.children.push(n); else pPr.children.splice(idx, 0, n); }
  demoted++;
}
console.log(`long body paragraphs removed from the table of contents: ${demoted}; duplicate auto-numbers switched off on typed-number headings: ${typedOff}`);
unchainLongNumberedBody();
console.log(`long numbered body paragraphs no longer chained to the next one: ${unchained}`);
doc.set('word/document.xml', serialize(root)); saveDocx(doc, outF);
console.log(`styles: list numbering removed from ${n} style(s) that had none in the original; rows un-chained from picture rows: ${k}`);
