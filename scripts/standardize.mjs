// standardize.mjs — ขั้นแรก: ตั้งค่าไฟล์ทุกเล่มให้เหมือนกัน ก่อนจัดความสวยงาม (ผู้ใช้สั่ง 25-11-2026)
//   หน้ากระดาษ/ขอบ/หัว-ท้ายกระดาษ, ฟอนต์และระยะของ Normal, Style หัวข้อทุกระดับ (ทั้ง "heading N" ของ Word และ "หัวข้อ N"),
//   สารบัญ / สารบัญรูป-ตาราง / คำบรรยาย ตามค่า "standard" ใน config.json
//   ไม่แตะ: เนื้อหา, ตัวเลข/ลำดับหัวข้อ (numPr), แนวกระดาษ (แนวนอนคงแนวนอน), หน้าปก
// ใช้: node standardize.mjs <in.docx> <out.docx> [config.json]
import fs from 'node:fs';
import { loadDocx, saveDocx } from './lib/docx.mjs';
const [, , inF, outF, cfgF] = process.argv;
if (!inF || !outF) { console.error('usage: node standardize.mjs in.docx out.docx [config.json]'); process.exit(1); }
const cfg = cfgF && fs.existsSync(cfgF) ? JSON.parse(fs.readFileSync(cfgF, 'utf8')) : {};
const STD = cfg.standard || {};
const FONT = cfg.heading_font || 'TH SarabunPSK';
const d = loadDocx(inF);
let st = d.get('word/styles.xml'); let doc = d.get('word/document.xml');
const log = [];

// ---------- 1) page setup: margins / header / footer distance (orientation & size kept) ----------
const cmTw = v => String(Math.round(v * 567));
let sectChanged = 0;
// the COVER section (first of several) keeps its own margins: its full-page background picture is positioned from the
// margins, so changing them left white strips at the paper edges
const nSect = (doc.match(/<w:sectPr\b/g) || []).length; let sIdx = 0;
doc = doc.replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, sp => {
  if (sIdx++ === 0 && nSect > 1) return sp;
  const land = /w:orient="landscape"/.test(sp);
  const pm = (land ? STD.landscape_margins_cm : STD.page_margins_cm) || {};
  return sp.replace(/<w:pgMar\b[^>]*\/>/, m => {
    let r = m;
    for (const k of ['top', 'bottom', 'left', 'right', 'header', 'footer']) {
      const v = pm[k]; if (v == null) continue; const re = new RegExp(`w:${k}="-?\\d+"`);
      r = re.test(r) ? r.replace(re, `w:${k}="${cmTw(v)}"`) : r.replace(/\/>$/, ` w:${k}="${cmTw(v)}"/>`);
    }
    if (r !== m) sectChanged++;
    return r;
  });
});log.push(`page margins set in ${sectChanged} section(s)`);

// ---------- 2) styles by NAME ----------
const styleRe = name => new RegExp(`<w:style\\b[^>]*>(?:(?!<\\/w:style>)[\\s\\S])*?<w:name w:val="${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}"\\/>[\\s\\S]*?<\\/w:style>`, 'i');
const ORDER_P = ['w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr', 'w:widowControl', 'w:numPr', 'w:suppressLineNumbers', 'w:pBdr', 'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap', 'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN', 'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind', 'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc', 'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange'];
const ORDER_R = ['w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps', 'w:strike', 'w:dstrike', 'w:outline', 'w:shadow', 'w:emboss', 'w:imprint', 'w:noProof', 'w:snapToGrid', 'w:vanish', 'w:webHidden', 'w:color', 'w:spacing', 'w:w', 'w:kern', 'w:position', 'w:sz', 'w:szCs', 'w:highlight', 'w:u', 'w:effect', 'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang', 'w:eastAsianLayout', 'w:specVanish', 'w:oMath'];
// put / replace one element inside <w:pPr> or <w:rPr> keeping schema order
function setIn(block, tag, xml, order) {
  const name = xml.match(/^<([\w:]+)/)[1];
  const open = `<${tag}>`, close = `</${tag}>`;
  let i = block.s.indexOf(open);
  if (i < 0) {   // create container: pPr before rPr, rPr at end (before tblPr etc.)
    const where = tag === 'w:pPr' ? (block.s.search(/<w:rPr>|<w:tblPr>|<w:trPr>|<w:tcPr>|<w:tblStylePr/) >= 0 ? block.s.search(/<w:rPr>|<w:tblPr>|<w:trPr>|<w:tcPr>|<w:tblStylePr/) : block.s.length)
      : (block.s.search(/<w:tblPr>|<w:trPr>|<w:tcPr>|<w:tblStylePr/) >= 0 ? block.s.search(/<w:tblPr>|<w:trPr>|<w:tcPr>|<w:tblStylePr/) : block.s.length);
    block.s = block.s.slice(0, where) + open + close + block.s.slice(where); i = block.s.indexOf(open);
  }
  const j = block.s.indexOf(close, i);
  let inner = block.s.slice(i + open.length, j);
  const own = new RegExp(`<${name}(?:\\s[^>]*)?(?:\\/>|>[\\s\\S]*?<\\/${name}>)`);
  if (own.test(inner)) inner = inner.replace(own, xml);
  else {
    const my = order.indexOf(name); const kids = [...inner.matchAll(/<(w:[\w]+)\b/g)];
    const after = kids.find(k => order.indexOf(k[1]) > my || order.indexOf(k[1]) < 0 && false);
    if (after) inner = inner.slice(0, after.index) + xml + inner.slice(after.index); else inner += xml;
  }
  block.s = block.s.slice(0, i + open.length) + inner + block.s.slice(j);
}
function applyLook(name, look) {
  const re = styleRe(name); const m = st.match(re); if (!m) return false;
  const b = { s: m[0].replace(/<\/w:style>$/, '') };
  // body of the style is everything between the opening tag and </w:style>; edit pPr / rPr inside it
  if (look.font !== false) setIn(b, 'w:rPr', `<w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:eastAsia="${FONT}" w:cs="${FONT}"/>`, ORDER_R);
  if (look.size) { setIn(b, 'w:rPr', `<w:sz w:val="${look.size * 2}"/>`, ORDER_R); setIn(b, 'w:rPr', `<w:szCs w:val="${look.size * 2}"/>`, ORDER_R); }
  if (look.bold != null) { setIn(b, 'w:rPr', look.bold ? '<w:b/>' : '<w:b w:val="0"/>', ORDER_R); setIn(b, 'w:rPr', look.bold ? '<w:bCs/>' : '<w:bCs w:val="0"/>', ORDER_R); }
  if (look.italic === false) { setIn(b, 'w:rPr', '<w:i w:val="0"/>', ORDER_R); setIn(b, 'w:rPr', '<w:iCs w:val="0"/>', ORDER_R); }
  if (look.color) setIn(b, 'w:rPr', `<w:color w:val="${look.color}"/>`, ORDER_R);
  if (look.align) setIn(b, 'w:pPr', `<w:jc w:val="${look.align}"/>`, ORDER_P);
  if (look.before != null || look.after != null || look.line != null) {
    const pt = v => String(Math.round(v * 20));
    setIn(b, 'w:pPr', `<w:spacing w:before="${pt(look.before || 0)}" w:after="${pt(look.after || 0)}" w:line="${look.line ? Math.round(look.line * 240) : 240}" w:lineRule="auto"/>`, ORDER_P);
  }
  if (look.first_line_cm != null || look.left_cm != null) setIn(b, 'w:pPr', `<w:ind w:left="${cmTw(look.left_cm || 0)}" w:firstLine="${cmTw(look.first_line_cm || 0)}"/>`, ORDER_P);
  st = st.replace(re, b.s + '</w:style>');
  return true;
}
const looks = STD.styles || {};
let styled = 0;
for (const [names, look] of Object.entries(looks)) for (const n of names.split('|')) if (applyLook(n.trim(), look)) styled++;
log.push(`styles set to the standard: ${styled}`);

// ---------- 3) document defaults: same base spacing / font for every book ----------
if (STD.defaults) {
  const D = STD.defaults;
  st = st.replace(/<w:pPrDefault>[\s\S]*?<\/w:pPrDefault>/, `<w:pPrDefault><w:pPr><w:spacing w:after="${Math.round((D.after || 0) * 20)}" w:line="${Math.round((D.line || 1) * 240)}" w:lineRule="auto"/></w:pPr></w:pPrDefault>`);
  st = st.replace(/(<w:rPrDefault><w:rPr>)([\s\S]*?)(<\/w:rPr><\/w:rPrDefault>)/, (m, a, inner, c) => {
    inner = inner.replace(/<w:rFonts\b[^>]*\/>/, '').replace(/<w:sz w:val="\d+"\/>/, '').replace(/<w:szCs w:val="\d+"\/>/, '');
    return a + `<w:rFonts w:ascii="${FONT}" w:eastAsia="${FONT}" w:hAnsi="${FONT}" w:cs="${FONT}"/>` + inner + `<w:sz w:val="${(D.size || 16) * 2}"/><w:szCs w:val="${(D.size || 16) * 2}"/>` + c;
  });
  log.push('document defaults (font / size / spacing) set');
}

// ---------- 4) force sizes (ผู้ใช้สั่ง "บังคับปรับขนาดตามเงื่อนไข"): hand-typed sizes on runs are removed so every
//      paragraph shows its style's standard size (body 16, chapter 20, TOC 16 …). Cover section, code font and
//      equations are left alone. Tables get their size later (16, dense 14, never below table_min_font_pt).
let forced = 0;
if (STD.force_sizes !== false) {
  const CODE = new RegExp(cfg.code_fonts || 'Consol|Courier|Mono', 'i');
  const bodyStart = doc.indexOf('<w:body>');
  const firstSect = doc.indexOf('<w:sectPr', bodyStart);           // end of the cover section (first section break)
  const coverEnd = firstSect > 0 && doc.indexOf('</w:body>') > firstSect + 50 && doc.lastIndexOf('<w:sectPr') !== firstSect ? doc.indexOf('</w:p>', firstSect) : bodyStart;
  const head = doc.slice(0, coverEnd), tail = doc.slice(coverEnd);
  const fixed = tail.replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/g, (m, inner) => {
    if (CODE.test((inner.match(/<w:rFonts\b[^>]*>/) || [''])[0])) return m;
    const out = inner.replace(/<w:sz w:val="\d+"\/>/, '').replace(/<w:szCs w:val="\d+"\/>/, '');
    if (out !== inner) forced++;
    return out.trim() ? `<w:rPr>${out}</w:rPr>` : '';   // empty rPr is dropped (keeps the XML round-trip exact)
  });
  doc = head + fixed;
}
log.push(`hand-typed text sizes reset to the standard: ${forced}`);

// ---------- 5) one font for every run, every slot (Latin / Thai "complex script" / East Asian) ----------
// Thai text tagged as complex script uses the CS font; without an explicit one Word falls back to the table style /
// theme (Angsana, Cordia -> Thai looks smaller than the English next to it). Code fonts are kept.
let fontRuns = 0;
{
  const CODE = new RegExp(cfg.code_fonts || 'Consol|Courier|Mono', 'i');
  const F = `<w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:eastAsia="${FONT}" w:cs="${FONT}"/>`;
  const bodyStart = doc.indexOf('<w:body>');
  doc = doc.slice(0, bodyStart) + doc.slice(bodyStart).replace(/<w:r(\s[^>]*)?>(<w:rPr>([\s\S]*?)<\/w:rPr>)?/g, (m, a = '', rp, inner) => {
    if (inner != null) {
      const rf = (inner.match(/<w:rFonts\b[^>]*\/>/) || [''])[0];
      if (CODE.test(rf)) return m;
      fontRuns++;
      return `<w:r${a}><w:rPr>${F}${inner.replace(/<w:rFonts\b[^>]*\/>/, '')}</w:rPr>`;
    }
    fontRuns++;
    return `<w:r${a}><w:rPr>${F}</w:rPr>`;
  });
  // table styles / character styles / anything still pointing at theme fonts
  st = st.replace(/<w:rFonts\b[^>]*\/>/g, rf => CODE.test(rf) ? rf : F);
}
log.push(`runs given the book font in every script slot: ${fontRuns}`);

d.set('word/styles.xml', st); d.set('word/document.xml', doc);
saveDocx(d, outF);
console.log('standardize: ' + log.join('; '));
