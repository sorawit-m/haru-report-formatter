// fix-tables.mjs  —  งานตารางที่ต้องแก้ใน XML (ใช้ได้ทั้ง Windows และ Mac)
//  1) ตาราง "หัว" กับ "เนื้อหา" ที่ถูกแยกเป็นตารางซ้อนในเซลล์ -> ดึงแถวขึ้นมารวมเป็นตารางเดียว
//  2) ตารางสองตารางติดกัน (คั่นด้วยย่อหน้าว่าง) จำนวนคอลัมน์เท่ากัน และตารางที่สองไม่มีหัว -> รวมเป็นตารางเดียว
//  3) แถวหัวตาราง: ตั้งให้ทำซ้ำทุกหน้า
//  4) ข้อความในเซลล์แถวข้อมูล: เลิกจัดแบบกระจาย (ตัวอักษรถูกถ่าง) -> ชิดซ้าย
//  5) ฟอนต์ในตารางให้เป็นฟอนต์เดียวกัน (ยกเว้นฟอนต์โค้ด)
// ใช้: node fix-tables.mjs <in.docx> <out.docx> [config.json]
import { loadDocx, saveDocx, parseXml, serialize, els, el, all, attr, setAttr, mk, replaceWith, remove, textOf, ensure } from './lib/docx.mjs';
import fs from 'node:fs';
import * as thaiWrap from './lib/thaiwrap.mjs';
if (process.env.RF_LOG) { const orig = console.log; console.log = (...a) => { orig(...a); try { fs.appendFileSync(process.env.RF_LOG, a.join(' ') + '\n', 'utf8'); } catch {} }; }

const [, , inF, outF, cfgF] = process.argv;
if (!inF || !outF) { console.error('usage: node fix-tables.mjs in.docx out.docx [config.json]'); process.exit(1); }
const cfg = cfgF && fs.existsSync(cfgF) ? JSON.parse(fs.readFileSync(cfgF, 'utf8')) : {};
const FONT = cfg.heading_font || 'TH SarabunPSK';
const CODE = new RegExp(cfg.code_fonts || 'Consol|Courier|Mono', 'i');

const doc = loadDocx(inF);
const xml = doc.get('word/document.xml');
const root = parseXml(xml);
if (serialize(root) !== xml) { console.error('XML round-trip mismatch — stop (file not changed)'); process.exit(2); }

// 0) whole-paragraph rewrites chosen by Claude for cells that still stretch (rewrites.tsv: เดิม<TAB>ใหม่, ผู้ใช้ให้สิทธิ์ปรับคำเต็มที่)
//    only when the paragraph's runs share the same look (so nothing like bold/red is lost); red text is never touched
let rewritten = 0;
if (process.env.RF_REWRITES && fs.existsSync(process.env.RF_REWRITES)) {
  const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const pairs = fs.readFileSync(process.env.RF_REWRITES, 'utf8').replace(/^﻿/, '').split(/\r?\n/).filter(l => l.includes('\t')).map(l => l.split('\t').map(esc));
  const look = r => { const rp = el(r, 'w:rPr'); return rp ? serialize(rp).replace(/<w:cs\/>|<w:lang[^>]*\/>|<w:szCs[^>]*\/>|<w:bCs[^>]*\/>|<w:rFonts[^>]*\/>/g, '') : ''; };
  for (const p of all(root, 'w:p')) {
    const runs = els(p, 'w:r').filter(r => els(r, 'w:t').length); if (!runs.length) continue;
    const s = textOf(p);
    const hit = pairs.find(([o]) => s.includes(o)); if (!hit) continue;
    if (new Set(runs.map(look)).size > 1 || /FF0000|C00000/i.test(serialize(p))) continue;
    if (els(p, 'w:r').some(r => !els(r, 'w:t').length && els(r).some(c => c.name !== 'w:rPr'))) continue;   // tabs/breaks/fields
    const ns = s.replace(hit[0], hit[1]);
    const t0 = els(runs[0], 'w:t')[0]; t0.children = [{ type: 'text', text: ns }]; if (!t0.attrs.some(a => a[0] === 'xml:space')) t0.attrs.push(['xml:space', 'preserve']);
    for (const t of els(runs[0], 'w:t').slice(1)) remove(t);
    for (const r of runs.slice(1)) remove(r);
    rewritten++; console.log(`  rewrite: ${s.slice(0, 60)}  ->  ${ns.slice(0, 70)}`);
  }
}
console.log(`cells rewritten (rewrites.tsv): ${rewritten}`);

const isEmptyPara = p => p.name === 'w:p' && !textOf(p).trim() && !all(p, 'w:drawing').length && !all(p, 'w:pict').length && !all(p, 'w:sectPr').length && !all(p, 'w:br').some(b => attr(b, 'w:type') === 'page');
const gridCount = tbl => { const g = el(tbl, 'w:tblGrid'); return g ? els(g, 'w:gridCol').length : 0; };
const rowCellCount = tr => els(tr, 'w:tc').reduce((s, tc) => { const tcPr = el(tc, 'w:tcPr'); const gs = tcPr && el(tcPr, 'w:gridSpan'); return s + (gs ? parseInt(attr(gs, 'w:val')) || 1 : 1); }, 0);
const firstCellText = tbl => { const tr = el(tbl, 'w:tr'); const tc = tr && el(tr, 'w:tc'); return tc ? textOf(tc).trim() : ''; };
const looksLikeHeader = (tbl) => { const tr = el(tbl, 'w:tr'); if (!tr) return false; const shd = all(tr, 'w:shd').some(s => { const f = (attr(s, 'w:fill') || '').toUpperCase(); return f && f !== 'AUTO' && f !== 'FFFFFF'; }); return shd; };

let unnested = 0, merged = 0, hdr = 0, aligned = 0, fonted = 0;

// which paragraph styles are "distributed / justified" (the style's own jc, following basedOn)
const stylesXml = doc.get('word/styles.xml') || '';
const styleJc = {}, styleBase = {}; let defaultStyleId = null;
for (const m of stylesXml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
  const id = (m[1].match(/w:styleId="([^"]*)"/) || [])[1]; if (!id) continue;
  if (/w:type="paragraph"/.test(m[1]) && /w:default="1"/.test(m[1])) defaultStyleId = id;
  const jc = (m[2].match(/<w:jc w:val="([^"]*)"/) || [])[1]; if (jc) styleJc[id] = jc;
  const b = (m[2].match(/<w:basedOn w:val="([^"]*)"/) || [])[1]; if (b) styleBase[id] = b;
}
function resolveJc(id) { for (let k = 0; id && k < 10; k++, id = styleBase[id]) if (styleJc[id]) return styleJc[id]; return null; }
function defaultDistributed(pPr) {
  const ps = pPr && el(pPr, 'w:pStyle'); const id = ps ? attr(ps, 'w:val') : defaultStyleId;
  return /^(thaiDistribute|distribute|both)$/.test(resolveJc(id) || '');
}

// ---- "กระจายแบบไทย" (Thai Distributed): เนื้อหาทั่วไป + ข้อความในเซลล์แถวข้อมูล ----
// ใช้ได้เฉพาะย่อหน้าที่: มีอักษรไทย, ไม่มี Shift+Enter (บรรทัดก่อนตัดจะถูกถ่าง), ไม่มีรูป, ไม่ใช่ฟอนต์โค้ด,
// ไม่ได้ตั้งกึ่งกลาง/ชิดขวาไว้, ไม่ใช่หัวข้อ/สารบัญ/คำบรรยาย/ชื่อเรื่อง
const BODY_THAI = (cfg.body_align || 'thaiDistribute') === 'thaiDistribute';
const CELL_THAI = (cfg.table_cell_align || 'thaiDistribute') === 'thaiDistribute';
const styleName = {}, styleOutline = {};
for (const m of stylesXml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
  const id = (m[1].match(/w:styleId="([^"]*)"/) || [])[1]; if (!id) continue;
  styleName[id] = ((m[2].match(/<w:name w:val="([^"]*)"/) || [])[1] || id).toLowerCase();
  if (/<w:outlineLvl /.test(m[2])) styleOutline[id] = true;
}
const SPECIAL_STYLE = /heading|toc|caption|title|subtitle|table of figures|header|footer|หัวเรื่อง|ชื่อเรื่อง|คำอธิบายภาพ|สารบัญ|code|โค้ด/;
function isSpecialStyle(id) { for (let k = 0; id && k < 10; k++, id = styleBase[id]) { if (styleOutline[id] || SPECIAL_STYLE.test(styleName[id] || id.toLowerCase())) return true; } return false; }
// กระจายไทยแล้วตัวอักษรถ่าง ("แ จ้ ง Search quota", "ท ะ เ บี ย น") เมื่อบรรทัดมีจุดตัดคำน้อย:
// คอลัมน์แคบ หรือข้อความมีภาษาอังกฤษ/ตัวเลขมาก -> ไม่ใช้กระจาย (ชิดซ้ายแทน)
const MIN_CELL_W = cfg.thai_distribute_min_cell_twips || 2300;      // ~4 ซม.
const MAX_LATIN_CELL = cfg.thai_distribute_max_latin_cell ?? 0.15;  // สัดส่วนอักษรอังกฤษ/ตัวเลข
const MAX_LATIN_BODY = cfg.thai_distribute_max_latin_body ?? 0.3;
function latinShare(t) { const th = (t.match(/[฀-๿]/g) || []).length, la = (t.match(/[A-Za-z0-9]/g) || []).length; return th + la ? la / (th + la) : 1; }
function cellWidth(p) { let tc = p.parent; while (tc && tc.name !== 'w:tc') tc = tc.parent; const w = tc && el(el(tc, 'w:tcPr') || { children: [] }, 'w:tcW'); return w && attr(w, 'w:type') === 'dxa' ? +attr(w, 'w:w') : Infinity; }
// decide per cell by simulating Word's line breaks with the real TH SarabunPSK widths (lib/thaiwrap.mjs):
//   every justified line nearly full -> Thai Distributed;  a line would be left with a big gap -> FALLBACK
//   (default "both" = justify: spreads only the spaces between words, never the letters, edges stay flush)
const GAP_EM = cfg.thai_distribute_max_gap_em ?? 0.065;
// ผู้ใช้เลือก (25-11-2026): ทุกเซลล์ที่มีภาษาไทยใช้ "กระจายแบบไทย" แม้บางบรรทัดตัวอักษรจะถ่างเล็กน้อย
// ("เต็มบรรทัด" ทำให้ช่องว่างระหว่างคำห่างโหว่ / "ชิดซ้าย" ขอบขวาไม่ตรง — ผู้ใช้ไม่เอาทั้งสองแบบ)
const FALLBACK = cfg.stretch_fallback_align || 'thaiDistribute';
let fallbackCells = 0;
function cellFontPt(p) { const r = all(p, 'w:r').find(x => textOf(x).trim()); const sz = r ? +effective(r, 'sz') : 32; return (sz || 32) / 2; }
function cellPadTw(p) { let t = p.parent; while (t && t.name !== 'w:tbl') t = t.parent; const m = t && el(el(t, 'w:tblPr') || { children: [] }, 'w:tblCellMar'); const l = m && el(m, 'w:left'), r = m && el(m, 'w:right'); return (l ? +attr(l, 'w:w') : 108) + (r ? +attr(r, 'w:w') : 108); }
function goodForThaiDistribute(p, inCell) {
  const t = textOf(p);
  if (inCell) {
    const w = cellWidth(p);
    if (thaiWrap.available && Number.isFinite(w)) return !thaiWrap.stretches(t, (w - cellPadTw(p)) / 20 / cellFontPt(p), GAP_EM) ? 'thaiDistribute' : FALLBACK;
    return w >= MIN_CELL_W && latinShare(t) <= MAX_LATIN_CELL ? 'thaiDistribute' : FALLBACK;
  }
  return latinShare(t) <= MAX_LATIN_BODY ? 'thaiDistribute' : FALLBACK;
}
function canThaiDistribute(p) {
  const pPr = el(p, 'w:pPr'); const ps = pPr && el(pPr, 'w:pStyle');
  if (isSpecialStyle(ps ? attr(ps, 'w:val') : defaultStyleId)) return false;
  if (pPr && el(pPr, 'w:outlineLvl')) return false;
  const jc = pPr && el(pPr, 'w:jc'); const cur = jc ? attr(jc, 'w:val') : resolveJc(ps ? attr(ps, 'w:val') : defaultStyleId);
  if (/^(center|right|end)$/.test(cur || '')) return false;
  const t = textOf(p); if (!/[฀-๿]/.test(t)) return false;
  if (all(p, 'w:br').some(b => !attr(b, 'w:type') || attr(b, 'w:type') === 'textWrapping')) return false;
  // lines laid out with tabs / runs of spaces ("เรื่อง <tab><tab> หน้า" on TOC pages) must not be distributed
  if (all(p, 'w:tab').length || / {4,}/.test(t)) return false;
  if (all(p, 'w:drawing').length || all(p, 'w:pict').length || all(p, 'w:fldSimple').length) return false;
  if (all(p, 'w:instrText').some(i => /TOC|PAGEREF/.test(textOf(i)))) return false;
  if (all(p, 'w:rFonts').some(f => CODE.test(attr(f, 'w:ascii') || attr(f, 'w:hAnsi') || ''))) return false;
  return true;
}
function setJc(p, val) {
  const pp = el(p, 'w:pPr') || ensure(p, 'w:pPr'); let jc = el(pp, 'w:jc');
  if (jc) { setAttr(jc, 'w:val', val); return; }
  const after = ['w:textDirection', 'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange'];
  jc = mk('w:jc', [['w:val', val]]); jc.parent = pp;
  const i = pp.children.findIndex(c => c.type === 'el' && after.includes(c.name));
  if (i < 0) pp.children.push(jc); else pp.children.splice(i, 0, jc);
}
let toThai = 0, bodyThai = 0, brJoined = 0;

// Shift+Enter กลางประโยค (ข้อความ "ตก" บรรทัดก่อนเต็มช่อง) -> เปลี่ยนเป็นช่องว่าง ให้ข้อความไหลเต็มบรรทัด
// คงไว้เมื่อ: บรรทัดถัดไปขึ้นต้นด้วยเลขข้อ/สัญลักษณ์รายการ, บรรทัดก่อนจบด้วย ":", ไม่มีอักษรไทยฝั่งใดฝั่งหนึ่ง (โค้ด/ที่อยู่), ฟอนต์โค้ด
// code / JSON / HTTP samples are laid out line by line on purpose — never join, merge or re-space them
// (API book: "ตัวอย่าง Response ⏎ HTTP/1.1 200 OK ⏎ { ⏎ "code": 200, …" was squashed into one line once)
const LOOKS_CODE = /^\s*[{}\[\]]|^\s*"[^"]*"\s*:|^\s*(HTTP\/\d|GET |POST |PUT |PATCH |DELETE |curl |SELECT |INSERT |UPDATE |<\?xml|<[a-z]+[ >])|[{\[,;]\s*$|^\s*\/\/|=>|^\s*(const|let|var|function|import|return)\b/;
const looksCode = t => t.split(/\r?\n/).some(l => LOOKS_CODE.test(l));
function cellHasCode(node) { let tc = node; while (tc && tc.name !== 'w:tc') tc = tc.parent; if (!tc) return false; return els(tc, 'w:p').some(p => looksCode(textOf(p))); }
const LIST_START =/^\s*(\d+(\.\d+)*[.)]|[-–•*▪●○◦]|\(?[a-zA-Zก-ฮ0-9]{1,2}\)|[ก-ฮ]\.|ข้อ\s*\d)/;
function joinSoftBreaks(p) {
  if (all(p, 'w:rFonts').some(f => CODE.test(attr(f, 'w:ascii') || attr(f, 'w:hAnsi') || ''))) return;
  if (cellHasCode(p)) return;
  // headings / captions / TOC lines: a Shift+Enter there is on purpose ("สถาปัตยกรรมข้อมูล⏎(DATA ARCHITECTURE)") -> keep
  { const pp = el(p, 'w:pPr'); const ps = pp && el(pp, 'w:pStyle'); if ((pp && el(pp, 'w:outlineLvl')) || isSpecialStyle(ps ? attr(ps, 'w:val') : defaultStyleId)) return; }
  const seq = [];
  (function walk(n) { for (const c of n.children || []) { if (c.type !== 'el') continue; if (c.name === 'w:t') seq.push({ t: c.children.map(x => x.text || '').join('') }); else if (c.name === 'w:br') { const ty = attr(c, 'w:type'); seq.push(!ty || ty === 'textWrapping' ? { br: c } : { t: '\n' }); } else if (c.name === 'w:tab') seq.push({ t: '\t' }); else walk(c); } })(p);
  for (let k = 0; k < seq.length; k++) {
    if (!seq[k].br) continue;
    let before = ''; for (let m = k - 1; m >= 0 && !seq[m].br; m--) before = seq[m].t + before;
    let after = ''; for (let m = k + 1; m < seq.length && !seq[m].br; m++) after += seq[m].t;
    const b = before.replace(/\s+$/, ''), a = after.replace(/^\s+/, '');
    if (!b || !a) continue;
    if (LIST_START.test(a) || /[:：]$/.test(b)) continue;
    if (!/[฀-๿]/.test(b + a)) continue;
    const br = seq[k].br;
    if (/\s$/.test(before) || /^\s/.test(after)) remove(br);
    else { const t = mk('w:t', [['xml:space', 'preserve']], [{ type: 'text', text: ' ' }]); replaceWith(br, [t]); }
    seq[k] = { t: ' ' }; brJoined++;
  }
  collapseSpaces(p);
}
// after joining: no double spaces ("รูป  ผู้รับบริการ  (JPG") — inside one text node, and across neighbouring text nodes
function collapseSpaces(p) {
  const pPr = el(p, 'w:pPr'); const ps = pPr && el(pPr, 'w:pStyle');
  if (isSpecialStyle(ps ? attr(ps, 'w:val') : defaultStyleId)) return;          // TOC header "เรื่อง      หน้า" etc.
  if (!/[฀-๿]/.test(textOf(p)) || / {5,}/.test(textOf(p))) return;     // spaces used for alignment -> leave
  if (cellHasCode(p) || looksCode(textOf(p))) return;                             // indentation in code samples
  const ts = all(p, 'w:t'); let prevEndsSpace = false;
  for (const t of ts) {
    let s = t.children.map(c => c.text || '').join(''); const orig = s;
    s = s.replace(/ {2,4}/g, ' ');
    if (prevEndsSpace) s = s.replace(/^ {1,4}/, '');
    if (s !== orig) { t.children = [{ type: 'text', text: s }]; if (!t.attrs.some(a => a[0] === 'xml:space')) t.attrs.push(['xml:space', 'preserve']); }
    if (s.length) prevEndsSpace = / $/.test(s);
  }
}

// ย่อหน้าที่ถูกตัดกลางประโยคในเซลล์เดียวกัน -> รวมเป็นย่อหน้าเดียว (เหมือนกด Cmd+C แล้ว Cmd+Shift+V บน Mac)
// คงแยกไว้: ย่อหน้าถัดไปขึ้นต้นด้วยเลขข้อ/bullet, ย่อหน้าก่อนจบด้วย ":", ย่อหน้าที่เป็นรายการ (numPr), มีรูป, ว่าง,
//           ตั้งกึ่งกลาง/ชิดขวา, Style ต่างกัน, หัวข้อ/คำบรรยาย, ไม่มีอักษรไทย, ฟอนต์โค้ด
let parasMerged = 0;
function plainPara(p) {
  const pPr = el(p, 'w:pPr');
  if (pPr && (el(pPr, 'w:numPr') || el(pPr, 'w:sectPr') || el(pPr, 'w:outlineLvl'))) return false;
  const ps = pPr && el(pPr, 'w:pStyle'); const sid = ps ? attr(ps, 'w:val') : defaultStyleId;
  if (isSpecialStyle(sid)) return false;
  const jc = pPr && el(pPr, 'w:jc'); const cur = jc ? attr(jc, 'w:val') : resolveJc(sid);
  if (/^(center|right|end)$/.test(cur || '')) return false;
  if (all(p, 'w:drawing').length || all(p, 'w:pict').length || all(p, 'w:object').length || all(p, 'w:instrText').length || all(p, 'w:fldSimple').length) return false;
  if (all(p, 'w:rFonts').some(f => CODE.test(attr(f, 'w:ascii') || attr(f, 'w:hAnsi') || ''))) return false;
  return !!textOf(p).trim();
}
function mergeCellParagraphs(tc) {
  let ps = els(tc, 'w:p');
  if (ps.some(p => looksCode(textOf(p)) || all(p, 'w:br').some(b => looksCode(textOf(p))))) return;   // code / JSON sample cell
  for (let k = 1; k < ps.length; k++) {
    const a = ps[k - 1], b = ps[k];
    if (!plainPara(a) || !plainPara(b)) continue;
    const sa = (el(el(a, 'w:pPr') || { children: [] }, 'w:pStyle') || null), sb = (el(el(b, 'w:pPr') || { children: [] }, 'w:pStyle') || null);
    if ((sa && attr(sa, 'w:val')) !== (sb && attr(sb, 'w:val'))) continue;
    const ta = textOf(a), tb = textOf(b);
    if (LIST_START.test(tb) || LIST_START.test(ta) && /^\s*\d/.test(tb) || /[:：]\s*$/.test(ta)) continue;
    if (!/[฀-๿]/.test(ta + tb)) continue;
    // move b's content (everything except pPr) to the end of a; Thai + Thai is one phrase cut by Enter -> no space
    // (same as Word for Mac paste: "ลายม่านตา" + "ในส่วน" -> "ลายม่านตาในส่วน"); otherwise one space
    const thaiJoin = /[฀-๿]$/.test(ta) && /^[฀-๿]/.test(tb);
    if (!thaiJoin && !/\s$/.test(ta) && !/^\s/.test(tb)) { const sp = mk('w:r', [], [mk('w:t', [['xml:space', 'preserve']], [{ type: 'text', text: ' ' }])]); sp.parent = a; a.children.push(sp); }
    for (const c of [...b.children]) { if (c.type === 'el' && c.name === 'w:pPr') continue; b.children.splice(b.children.indexOf(c), 1); c.parent = a; a.children.push(c); }
    remove(b); parasMerged++;
    ps = els(tc, 'w:p'); k--;
  }
}

// ติดป้าย "ข้อความไทย" (<w:cs/>) ให้ข้อความไทยทุกที่ แยกคำอังกฤษ/ตัวเลขเป็น run ของตัวเอง
// (สิ่งที่ Word บน Mac ทำตอน Cmd+Shift+V — ทำให้ตัดคำไทยตรงรอยคำ ไม่ขาดกลางคำ; Windows แสดงผลเหมือนเดิม)
let csTagged = 0, runsSplit = 0;
const RPR_AFTER_CS = ['w:em', 'w:lang', 'w:eastAsianLayout', 'w:specVanish', 'w:oMath', 'w:rPrChange'];
// effective size / bold of a run (run -> character style -> paragraph style chain -> defaults) so that the Thai
// ("complex script") size and bold match what the run shows today
const stRun = {};
for (const m of stylesXml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
  const id = (m[1].match(/w:styleId="([^"]*)"/) || [])[1]; if (!id) continue;
  const rp = (m[2].match(/<w:rPr>([\s\S]*?)<\/w:rPr>/) || [])[1] || '';
  stRun[id] = { sz: (rp.match(/<w:sz w:val="(\d+)"/) || [])[1], szCs: (rp.match(/<w:szCs w:val="(\d+)"/) || [])[1], b: /<w:b\/>|<w:b w:val="(1|true|on)"/.test(rp) ? true : /<w:b w:val="(0|false|off)"/.test(rp) ? false : undefined, bCs: /<w:bCs\/>|<w:bCs w:val="(1|true|on)"/.test(rp) ? true : /<w:bCs w:val="(0|false|off)"/.test(rp) ? false : undefined };
}
const defRpr = (stylesXml.match(/<w:rPrDefault>[\s\S]*?<\/w:rPrDefault>/) || [''])[0];
const DEF_SZ = (defRpr.match(/<w:sz w:val="(\d+)"/) || [])[1] || '20';
function styleProp(id, k) { for (let n = 0; id && n < 12; n++, id = styleBase[id]) if (stRun[id] && stRun[id][k] !== undefined) return stRun[id][k]; return undefined; }
function effective(r, k) {
  const rPr = el(r, 'w:rPr'); const p = r.parent && (r.parent.name === 'w:p' ? r.parent : r.parent.parent);
  const own = rPr && el(rPr, k === 'sz' ? 'w:sz' : 'w:b');
  if (own) return k === 'sz' ? attr(own, 'w:val') : !/^(0|false|off)$/.test(attr(own, 'w:val') || '1');
  const rs = rPr && el(rPr, 'w:rStyle'); if (rs) { const v = styleProp(attr(rs, 'w:val'), k); if (v !== undefined) return v; }
  const pPr = p && el(p, 'w:pPr'); const ps = pPr && el(pPr, 'w:pStyle');
  const v = styleProp(ps ? attr(ps, 'w:val') : defaultStyleId, k); if (v !== undefined) return v;
  return k === 'sz' ? DEF_SZ : false;
}
function syncCsFormatting(r, rPr) {
  const sz = effective(r, 'sz'); const b = effective(r, 'b');
  const put = (name, val) => {   // keep schema order: b,bCs,i,iCs ... sz,szCs ...
    if (el(rPr, name)) { if (val !== null) setAttr(el(rPr, name), 'w:val', val); return; }
    const n = mk(name, val === null ? [] : [['w:val', val]]); n.parent = rPr;
    const before = name === 'w:bCs' ? 'w:b' : 'w:sz';
    const bi = rPr.children.findIndex(c => c.type === 'el' && c.name === before);
    if (bi >= 0) { rPr.children.splice(bi + 1, 0, n); return; }
    const order = ['w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps', 'w:strike', 'w:dstrike', 'w:outline', 'w:shadow', 'w:emboss', 'w:imprint', 'w:noProof', 'w:snapToGrid', 'w:vanish', 'w:webHidden', 'w:color', 'w:spacing', 'w:w', 'w:kern', 'w:position', 'w:sz', 'w:szCs'];
    const my = order.indexOf(name);
    const i = rPr.children.findIndex(c => c.type === 'el' && (order.indexOf(c.name) > my || order.indexOf(c.name) < 0));
    if (i < 0) rPr.children.push(n); else rPr.children.splice(i, 0, n);
  };
  if (sz) put('w:szCs', String(sz));
  if (b === true && !el(rPr, 'w:bCs')) put('w:bCs', null);
}
function addCs(rPr, r) {
  if (r) syncCsFormatting(r, rPr);
  if (el(rPr, 'w:cs')) return;
  const n = mk('w:cs'); n.parent = rPr;
  const i = rPr.children.findIndex(c => c.type === 'el' && RPR_AFTER_CS.includes(c.name));
  if (i < 0) rPr.children.push(n); else rPr.children.splice(i, 0, n);
}
const cloneEl = n => n.type === 'el' ? (() => { const e = mk(n.name, n.attrs.map(a => [...a]), n.children.map(cloneEl)); return e; })() : { ...n };
function tagThaiRuns(scope) {
  for (const r of all(scope, 'w:r')) {
    if (!r.parent) continue;
    const rPr = el(r, 'w:rPr');
    const rf = rPr && el(rPr, 'w:rFonts'); if (rf && CODE.test(attr(rf, 'w:ascii') || attr(rf, 'w:hAnsi') || '')) continue;
    const kids = els(r).filter(c => c.name !== 'w:rPr');
    const ts = kids.filter(c => c.name === 'w:t'); if (!ts.length) continue;
    const text = ts.map(t => textOf({ children: [t] })).join('');
    if (!/[฀-๿]/.test(text)) continue;
    const hasLatin = /[A-Za-z0-9]/.test(text);
    if (!hasLatin || kids.length !== 1) { addCs(rPr || ensure(r, 'w:rPr'), r); csTagged++; continue; }
    // one w:t with Thai + Latin -> split into segments (neutral chars stay with the current segment)
    const segs = []; let cur = null;
    for (const ch of text.match(/&[#A-Za-z0-9]+;|[\s\S]/gu)) {   // raw XML text: keep entities (&amp;) whole
      const cls = ch.length > 1 && ch[0] === '&' ? null : /[฀-๿]/.test(ch) ? 'T' : /[A-Za-z0-9]/.test(ch) ? 'L' : null;
      if (!cur) cur = { cls: cls || 'T', s: '' };
      else if (cls && cls !== cur.cls) { segs.push(cur); cur = { cls, s: '' }; }
      cur.s += ch;
    }
    segs.push(cur);
    const news = segs.map(sg => {
      const nr = mk('w:r', r.attrs.map(a => [...a]), []);
      const rp = rPr ? cloneEl(rPr) : null; if (rp) { rp.parent = nr; nr.children.push(rp); }
      if (sg.cls === 'T') { const rr = rp || (() => { const x = mk('w:rPr'); x.parent = nr; nr.children.unshift(x); return x; })(); addCs(rr, r); }
      const t = mk('w:t', [['xml:space', 'preserve']], [{ type: 'text', text: sg.s }]); t.parent = nr; nr.children.push(t);
      return nr;
    });
    replaceWith(r, news); runsSplit++; csTagged++;
  }
}

// 1) nested table inside a single merged cell -> lift its rows into the outer table
for (const outer of all(root, 'w:tbl')) {
  if (!outer.parent) continue;
  for (const tr of els(outer, 'w:tr')) {
    const tcs = els(tr, 'w:tc'); if (tcs.length !== 1) continue;
    const inner = els(tcs[0], 'w:tbl'); if (inner.length !== 1) continue;
    const others = els(tcs[0]).filter(c => c.name !== 'w:tcPr' && c !== inner[0]);
    if (!others.every(isEmptyPara)) continue;
    if (gridCount(inner[0]) && gridCount(outer) && gridCount(inner[0]) !== gridCount(outer)) continue;
    const rows = els(inner[0], 'w:tr');
    replaceWith(tr, rows);
    unnested++;
  }
  // drop rows that became completely empty
  for (const tr of els(outer, 'w:tr')) { if (!textOf(tr).trim() && !all(tr, 'w:drawing').length && !all(tr, 'w:tbl').length && els(outer, 'w:tr').length > 1 && tr !== el(outer, 'w:tr')) { remove(tr); } }
}

// 2) adjacent tables (only empty paragraphs between), same columns, second without its own header -> one table
const body = el(el(root, 'w:document'), 'w:body');
let kids = els(body);
for (let i = 0; i < kids.length; i++) {
  const a = kids[i]; if (a.name !== 'w:tbl' || !a.parent) continue;
  let j = i + 1; const gap = [];
  // gap may hold empty paragraphs and page-break-only paragraphs (the table simply continues on the next page)
  const emptyOrBreak = p => p.name === 'w:p' && !textOf(p).trim() && !all(p, 'w:drawing').length && !all(p, 'w:sectPr').length;
  while (j < kids.length && emptyOrBreak(kids[j])) { gap.push(kids[j]); j++; }
  const b = kids[j]; if (!b || b.name !== 'w:tbl') continue;
  // same number of columns: compare the widest row (grids may differ by a hair-thin extra column)
  const mainCells = t => { const c = {}; els(t, 'w:tr').forEach(tr => { const n = rowCellCount(tr); c[n] = (c[n] || 0) + 1; }); return +Object.entries(c).sort((x, y) => y[1] - x[1])[0][0]; };
  if (mainCells(a) !== mainCells(b) || looksLikeHeader(b) || !looksLikeHeader(a)) continue;
  if (els(a, 'w:tr').length > 1 && firstCellText(b) === firstCellText(a)) continue;
  gap.forEach(remove);
  for (const tr of els(b, 'w:tr')) { remove(tr); tr.parent = a; a.children.push(tr); }
  remove(b); merged++;
  kids = els(body); i = Math.max(-1, kids.indexOf(a) - 1);
}

// 3–5) header repeat, left-align data cells, one font inside tables
for (const tbl of all(root, 'w:tbl')) {
  const trs = els(tbl, 'w:tr'); if (!trs.length) continue;
  const head = looksLikeHeader(tbl) && trs.length > 1;
  if (head) { const trPr = ensure(trs[0], 'w:trPr'); if (!el(trPr, 'w:tblHeader')) { trPr.children.push(mk('w:tblHeader')); el(trPr, 'w:tblHeader').parent = trPr; hdr++; } }
  trs.forEach((tr, idx) => {
    if (head && idx === 0) return;
    if (cfg.merge_broken_paragraphs_in_cells !== false) for (const tc of els(tr, 'w:tc')) mergeCellParagraphs(tc);
    for (const p of all(tr, 'w:p', [], 'w:tbl')) {
      const pPr = el(p, 'w:pPr'); const jc = pPr && el(pPr, 'w:jc');
      if (cfg.join_soft_breaks_in_cells !== false) joinSoftBreaks(p);
      if (CELL_THAI && canThaiDistribute(p)) {
        const how = goodForThaiDistribute(p, true);
        setJc(p, how); if (how === 'thaiDistribute') toThai++; else fallbackCells++;
        continue;
      }
      const isCode = all(p, 'w:rFonts').some(f => CODE.test(attr(f, 'w:ascii') || attr(f, 'w:hAnsi') || ''));
      // no ragged (left) cells: text without Thai / with Shift+Enter -> justify (spaces only); code stays left
      { const ps0 = pPr && el(pPr, 'w:pStyle'); const sid0 = ps0 ? attr(ps0, 'w:val') : defaultStyleId;
        // Thai text that keeps a Shift+Enter (list lines): justify, because Word does not stretch a line ending in
        // Shift+Enter under "both" (doNotExpandShiftReturn) but may under Thai Distributed
        if (!isCode && !isSpecialStyle(sid0) && /[฀-๿]/.test(textOf(p)) && (!jc || /^(left|start)$/.test(attr(jc, 'w:val') || ''))) { setJc(p, all(p, 'w:br').length ? 'both' : FALLBACK); fallbackCells++; continue; } }
      if (jc) { if (/^(thaiDistribute|distribute|both)$/.test(attr(jc, 'w:val') || '')) { setAttr(jc, 'w:val', isCode ? 'left' : 'both'); aligned++; } }
      else if (defaultDistributed(pPr)) {
        // alignment inherited from a "distributed" style -> set left explicitly (schema order: jc goes before these)
        const pp = pPr || ensure(p, 'w:pPr');
        const after = ['w:textDirection', 'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange'];
        const n = mk('w:jc', [['w:val', isCode ? 'left' : 'both']]); n.parent = pp;
        const idx = pp.children.findIndex(c => c.type === 'el' && after.includes(c.name));
        if (idx < 0) pp.children.push(n); else pp.children.splice(idx, 0, n);
        aligned++;
      }
    }
  });
  if (cfg.table_one_font !== false) {
    for (const r of all(tbl, 'w:r')) {
      const rPr = el(r, 'w:rPr'); const rf = rPr && el(rPr, 'w:rFonts');
      const cur = rf ? (attr(rf, 'w:ascii') || attr(rf, 'w:cs') || '') : '';
      if (CODE.test(cur) || /Symbol|Math|Emoji|Wingdings/i.test(cur)) continue;
      const rp = rPr || ensure(r, 'w:rPr');
      const f = el(rp, 'w:rFonts') || (() => { const n = mk('w:rFonts'); n.parent = rp; rp.children.unshift(n); return n; })();
      f.attrs = f.attrs.filter(([k]) => !/Theme$/.test(k) && k !== 'w:hint');
      for (const k of ['w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia']) setAttr(f, k, FONT);
      fonted++;
    }
  }
}

// 5a) label columns (every data cell a short label: "ส่วนของระบบ", "รหัส", "หัวข้อ") -> centred.
//     ผู้ใช้เลือก (25-11-2026): กระจายไทยในคำสั้น 2 บรรทัด ("ทะเบียน / ผู้รับบริการ") บรรทัดแรกถ่าง -> จัดกึ่งกลางเฉพาะคอลัมน์ป้าย
//     Only tables with a header row; key-value tables (shaded first column: Use Case ID | UC-007) are left alone.
let labelCols = 0;
if (cfg.center_label_columns !== false) {
  const shaded = tc => all(tc, 'w:shd').some(s => { const f = (attr(s, 'w:fill') || '').toUpperCase(); return f && f !== 'AUTO' && f !== 'FFFFFF'; });
  const cand = [];   // tables that qualify structurally
  for (const tbl of all(root, 'w:tbl')) {
    const trs = els(tbl, 'w:tr'); if (trs.length < 2 || !looksLikeHeader(tbl)) continue;
    const n = els(trs[0], 'w:tc').length;
    // rows with one cell spanning the whole table (section / picture rows) are not data rows
    const data = trs.slice(1).filter(tr => els(tr, 'w:tc').length === n);
    if (data.length < Math.max(1, (trs.length - 1) * 0.5)) continue;
    if (data.every(tr => shaded(els(tr, 'w:tc')[0]))) continue;           // key-value table
    cand.push({ tbl, data, n, key: n + '#' + els(trs[0], 'w:tc').map(tc => textOf(tc).trim()).join('|') });
  }
  // decide per header kind over ALL tables of that kind (tables with the same header end up alike)
  const kinds = new Map(); for (const c of cand) { if (!kinds.has(c.key)) kinds.set(c.key, []); kinds.get(c.key).push(c); }
  for (const list of kinds.values()) {
    const n = list[0].n;
    for (let j = 0; j < n; j++) {
      const tcs = list.flatMap(c => c.data.map(tr => els(tr, 'w:tc')[j]));
      const texts = tcs.map(tc => textOf(tc).replace(/\s+/g, ' ').trim()).filter(Boolean);
      if (texts.length < 2) continue;
      const short = texts.filter(t => [...t].length <= 35).length;
      if (short / texts.length < 0.8 || Math.max(...texts.map(t => [...t].length)) > 50) continue;
      if (tcs.some(tc => all(tc, 'w:drawing').length || all(tc, 'w:rFonts').some(f => CODE.test(attr(f, 'w:ascii') || '')))) continue;
      for (const tc of tcs) for (const p of els(tc, 'w:p')) setJc(p, 'center');
      labelCols++;
    }
  }
}
console.log(`short-label columns centred: ${labelCols}`);

// 5b) column widths that make Thai Distributed stretch letters ("ท ะ เ บี ย น" in a narrow column):
//     move width from roomy columns to cramped ones. Total table width never changes, no column is made narrower than
//     its longest word, and ALL tables with the same header get the SAME widths (ผู้ใช้สั่ง: ขยายต้องขยายทุกตาราง).
let colTablesChanged = 0, colGroups = 0;
if (thaiWrap.available && cfg.balance_column_widths !== false) {
  const simple = tbl => !all(tbl, 'w:vMerge').length && !all(tbl, 'w:gridSpan').some(g => +attr(g, 'w:val') > 1) && !all(tbl, 'w:tbl').length;
  const groups = new Map();
  for (const tbl of all(root, 'w:tbl')) {
    const trs = els(tbl, 'w:tr'); if (trs.length < 2 || !simple(tbl)) continue;
    const g = el(tbl, 'w:tblGrid'); const cols = g ? els(g, 'w:gridCol') : []; if (cols.length < 2) continue;
    if (!trs.every(tr => els(tr, 'w:tc').length === cols.length)) continue;
    const head = els(trs[0], 'w:tc').map(tc => textOf(tc).trim()).join('|');
    const key = cols.length + '#' + (looksLikeHeader(tbl) ? head : 'nohead:' + Math.random());
    if (!groups.has(key)) groups.set(key, []); groups.get(key).push(tbl);
  }
  for (const tbls of groups.values()) {
    const first = tbls[0]; const n = els(el(first, 'w:tblGrid'), 'w:gridCol').length;
    // start from the widths most tables in the group use (tables of one kind sometimes differ a little -> they all end up identical)
    const widthsOf = t => els(el(t, 'w:tblGrid'), 'w:gridCol').map(c => +attr(c, 'w:w'));
    const tally = new Map(); for (const t of tbls) { const k = widthsOf(t).join(','); tally.set(k, (tally.get(k) || 0) + 1); }
    const W = [...tally.entries()].sort((a, b) => b[1] - a[1])[0][0].split(',').map(Number);
    if (W.some(x => !x)) continue;
    const total = W.reduce((a, b) => a + b, 0);
    if (!tbls.every(t => Math.abs(widthsOf(t).reduce((a, b) => a + b, 0) - total) < 60)) continue;   // same overall width only
    const pad = cellPadTw(all(first, 'w:p')[0] || first);
    const cells = Array.from({ length: n }, () => []);   // [{text, pt}] per column (data rows)
    for (const t of tbls) els(t, 'w:tr').forEach((tr, ri) => { if (ri === 0 && looksLikeHeader(t)) return; els(tr, 'w:tc').forEach((tc, j) => { for (const p of els(tc, 'w:p')) { const s = textOf(p); if (/[฀-๿]/.test(s) && s.trim().length > 3 && !all(p, 'w:br').length) cells[j].push({ s, pt: cellFontPt(p) }); } }); });
    // no column narrower than its longest word in ANY row (header, codes like "BR-CLN-01", English) or than it is now if it already wraps a word
    const allText = Array.from({ length: n }, () => []);
    for (const t of tbls) for (const tr of els(t, 'w:tr')) els(tr, 'w:tc').forEach((tc, j) => { for (const p of els(tc, 'w:p')) { const s = textOf(p); if (s.trim()) allText[j].push({ s, pt: cellFontPt(p) }); } });
    const minW = allText.map((cs, j) => { const need = Math.max(500, ...cs.map(c => Math.ceil(thaiWrap.longestUnit(c.s) * c.pt * 20) + pad + 60)); return need > W[j] ? W[j] : need; });
    // short labels ("ทะเบียนผู้รับบริการ", "ค้นหาด้วย"): Word keeps a space-free Thai chunk together more often than ICU splits it,
    // so a chunk that does not fit is wrapped with huge letter gaps -> penalise widths narrower than the chunk
    const chunkW = c => Math.max(0, ...c.s.trim().split(/\s+/).filter(x => /[฀-๿]/.test(x)).map(x => thaiWrap.chunkWidth(x)));
    const score = (j, w) => cells[j].reduce((a, c) => {
      const em = (w - pad) / 20 / c.pt;
      let s = Math.max(0, thaiWrap.maxGap(c.s, em) - GAP_EM);
      if (c.s.trim().length <= 30) { const cw = chunkW(c); if (cw > em && cw < em * 1.8) s += (cw - em) * 0.6; }
      return a + s;
    }, 0);
    const cur = [...W]; let sc = cur.map((w, j) => score(j, w)); const STEP = 60;
    for (let it = 0; it < 120; it++) {
      const need = sc.indexOf(Math.max(...sc)); if (sc[need] <= 0) break;
      let best = null;
      for (let d = 0; d < n; d++) {
        if (d === need || cur[d] - STEP < minW[d]) continue;
        const gain = sc[need] - score(need, cur[need] + STEP), loss = score(d, cur[d] - STEP) - sc[d];
        if (gain - loss > 1e-6 && (!best || gain - loss > best.v)) best = { d, v: gain - loss };
      }
      if (!best) break;
      cur[need] += STEP; cur[best.d] -= STEP; sc[need] = score(need, cur[need]); sc[best.d] = score(best.d, cur[best.d]);
    }
    if (cur.every((w, i) => w === W[i]) && tally.size === 1) continue;
    colGroups++;
    console.log(`  same-header tables x${tbls.length}: widths ${[...tally.keys()].join(' / ')} -> ${cur.join(',')}`);
    for (const t of tbls) {
      els(el(t, 'w:tblGrid'), 'w:gridCol').forEach((c, i) => setAttr(c, 'w:w', String(cur[i])));
      for (const tr of els(t, 'w:tr')) els(tr, 'w:tc').forEach((tc, i) => { const w = el(el(tc, 'w:tcPr') || { children: [] }, 'w:tcW'); if (w && attr(w, 'w:type') === 'dxa') setAttr(w, 'w:w', String(cur[i])); });
      colTablesChanged++;
    }
  }
}
console.log(`column widths balanced: ${colTablesChanged} table(s) in ${colGroups} group(s) of same-header tables`);

// 6) every row ends at the same right edge: rows whose cell widths (tcW, plus w:wAfter / gridAfter) add up to less or more
//    than the widest row get their LAST cell stretched / shrunk (skipped for tables with vertically merged cells)
let evened = 0;
for (const tbl of all(root, 'w:tbl')) {
  if (all(tbl, 'w:vMerge').length) continue;
  const trs = els(tbl, 'w:tr'); if (trs.length < 2) continue;
  const rowW = tr => { let s = 0; for (const tc of els(tr, 'w:tc')) { const p = el(tc, 'w:tcPr'); const w = p && el(p, 'w:tcW'); if (!w || attr(w, 'w:type') !== 'dxa') return null; s += +attr(w, 'w:w'); } const tp = el(tr, 'w:trPr'); const wa = tp && el(tp, 'w:wAfter'); if (wa) s += +attr(wa, 'w:w') || 0; return s; };
  let changed = false;
  // empty space left after a row (gridAfter / wAfter) makes the right edge look stepped -> give that space to the last cell
  for (const tr of trs) {
    const tp = el(tr, 'w:trPr'); const ga = tp && el(tp, 'w:gridAfter'); if (!ga) continue; const wa = el(tp, 'w:wAfter');
    const n = +attr(ga, 'w:val') || 0; const add = wa ? +attr(wa, 'w:w') || 0 : 0;
    const tcs = els(tr, 'w:tc'); const last = tcs[tcs.length - 1]; const p = el(last, 'w:tcPr'); if (!p || n <= 0) continue;
    let gs = el(p, 'w:gridSpan'); if (!gs) { gs = mk('w:gridSpan', [['w:val', '1']]); gs.parent = p; const iw = p.children.findIndex(c => c.type === 'el' && c.name === 'w:tcW'); p.children.splice(iw + 1, 0, gs); }
    setAttr(gs, 'w:val', String((+attr(gs, 'w:val') || 1) + n));
    const w = el(p, 'w:tcW'); if (w && attr(w, 'w:type') === 'dxa') setAttr(w, 'w:w', String(+attr(w, 'w:w') + add));
    tp.children.splice(tp.children.indexOf(ga), 1); if (wa) tp.children.splice(tp.children.indexOf(wa), 1);
    changed = true;
  }
  const ws = trs.map(rowW); if (ws.some(x => x === null)) { if (changed) evened++; continue; }
  const target = Math.max(...ws);
  trs.forEach((tr, i) => {
    const diff = target - ws[i]; if (Math.abs(diff) < 30) return;   // < ~0.5 mm: ignore
    const tp = el(tr, 'w:trPr'); const wa = tp && el(tp, 'w:wAfter'); const ga = tp && el(tp, 'w:gridAfter');
    if (wa) { wa.parent.children.splice(wa.parent.children.indexOf(wa), 1); } if (ga) { ga.parent.children.splice(ga.parent.children.indexOf(ga), 1); }
    const tcs = els(tr, 'w:tc'); const last = tcs[tcs.length - 1]; const w = el(el(last, 'w:tcPr'), 'w:tcW');
    setAttr(w, 'w:w', String(+attr(w, 'w:w') + target - ws[i] + (wa ? +attr(wa, 'w:w') || 0 : 0)));
    changed = true;
  });
  if (changed) evened++;
}
console.log(`tables with uneven row edges evened: ${evened}`);

// 6b) body text broken mid-sentence into two paragraphs -> one paragraph
//     (outside tables we only merge when the break is clearly inside a sentence: the first part ends with,
//      or the second part starts with, a joining word such as และ/หรือ/ของ/ที่/โดย/เพื่อ/ซึ่ง/กับ/ใน/จาก/ให้)
const JOIN_END = /(และ|หรือ|ของ|ที่|โดย|เพื่อ|ซึ่ง|กับ|ใน|จาก|ให้|แก่|แต่|เช่น|คือ|ได้แก่|รวมถึง|ตาม|จะ|ต้อง|ได้|ไป|มา|ผ่าน|ด้วย)\s*$/;
const JOIN_START = /^\s*(และ|หรือ|ซึ่ง|โดย|เพื่อ|ที่|ของ|กับ|จาก|ให้|แต่|รวมถึง|ตลอดจน|ด้วย)/;
let bodyMerged = 0;
// OFF by default: on a real book it merged two genuine paragraphs ("…ประชาชนทั่วไป" + "จากปัญหา…"). Enable only on request.
if (cfg.merge_broken_paragraphs_in_body === true) {
  let kids2 = els(body);
  for (let k = 1; k < kids2.length; k++) {
    const a = kids2[k - 1], b = kids2[k];
    if (a.name !== 'w:p' || b.name !== 'w:p' || !plainPara(a) || !plainPara(b)) continue;
    const sa = el(el(a, 'w:pPr') || { children: [] }, 'w:pStyle'), sb = el(el(b, 'w:pPr') || { children: [] }, 'w:pStyle');
    if ((sa && attr(sa, 'w:val')) !== (sb && attr(sb, 'w:val'))) continue;
    const ta = textOf(a), tb = textOf(b);
    if (LIST_START.test(tb) || /[:：.]\s*$/.test(ta)) continue;
    if (!/[฀-๿]$/.test(ta.trim()) || !/^[฀-๿]/.test(tb.trim())) continue;
    if (!JOIN_END.test(ta) && !JOIN_START.test(tb)) continue;
    if (!/\s$/.test(ta) && !/^\s/.test(tb)) { const sp = mk('w:r', [], [mk('w:t', [['xml:space', 'preserve']], [{ type: 'text', text: ' ' }])]); sp.parent = a; a.children.push(sp); }
    for (const c of [...b.children]) { if (c.type === 'el' && c.name === 'w:pPr') continue; b.children.splice(b.children.indexOf(c), 1); c.parent = a; a.children.push(c); }
    remove(b); bodyMerged++; kids2 = els(body); k--;
  }
}
console.log(`body paragraphs broken mid-sentence merged: ${bodyMerged}`);

// 6c) empty paragraph sitting between two items of the same numbered list / same heading-style series
//     (e.g. a blank line between 1.2.3 and 1.2.4) -> remove the blank line
let gapsRemoved = 0;
{
  const pst = p => { const pp = el(p, 'w:pPr'); const s = pp && el(pp, 'w:pStyle'); return s ? attr(s, 'w:val') : null; };
  const styleNum = {}; for (const m of stylesXml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) { const id = (m[1].match(/w:styleId="([^"]*)"/) || [])[1]; if (id && /<w:numPr>/.test(m[2])) styleNum[id] = true; }
  const chainNum = id => { for (let n = 0; id && n < 10; n++, id = styleBase[id]) if (styleNum[id]) return true; return false; };
  const numbered = p => { const pp = el(p, 'w:pPr'); return !!(pp && el(pp, 'w:numPr')) || chainNum(pst(p)); };
  const blank = p => p.name === 'w:p' && !textOf(p).trim() && !all(p, 'w:drawing').length && !all(p, 'w:pict').length && !all(p, 'w:sectPr').length && !all(p, 'w:br').some(b => attr(b, 'w:type') === 'page') && !all(p, 'w:bookmarkStart').length && !all(p, 'w:fldChar').length;
  let ks = els(body);
  for (let k = 1; k < ks.length - 1; k++) {
    const a = ks[k - 1], g = ks[k], b = ks[k + 1];
    if (a.name !== 'w:p' || b.name !== 'w:p' || !blank(g)) continue;
    if (!textOf(a).trim() || !textOf(b).trim()) continue;
    if (pst(a) && pst(a) === pst(b) && numbered(a)) { remove(g); gapsRemoved++; ks = els(body); k--; }
  }
}
console.log(`blank lines between numbered items removed: ${gapsRemoved}`);

// 7) body paragraphs (outside tables, after the cover section) -> Thai Distributed
if (BODY_THAI) {
  const bodyParas = els(body).filter(c => c.name === 'w:p');
  let pastCover = !bodyParas.some(c => { const pp = el(c, 'w:pPr'); return pp && el(pp, 'w:sectPr'); });   // one section only: no separate cover
  for (const c of els(body)) {
    if (c.name !== 'w:p') continue;
    const pPr = el(c, 'w:pPr');
    if (!pastCover) { if (pPr && el(pPr, 'w:sectPr')) pastCover = true; continue; }   // first section = cover: leave alone
    if (cfg.join_soft_breaks_in_body !== false) joinSoftBreaks(c);
    if (canThaiDistribute(c)) {
      const jc = pPr && el(pPr, 'w:jc');
      if (goodForThaiDistribute(c, false) === 'thaiDistribute') { if (!jc || attr(jc, 'w:val') !== 'thaiDistribute') { setJc(c, 'thaiDistribute'); bodyThai++; } }
      else if (!jc || /^(thaiDistribute|distribute)$/.test(attr(jc, 'w:val'))) setJc(c, FALLBACK);
    }
  }
}
// 7a) long English names/keys in table cells ("EMAIL_ENV_PLAN_APPROVED", "AUTH.SESSION_IDLE_MINUTES", "triage_keywords")
//     cannot wrap, so Word cuts them mid-word ("APPROVE / D") or leaves a lonely Thai word that Thai Distributed stretches.
//     Insert Word's invisible "No-Width Optional Break" (U+200B) after _ . / inside such names -> they wrap at those points.
//     Nothing visible changes (ผู้ใช้เลือกวิธีนี้ 25-11-2026). Not in code-font runs.
let zwspAdded = 0;
if (cfg.optional_breaks_in_names !== false) {
  const NAME = /[A-Za-z0-9{(][A-Za-z0-9_.\/{}()-]{10,}/g;
  for (const tc of all(root, 'w:tc')) { if (els(tc, 'w:p').some(p => looksCode(textOf(p)))) continue;   // code / JSON samples: copy-paste must stay exact
  for (const r of all(tc, 'w:r')) {
    const rp = el(r, 'w:rPr'); const rf = rp && el(rp, 'w:rFonts'); if (rf && CODE.test(attr(rf, 'w:ascii') || attr(rf, 'w:hAnsi') || '')) continue;
    for (const t of els(r, 'w:t')) {
      const s = t.children.map(c => c.text || '').join('');
      if (!/[_.\/]/.test(s) || /https?:|www\./i.test(s)) continue;
      // not in {placeholders}, not in URL/API paths ("/api/…", "HTTP/1.1") — only names joined by _ or .
      const ns = s.replace(NAME, m => /[_.]/.test(m) && /[A-Za-z]/.test(m) && !/\//.test(m) && !/^\{.*\}?$/.test(m) && !/\{[^}]*$/.test(m) ? m.replace(/([_.])(?=[A-Za-z0-9{(])/g, (x, c) => { zwspAdded++; return c + '​'; }) : m);
      if (ns !== s) t.children = [{ type: 'text', text: ns }];
    }
  } }
}
console.log(`invisible break points added inside long names: ${zwspAdded}`);

// 7a1) a column holding only code names (EMAIL_ENV_PLAN_APPROVED …) can be narrower now that names wrap at "_":
//      give that width to Thai columns that stretch (ผู้ใช้เลือก 25-11-2026: "ลดคอลัมน์รหัส เพิ่มให้ผู้รับ/หัวเรื่อง").
//      Total width unchanged; tables with the same header get the same widths.
let codeColsShrunk = 0;
if (thaiWrap.available && cfg.shrink_code_columns !== false) {
  const kinds = new Map();
  for (const tbl of all(root, 'w:tbl')) {
    const trs = els(tbl, 'w:tr'); if (trs.length < 3 || !looksLikeHeader(tbl)) continue;
    const g = el(tbl, 'w:tblGrid'); const cols = g ? els(g, 'w:gridCol') : []; const n = cols.length;
    // rows are either n plain cells, or one cell spanning the whole table (picture / section rows)
    const full = tr => { const tcs = els(tr, 'w:tc'); return tcs.length === 1 && +attr(el(el(tcs[0], 'w:tcPr') || { children: [] }, 'w:gridSpan') || { attrs: [] }, 'w:val') === n; };
    if (n < 3 || !trs.every(tr => (els(tr, 'w:tc').length === n && !all(tr, 'w:gridSpan').some(x => +attr(x, 'w:val') > 1)) || full(tr))) continue;
    const key = n + '#' + els(trs[0], 'w:tc').map(tc => textOf(tc).trim()).join('|');
    if (!kinds.has(key)) kinds.set(key, []); kinds.get(key).push(tbl);
  }
  for (let tbls of kinds.values()) {
    // tables of one kind often differ by a few twips -> start from the widths most of them use; all end up identical
    const wOf = t => els(el(t, 'w:tblGrid'), 'w:gridCol').map(c => +attr(c, 'w:w'));
    const tl = new Map(); for (const t of tbls) { const k = wOf(t).join(','); tl.set(k, (tl.get(k) || 0) + 1); }
    const W = [...tl.entries()].sort((a, b) => b[1] - a[1])[0][0].split(',').map(Number); const n = W.length;
    const tot = W.reduce((a, b) => a + b, 0);
    tbls = tbls.filter(t => Math.abs(wOf(t).reduce((a, b) => a + b, 0) - tot) < tot * 0.05);   // odd-sized ones keep their own widths
    if (!tbls.length) continue;
    const colCells = j => tbls.flatMap(t => els(t, 'w:tr').slice(1).filter(tr => els(tr, 'w:tc').length === n).map(tr => els(tr, 'w:tc')[j]));
    const pad = cellPadTw(all(tbls[0], 'w:p')[0] || tbls[0]); const pt = n >= (cfg.dense_table_min_cols || 7) ? (cfg.dense_table_font_pt || 14) : cellFontPt(all(tbls[0], 'w:p')[1] || all(tbls[0], 'w:p')[0]);
    const cur = [...W]; let freed = 0;
    for (let j = 0; j < n; j++) {
      const texts = colCells(j).map(tc => textOf(tc).trim()).filter(Boolean);
      const codeCol = texts.every(t => /^[A-Za-z0-9_.\-​]+$/.test(t)) && texts.some(t => /​/.test(t));
      // short-label columns ("PASS", "1-1", "VIEW_001", "การดูข้อมูล"): fit the longest label, header word included
      const labelCol = texts.length >= 2 && texts.every(t => [...t.replace(/​/g, '')].length <= 15);
      if (texts.length < 2 || (!codeCol && !labelCol)) continue;
      const hc = els(els(tbls[0], 'w:tr')[0], 'w:tc')[j]; const headTxt = hc ? textOf(hc).trim() : '';
      const pieces = codeCol ? texts.flatMap(t => t.split('​')) : texts.concat(headTxt.split(/\s+/));
      // measure with the table's REAL text size (largest run size in the column) + bold margin, never narrower than now-200
      const colPt = Math.max(pt, ...colCells(j).flatMap(tc => all(tc, 'w:sz').map(s => (+attr(s, 'w:val') || 0) / 2)));
      const need = Math.ceil(Math.max(...pieces.map(x => thaiWrap.longestUnit(codeCol ? x : x.replace(/\s+/g, '\u00A0')))) * colPt * 20) + pad + 200;
      if (need < W[j] - 200) { freed += W[j] - need; cur[j] = need; }
    }
    if (process.env.RF_DEBUG && !freed) console.log('  dbg nofreed', tbls.length, W.join(','));
    if (!freed) continue;
    // hand the freed width to Thai columns, the more they stretch the more they get
    const thaiCols = [...Array(n).keys()].filter(j => cur[j] === W[j] && colCells(j).some(tc => /[฀-๿]/.test(textOf(tc))));
    const stretch = j => colCells(j).reduce((a, tc) => a + Math.max(0, thaiWrap.maxGap(textOf(tc), (cur[j] - pad) / 20 / pt) - GAP_EM), 0) + 0.01;
    const wts = thaiCols.map(stretch); const sum = wts.reduce((a, b) => a + b, 0);
    let given = 0; thaiCols.forEach((j, k) => { const add = k === thaiCols.length - 1 ? freed - given : Math.round(freed * wts[k] / sum / 10) * 10; cur[j] += add; given += add; });
    // keep the change only if the group as a whole stretches less than before
    const total = ws => [...Array(n).keys()].reduce((a, j) => a + colCells(j).reduce((b, tc) => b + Math.max(0, thaiWrap.maxGap(textOf(tc), (ws[j] - pad) / 20 / pt) - GAP_EM), 0), 0);
    if (process.env.RF_DEBUG) console.log('  dbg kind', tbls.length, W.join(','), '->', cur.join(','), 'freed', freed, 'thaiCols', thaiCols.join('/'), 'before', total(W).toFixed(2), 'after', total(cur).toFixed(2));
    if (!thaiCols.length || total(cur) >= total(W) - 0.01) continue;
    for (const t of tbls) {
      els(el(t, 'w:tblGrid'), 'w:gridCol').forEach((c, i) => setAttr(c, 'w:w', String(cur[i])));
      for (const tr of els(t, 'w:tr')) { const tcs = els(tr, 'w:tc'); tcs.forEach((tc, i) => { const w = el(el(tc, 'w:tcPr') || { children: [] }, 'w:tcW'); if (w && attr(w, 'w:type') === 'dxa') setAttr(w, 'w:w', String(tcs.length === 1 ? cur.reduce((a, b) => a + b, 0) : cur[i])); }); }
      codeColsShrunk++;
    }
    console.log(`  code column narrowed: ${W.join(',')} -> ${cur.join(',')} | ${textOf(els(els(tbls[0], 'w:tr')[1], 'w:tc')[0]).replace(/​/g, '').trim().slice(0, 40)}`);
  }
}
console.log(`tables whose code-name column was narrowed: ${codeColsShrunk}`);

// 7a2) connector / filler words where Thai Distributed would still stretch letters badly (ผู้ใช้อนุญาต 25-11-2026:
//      "เติมคำสร้อย/คำเชื่อมที่ไม่กระทบเนื้อหาได้"). Only in front of a system message / e-mail subject, never inside it,
//      only when the simulated stretch really drops; every change is printed as "filler:" for the summary.
const fillers = [];
if (thaiWrap.available && cfg.filler_words !== false) {
  const CANDS = [
    // many lengths, so a cell can pick the one whose line breaks fall best (ผู้ใช้ให้สิทธิ์ปรับคำเต็มที่ 25-11-2026)
    [/^แจ้ง\s+(?=[A-Za-z"“])/, ['แจ้งข้อความ ', 'ระบบแจ้ง ', 'ระบบแจ้งข้อความ ', 'ระบบแจ้งข้อความว่า ', 'ระบบจะแจ้งข้อความ ', 'ระบบจะแจ้งข้อความว่า ',
      'แจ้งเตือนด้วยข้อความ ', 'ระบบแสดงข้อความแจ้งเตือน ', 'ระบบแสดงข้อความแจ้งเตือนว่า ', 'ระบบจะแสดงข้อความแจ้งเตือนว่า ', 'ระบบจะแสดงข้อความแจ้งเตือนผู้ใช้ว่า ']],
    [/^(.*?\S)\s+แจ้ง\s+(?=[A-Za-z"“])/, ['$1 แจ้งข้อความ ', '$1 ระบบแจ้งข้อความ ', '$1 ระบบจะแจ้งข้อความว่า ', '$1 ระบบแสดงข้อความแจ้งเตือน ', '$1 ระบบจะแสดงข้อความแจ้งเตือนว่า ']],
    [/^(?=\[[A-Za-z])/, ['หัวเรื่อง ', 'ใช้หัวเรื่อง ']],
    // ผู้ใช้เลือก (25-11-2026) สำหรับคอลัมน์ผู้รับ: "บุคลากรของสถานบริการ…" -> "ส่งถึงผู้ใช้ที่เป็นบุคลากรของสถานบริการ…"
    [/^บุคลากรของ/, ['ส่งถึงผู้ใช้ที่เป็นบุคลากรของ', 'ส่งถึงบุคลากรของ']],
  ];
  // group cells by (same table header, column) so a column is changed the SAME way everywhere (no half-filled columns)
  const groups = new Map();
  for (const tbl of all(root, 'w:tbl')) {
    const trs = els(tbl, 'w:tr'); if (!trs.length) continue;
    const head = looksLikeHeader(tbl) ? els(trs[0], 'w:tc').map(tc => textOf(tc).trim()).join('|') : 'T' + Math.random();
    trs.forEach((tr, ri) => { if (ri === 0 && looksLikeHeader(tbl)) return; els(tr, 'w:tc').forEach((tc, ci) => { for (const p of els(tc, 'w:p')) {
      const s = textOf(p); if (!/[฀-๿]/.test(s)) continue;
      const jc = el(el(p, 'w:pPr') || { children: [] }, 'w:jc'); if (jc && /center|right/.test(attr(jc, 'w:val'))) continue;
      const w = cellWidth(p); if (!Number.isFinite(w)) continue;
      const k = head + '#' + ci; if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push({ p, s, em: (w - cellPadTw(p)) / 20 / cellFontPt(p) });
    } }); });
  }
  const applyTo = (p, re, rep) => {   // replace the matched start inside the first text node(s) (raw XML; patterns hold no entities)
    const ts = all(p, 'w:t'); let acc = '';
    for (const t of ts) {
      const tt = t.children.map(c => c.text || '').join('');
      if (re.test(acc + tt)) { const full = (acc + tt).replace(re, rep); if (!full.startsWith(acc)) return false; t.children = [{ type: 'text', text: full.slice(acc.length) }]; if (!t.attrs.some(a => a[0] === 'xml:space')) t.attrs.push(['xml:space', 'preserve']); return true; }
      acc += tt;
    }
    return false;
  };
  for (const cells of groups.values()) {
    for (const [re, reps] of CANDS) {
      const hit = cells.filter(c => re.test(c.s)); if (!hit.length) continue;
      // every cell picks the wording whose line breaks look best (shortest option wins a tie) — beauty first
      for (const c of hit) {
        const g0 = thaiWrap.maxGap(c.s, c.em); if (g0 <= 0.1) continue;
        const own = reps.map((r, i) => ({ r, i, g: thaiWrap.maxGap(c.s.replace(re, r), c.em) })).sort((a, b) => (a.g - b.g) || (a.i - b.i))[0];
        if (own.g > g0 - 0.03) continue;
        if (applyTo(c.p, re, own.r)) { const ns = c.s.replace(re, own.r); fillers.push(`${c.s.slice(0, 60)}  ->  ${ns.slice(0, 70)}`); c.s = ns; }
      }
    }
  }
}
console.log(`filler words added (cells that stretched badly): ${fillers.length}`);
fillers.forEach(x => console.log('  filler: ' + x));

// 7b) hand-made character spacing (Font > Advanced > Spacing: condensed/expanded, e.g. -0.5 pt) used to squeeze text onto a
//     line. With Thai Distributed it pulls tone marks off their letters ("น ั บ จ า ก"). The team fixed it by copy -> paste as
//     plain text; we remove it everywhere except the cover page (titles may use spacing on purpose).
let charSpacingRemoved = 0;
if (cfg.strip_char_spacing !== false) {
  const bodyKids = els(body); let coverEnd = bodyKids.findIndex(c => c.name === 'w:p' && el(el(c, 'w:pPr') || { children: [] }, 'w:sectPr'));
  if (coverEnd < 0) coverEnd = -1;
  bodyKids.forEach((c, i) => {
    if (i <= coverEnd) return;
    for (const r of all(c, 'w:r')) { const rp = el(r, 'w:rPr'); if (!rp) continue; for (const s of els(rp, 'w:spacing')) { remove(s); charSpacingRemoved++; } for (const s of els(rp, 'w:w')) { remove(s); charSpacingRemoved++; } }
  });
}
console.log(`hand-made character spacing removed: ${charSpacingRemoved}`);

// 7b2) manual page break (Ctrl+Enter) at the start of a paragraph / alone in an empty paragraph -> "page break before"
//      on the paragraph that follows. A page that is already full (e.g. a screenshot row) no longer leaves a blank page.
let pbConverted = 0;
{
  const setPBB = p => { const pp = el(p, 'w:pPr') || ensure(p, 'w:pPr'); if (el(pp, 'w:pageBreakBefore')) return; const n = mk('w:pageBreakBefore'); n.parent = pp; const i = pp.children.findIndex(c => c.type === 'el' && !['w:pStyle', 'w:keepNext', 'w:keepLines'].includes(c.name)); if (i < 0) pp.children.push(n); else pp.children.splice(i, 0, n); };
  const isPB = b => attr(b, 'w:type') === 'page';
  const kids = els(body);
  for (let i = 0; i < kids.length; i++) {
    const p = kids[i]; if (p.name !== 'w:p' || el(el(p, 'w:pPr') || { children: [] }, 'w:sectPr')) continue;
    const brs = all(p, 'w:br').filter(isPB); if (brs.length !== 1) continue;
    const before = []; let seen = false; (function walk(n) { for (const c of n.children || []) { if (c.type !== 'el') continue; if (c === brs[0]) { seen = true; return; } if (c.name === 'w:t' && !seen) before.push(textOf({ children: [c] })); else if (!seen) walk(c); } })(p);
    if (before.join('').trim()) continue;                                    // break comes after some text -> leave it
    const after = textOf(p).trim();
    if (after) { remove(brs[0]); setPBB(p); pbConverted++; continue; }        // "⇟ข้อมูลผลการประเมิน CCAT" heading
    const next = kids[i + 1];
    if (next && next.name === 'w:p' && !all(p, 'w:drawing').length && !all(p, 'w:bookmarkStart').length) { setPBB(next); remove(p); pbConverted++; }
  }
}
console.log(`manual page breaks turned into "page break before": ${pbConverted}`);

// 7b3) tables narrower than the text area -> widen to the full text width (ผู้ใช้: "ยังสามารถขยายตารางได้ เพื่อไม่ให้แปลก").
//      Every column first gets at least its longest word ("Category", "Standard", "Case ID" no longer cut), the rest of the
//      extra width is shared in proportion to the current widths. Same input -> same output, so same-header tables stay alike.
let tablesWidened = 0, denseTables = 0;
if (cfg.widen_tables_to_text !== false) {
  const kids = els(body); let secIdx = 0; const secOf = new Map(); const sects = [];
  kids.forEach(k => { if (k.name === 'w:p') { const sp = el(el(k, 'w:pPr') || { children: [] }, 'w:sectPr'); if (sp) sects.push(sp); } });
  const finalSect = el(body, 'w:sectPr'); if (finalSect) sects.push(finalSect);
  kids.forEach(k => { if (k.name === 'w:tbl') secOf.set(k, sects[secIdx] || finalSect); if (k.name === 'w:p' && el(el(k, 'w:pPr') || { children: [] }, 'w:sectPr')) secIdx++; });
  for (const [tbl, sp] of secOf) {
    if (!sp) continue;
    const pg = el(sp, 'w:pgSz'), mar = el(sp, 'w:pgMar'); if (!pg || !mar) continue;
    const textW = +attr(pg, 'w:w') - (+attr(mar, 'w:left') || 0) - (+attr(mar, 'w:right') || 0) - (+attr(mar, 'w:gutter') || 0);
    // a negative indent pulls the table into the left margin; then it may also run to the right margin (no lopsided gap)
    const tblPr = el(tbl, 'w:tblPr') || ensure(tbl, 'w:tblPr'); let ind = el(tblPr, 'w:tblInd'); let indW = ind && attr(ind, 'w:type') === 'dxa' ? +attr(ind, 'w:w') || 0 : 0;
    if (indW < 0) { setAttr(ind, 'w:w', '0'); indW = 0; }   // left margin = binding: tables never reach into it
    const g = el(tbl, 'w:tblGrid'); const cols = g ? els(g, 'w:gridCol') : []; if (cols.length < 2) continue;
    // dense tables (Test Case: 7+ narrow columns) — ผู้ใช้เลือก 25-11-2026: reach 0.5" into both margins + table text 14 pt
    // count the cells a reader sees per row (merged "Table Name" rows inflate the grid), not grid columns
    const shownCols = Math.max(...els(tbl, 'w:tr').map(tr => els(tr, 'w:tc').length));
    const dense = shownCols >= (cfg.dense_table_min_cols || 7); let rightBleed = 0;
    if (dense) {
      // left margin is for binding (ผู้ใช้สั่ง): never reach into it — start at the left margin, may run to the right
      // up to dense_table_right_edge_min_cm from the paper edge
      const pt2 = String((cfg.dense_table_font_pt || 14) * 2);
      const rm = +attr(mar, 'w:right') || 0; const keep = Math.round((cfg.dense_table_right_edge_min_cm ?? 1.0) * 567);
      if (ind && (+attr(ind, 'w:w') || 0) < 0) { setAttr(ind, 'w:w', '0'); indW = 0; }
      rightBleed = Math.max(0, rm - keep);      for (const r of all(tbl, 'w:r')) { const rp = el(r, 'w:rPr') || ensure(r, 'w:rPr'); for (const [nm, v] of [['w:sz', pt2], ['w:szCs', pt2]]) { let e = el(rp, nm); if (!e) { e = mk(nm, [['w:val', v]]); e.parent = rp; const after = ['w:highlight', 'w:u', 'w:effect', 'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang', 'w:eastAsianLayout', 'w:specVanish', 'w:oMath']; const i = rp.children.findIndex(c => c.type === 'el' && (after.includes(c.name) || (nm === 'w:sz' && c.name === 'w:szCs'))); if (i < 0) rp.children.push(e); else rp.children.splice(i, 0, e); } else setAttr(e, 'w:val', v); } }
      denseTables++;
    }
    const avail = textW - indW + rightBleed;   // from (left margin + indent) to (right margin + bleed)
    const W = cols.map(c => +attr(c, 'w:w') || 0); const sum = W.reduce((a, b) => a + b, 0);
    if (!sum || sum >= avail * 0.97 || sum < avail * 0.5) continue;         // already full width / tiny tables on purpose
    const trs = els(tbl, 'w:tr');
    if (!trs.every(tr => els(tr, 'w:tc').reduce((a, tc) => a + (+attr(el(el(tc, 'w:tcPr') || { children: [] }, 'w:gridSpan') || { attrs: [] }, 'w:val') || 1), 0) === cols.length)) continue;
    // longest word per grid column (single-span cells only)
    const pad = cellPadTw(all(tbl, 'w:p')[0] || tbl); const need = W.map(() => 0);
    for (const tr of trs) { let gi = 0; for (const tc of els(tr, 'w:tc')) { const span = +attr(el(el(tc, 'w:tcPr') || { children: [] }, 'w:gridSpan') || { attrs: [] }, 'w:val') || 1; if (span === 1) for (const p of els(tc, 'w:p')) { const s = textOf(p).replace(/​/g, ' '); if (s.trim() && thaiWrap.available) need[gi] = Math.max(need[gi], Math.ceil(thaiWrap.longestUnit(s) * cellFontPt(p) * 20) + pad + 60); } gi += span; } }
    let extra = avail - sum; const cur = [...W];
    for (let j = 0; j < cur.length && extra > 0; j++) { const add = Math.min(extra, Math.max(0, need[j] - cur[j])); cur[j] += add; extra -= add; }
    if (extra > 0) {
      // the rest goes to columns that hold long text (they wrap and stretch); short-label columns ("PASS", "1-1") get none
      const colTexts = W.map(() => []);
      for (const tr of trs) { const tcs = els(tr, 'w:tc'); if (tcs.length !== cur.length) continue; tcs.forEach((tc, j) => { const s = textOf(tc).replace(/\u200B/g, '').trim(); if (s) colTexts[j].push(s); }); }
      const fpt = dense ? (cfg.dense_table_font_pt || 14) : cellFontPt(all(tbl, 'w:p')[1] || all(tbl, 'w:p')[0]);
      const wt = colTexts.map((ts, j) => ts.some(s => [...s].length > 15) ? ts.reduce((a, s) => a + Math.max(0.05, thaiWrap.maxGap(s, (cur[j] - pad) / 20 / fpt)), 0) + 0.5 : 0);
      let sumW = wt.reduce((a, b) => a + b, 0); const base = sumW ? wt : cur; sumW = sumW || cur.reduce((a, b) => a + b, 0);
      let given = 0; const last = base.reduce((m, v, j) => (v > 0 ? j : m), cur.length - 1);
      cur.forEach((w, j) => { const add = j === last ? extra - given : Math.floor(extra * base[j] / sumW); cur[j] += add; given += add; });
    }
    cols.forEach((c, j) => setAttr(c, 'w:w', String(cur[j])));
    for (const tr of trs) { let gi = 0; for (const tc of els(tr, 'w:tc')) { const tcPr = el(tc, 'w:tcPr'); const span = +attr(el(tcPr || { children: [] }, 'w:gridSpan') || { attrs: [] }, 'w:val') || 1; const w = tcPr && el(tcPr, 'w:tcW'); if (w && attr(w, 'w:type') === 'dxa') setAttr(w, 'w:w', String(cur.slice(gi, gi + span).reduce((a, b) => a + b, 0))); gi += span; } }
    const tw = tblPr && el(tblPr, 'w:tblW'); if (tw) { setAttr(tw, 'w:type', 'dxa'); setAttr(tw, 'w:w', String(avail)); }
    tablesWidened++;
  }
}
console.log(`tables widened to the text width: ${tablesWidened} (dense 7+ column tables, 14 pt, run into the right margin: ${denseTables})`);

// 7b4) every column at least as wide as its widest word / code / header word ("TCASE_001", "Test Result" -> no
//      "TCASE_0/01", "Resul/t"). Missing width is taken from the columns with the most spare room. Table width unchanged.
//      Also: text in tables never smaller than table_min_font_pt (ผู้ใช้: ตัวอักษรตาราง 8–9 pt เล็กเกินไป).
let colsFixed = 0, smallFontRuns = 0;
if (thaiWrap.available) {
  const MINPT = cfg.table_min_font_pt || 14;
  for (const tbl of all(root, 'w:tbl')) {
    for (const sz of [...all(tbl, 'w:sz'), ...all(tbl, 'w:szCs')]) { const v = +attr(sz, 'w:val') || 0; if (v && v < MINPT * 2 && !(sz.parent && all(sz.parent, 'w:rFonts').some(f => CODE.test(attr(f, 'w:ascii') || '')))) { setAttr(sz, 'w:val', String(MINPT * 2)); smallFontRuns++; } }
    const g = el(tbl, 'w:tblGrid'); const cols = g ? els(g, 'w:gridCol') : []; const n = cols.length; if (n < 2) continue;
    const trs = els(tbl, 'w:tr');
    const plain = trs.filter(tr => els(tr, 'w:tc').length === n && !all(tr, 'w:gridSpan').some(x => +attr(x, 'w:val') > 1));
    if (plain.length < Math.max(1, trs.length * 0.5)) continue;
    const W = cols.map(c => +attr(c, 'w:w') || 0); if (W.some(w => !w)) continue;
    const pad = cellPadTw(all(tbl, 'w:p')[0] || tbl);
    const need = W.map((_, j) => Math.max(0, ...plain.map(tr => { const tc = els(tr, 'w:tc')[j]; if (all(tc, 'w:drawing').length) return 0; return Math.max(0, ...els(tc, 'w:p').map(p => { const s = textOf(p); if (!s.trim() || looksCode(s)) return 0; const ptx = Math.max(cellFontPt(p), ...all(p, 'w:sz').map(z => (+attr(z, 'w:val') || 0) / 2)); return Math.ceil(thaiWrap.longestUnit(s) * ptx * 20 * 1.1) + pad + 250; })); })));
    let deficit = need.reduce((a, nd, j) => a + Math.max(0, nd - W[j]), 0);
    if (process.env.RF_DEBUG) console.log('  dbgA', n, W.join(','), 'need', need.join(','));
    if (!deficit) continue;
    const spare = W.map((w, j) => Math.max(0, w - need[j] - 300));
    const totalSpare = spare.reduce((a, b) => a + b, 0);
    if (process.env.RF_DEBUG) console.log('  dbg7b4', W.join(','), 'need', need.join(','), 'deficit', deficit, 'spare', totalSpare);
    let spareUse = spare, sumSpare = totalSpare;
    if (sumSpare < deficit) { spareUse = W.map((w, j) => Math.max(0, w - need[j])); sumSpare = spareUse.reduce((a, b) => a + b, 0); }
    if (sumSpare < deficit) {   // still short: every column gets need scaled to the table width (forced, ผู้ใช้สั่งบังคับ)
      const tot = W.reduce((a, b) => a + b, 0), nsum = need.reduce((a, b) => a + b, 0);
      const cur2 = need.map(nd => Math.floor(nd * tot / nsum)); cur2[cur2.length - 1] += tot - cur2.reduce((a, b) => a + b, 0);
      cols.forEach((c, j) => setAttr(c, 'w:w', String(cur2[j])));
      for (const tr of trs) { const tcs = els(tr, 'w:tc'); let gi = 0; for (const tc of tcs) { const tcPr = el(tc, 'w:tcPr'); const span = +attr(el(tcPr || { children: [] }, 'w:gridSpan') || { attrs: [] }, 'w:val') || 1; const w = tcPr && el(tcPr, 'w:tcW'); if (w && attr(w, 'w:type') === 'dxa') setAttr(w, 'w:w', String(cur2.slice(gi, gi + span).reduce((a, b) => a + b, 0))); gi += span; } }
      colsFixed++; continue;
    }
    const cur = [...W]; need.forEach((nd, j) => { if (nd > cur[j]) cur[j] = nd; });
    let taken = 0; const donors = spareUse.map((s, j) => [s, j]).filter(x => x[0] > 0);
    donors.forEach(([s, j], k) => { const t = k === donors.length - 1 ? deficit - taken : Math.round(deficit * s / sumSpare); cur[j] -= t; taken += t; });
    cols.forEach((c, j) => setAttr(c, 'w:w', String(cur[j])));
    for (const tr of trs) { const tcs = els(tr, 'w:tc'); let gi = 0; for (const tc of tcs) { const tcPr = el(tc, 'w:tcPr'); const span = +attr(el(tcPr || { children: [] }, 'w:gridSpan') || { attrs: [] }, 'w:val') || 1; const w = tcPr && el(tcPr, 'w:tcW'); if (w && attr(w, 'w:type') === 'dxa') setAttr(w, 'w:w', String(cur.slice(gi, gi + span).reduce((a, b) => a + b, 0))); gi += span; } }
    colsFixed++;
  }
}
console.log(`tables whose too-narrow columns were widened to fit their longest word: ${colsFixed}; table runs raised to the minimum font size: ${smallFontRuns}`);

// 7b5) a heading with a typed sub-number ("3.3 โครงสร้างฐานข้อมูล") wearing a CHAPTER style -> use the level-2 style
//      (it showed as a chapter line in the table of contents)
let levelFixed = 0;
{
  const st = doc.get('word/styles.xml') || ''; const S = {};
  for (const m of st.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) { const id = (m[1].match(/w:styleId="([^"]*)"/) || [])[1]; if (id) S[id] = { o: (m[2].match(/<w:outlineLvl w:val="(\d)"/) || [])[1], base: (m[2].match(/<w:basedOn w:val="([^"]*)"/) || [])[1], name: (m[2].match(/<w:name w:val="([^"]*)"/) || [])[1] }; }
  const lvl = id => { for (let i = 0; id && i < 12; i++, id = S[id] && S[id].base) if (S[id] && S[id].o != null) return +S[id].o; return 9; };
  const counts = {}; for (const p of all(body, 'w:p')) { const ps = el(el(p, 'w:pPr') || { children: [] }, 'w:pStyle'); if (ps) { const id = attr(ps, 'w:val'); if (lvl(id) === 1) counts[id] = (counts[id] || 0) + 1; } }
  const l2 = Object.entries(counts).sort((a, b) => b[1] - a[1])[0];
  if (l2) for (const p of all(body, 'w:p')) {
    if (p.parent && p.parent.name === 'w:tc') continue;
    const ps = el(el(p, 'w:pPr') || { children: [] }, 'w:pStyle'); if (!ps || lvl(attr(ps, 'w:val')) !== 0) continue;
    if (/^\s*\d+\.\d+(\.\d+)*\s+\S/.test(textOf(p))) { setAttr(ps, 'w:val', l2[0]); levelFixed++; }
  }
}
console.log(`sub-numbered headings moved from chapter level to level 2: ${levelFixed}`);

// 7c) table-of-contents look: (1) TOC / figure-list styles often keep the template's theme fonts (Calibri + Angsana/Cordia,
//     10–12 pt) -> Thai and English in different fonts/sizes. Set them to the book font, TOC size. (2) the "เรื่อง … หน้า"
//     header line: no first-line indent, left aligned, so "หน้า" stays on the same line.
let tocStyles = 0, tocHeads = 0;
{
  const TOC_SZ = String((cfg.toc_font_pt || 16) * 2);
  const fixed = (doc.get('word/styles.xml') || '').replace(/<w:style\b[^>]*>[\s\S]*?<\/w:style>/g, s => {
    const name = ((s.match(/<w:name w:val="([^"]*)"/) || [])[1] || '').toLowerCase();
    if (!/^(toc \d|table of figures|toc heading)$/.test(name)) return s;
    let r = s;
    const fonts = `<w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:eastAsia="${FONT}" w:cs="${FONT}"/>`;
    if (/<w:rPr>/.test(r)) {
      r = /<w:rFonts\b[^>]*\/>/.test(r) ? r.replace(/<w:rFonts\b[^>]*\/>/, fonts) : r.replace(/<w:rPr>/, '<w:rPr>' + fonts);
      if (name !== 'toc heading') {
        r = /<w:sz w:val="\d+"\/>/.test(r) ? r.replace(/<w:sz w:val="\d+"\/>/, `<w:sz w:val="${TOC_SZ}"/>`) : r.replace(/<\/w:rPr>/, `<w:sz w:val="${TOC_SZ}"/></w:rPr>`);
        r = /<w:szCs w:val="\d+"\/>/.test(r) ? r.replace(/<w:szCs w:val="\d+"\/>/, `<w:szCs w:val="${TOC_SZ}"/>`) : r.replace(/<\/w:rPr>/, `<w:szCs w:val="${TOC_SZ}"/></w:rPr>`);
      }
    } else r = r.replace(/<\/w:style>$/, `<w:rPr>${fonts}${name !== 'toc heading' ? `<w:sz w:val="${TOC_SZ}"/><w:szCs w:val="${TOC_SZ}"/>` : ''}</w:rPr></w:style>`);
    if (r !== s) tocStyles++;
    return r;
  });
  if (tocStyles) doc.set('word/styles.xml', fixed);
  const tocTab = (serialize(body).match(/<w:tab w:val="right" w:leader="dot" w:pos="(\d+)"/) || [])[1] || '9061';
  for (const p of all(body, 'w:p')) {
    const t = textOf(p).replace(/\s+/g, ' ').trim();
    if (!/^(เรื่อง|รายการ|ตารางที่|รูปที่|ภาพที่)\s+หน้า$/.test(t) || !all(p, 'w:tab').length && !/ {3,}/.test(textOf(p))) continue;
    // rebuild as  "เรื่อง <one right tab at the TOC edge> หน้า"  (many tabs + spaces wrapped "หน้า" onto its own line)
    { const runs = els(p, 'w:r'); const r0 = runs.find(r => els(r, 'w:t').length) || runs[0];
      const rp = el(r0, 'w:rPr'); const left = t.split(' ')[0];
      const mkRun = kids => { const r = mk('w:r', [], []); if (rp) { const c = parseXml(serialize(rp)).children[0]; c.parent = r; r.children.push(c); } kids.forEach(k => { k.parent = r; r.children.push(k); }); return r; };
      runs.forEach(remove);
      [mkRun([mk('w:t', [], [{ type: 'text', text: left }])]), mkRun([mk('w:tab')]), mkRun([mk('w:t', [], [{ type: 'text', text: 'หน้า' }])])].forEach(r => { r.parent = p; p.children.push(r); });
      const pp0 = el(p, 'w:pPr') || ensure(p, 'w:pPr'); let tabs = el(pp0, 'w:tabs'); if (tabs) remove(tabs);
      tabs = mk('w:tabs', [], [mk('w:tab', [['w:val', 'right'], ['w:pos', tocTab]])]); tabs.parent = pp0;
      const i = pp0.children.findIndex(c => c.type === 'el' && !['w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr', 'w:widowControl', 'w:numPr', 'w:suppressLineNumbers', 'w:pBdr', 'w:shd'].includes(c.name));
      if (i < 0) pp0.children.push(tabs); else pp0.children.splice(i, 0, tabs); }
    const pp = el(p, 'w:pPr') || ensure(p, 'w:pPr');
    let ind = el(pp, 'w:ind'); if (!ind) { ind = mk('w:ind'); ind.parent = pp; const i = pp.children.findIndex(c => c.type === 'el' && ['w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc', 'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange'].includes(c.name)); if (i < 0) pp.children.push(ind); else pp.children.splice(i, 0, ind); }
    ind.attrs = [['w:left', '0'], ['w:firstLine', '0']];
    setJc(p, 'left'); tocHeads++;
  }
}
console.log(`TOC styles set to the book font: ${tocStyles}; "เรื่อง … หน้า" header lines fixed: ${tocHeads}`);

// 7d) table cells that wear a HEADING style ("Test ID" header cells) show up as entries in the table of contents
//     ("Test ID ……… 28" × 150). Give them the normal paragraph style and keep their look as direct formatting.
let cellHeadings = 0;
{
  const st = doc.get('word/styles.xml') || '';
  const sty = {}; for (const m of st.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) { const id = (m[1].match(/w:styleId="([^"]*)"/) || [])[1]; if (id) sty[id] = m[2]; }
  const outline = id => { for (let i = 0; id && i < 12; i++) { const s = sty[id]; if (!s) return 9; const o = s.match(/<w:outlineLvl w:val="(\d)"/); if (o) return +o[1]; id = (s.match(/<w:basedOn w:val="([^"]*)"/) || [])[1]; } return 9; };
  const look = id => { const out = {}; const chain = []; for (let i = 0; id && i < 12; i++) { const s = sty[id]; if (!s) break; chain.unshift(s); id = (s.match(/<w:basedOn w:val="([^"]*)"/) || [])[1]; }
    for (const s of chain) { const rp = (s.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/) || [])[1] || ''; for (const nm of ['b', 'bCs', 'sz', 'szCs', 'color']) { const m = rp.match(new RegExp(`<w:${nm}(?: [^>]*)?\\/>`)); if (m) out[nm] = m[0]; } const jc = (s.match(/<w:pPr>[\s\S]*?(<w:jc [^>]*\/>)/) || [])[1]; if (jc) out.jc = jc; }
    return out; };
  for (const tc of all(body, 'w:tc')) for (const p of els(tc, 'w:p')) {
    const pPr = el(p, 'w:pPr'); const ps = pPr && el(pPr, 'w:pStyle'); if (!ps) continue;
    const id = attr(ps, 'w:val'); if (outline(id) >= 9) continue;
    const lk = look(id);
    for (const r of els(p, 'w:r')) {
      const rp = el(r, 'w:rPr') || ensure(r, 'w:rPr');
      for (const nm of ['b', 'bCs', 'color', 'sz', 'szCs']) if (lk[nm] && !el(rp, 'w:' + nm)) { const n = parseXml(lk[nm]).children[0]; n.parent = rp; const order = ['w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps', 'w:strike', 'w:dstrike', 'w:outline', 'w:shadow', 'w:emboss', 'w:imprint', 'w:noProof', 'w:snapToGrid', 'w:vanish', 'w:webHidden', 'w:color', 'w:spacing', 'w:w', 'w:kern', 'w:position', 'w:sz', 'w:szCs']; const my = order.indexOf('w:' + nm); const i = rp.children.findIndex(c => c.type === 'el' && (order.indexOf(c.name) > my || order.indexOf(c.name) < 0)); if (i < 0) rp.children.push(n); else rp.children.splice(i, 0, n); }
    }
    if (lk.jc && !el(pPr, 'w:jc')) setJc(p, attr(parseXml(lk.jc).children[0], 'w:val'));
    remove(ps); cellHeadings++;
  }
}
console.log(`heading-styled table cells taken out of the TOC: ${cellHeadings}`);

// 8) Thai text tagged as complex script everywhere (body, tables, headings)
if (cfg.tag_thai_text !== false) tagThaiRuns(body);
console.log(`thai-distributed: body paragraphs=${bodyThai} table cells=${toThai} | soft line breaks joined=${brJoined} | cell paragraphs merged=${parasMerged} | thai runs tagged=${csTagged} (split=${runsSplit})`);

doc.set('word/document.xml', serialize(root));
saveDocx(doc, outF);
console.log(`nested tables lifted=${unnested} adjacent tables merged=${merged} header rows set=${hdr} cells un-distributed=${aligned} runs font-unified=${fonted}`);
