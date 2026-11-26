// format-docx.mjs — จัดรูปแบบเล่มรายงานโดยแก้ XML ตรง (ใช้ได้ทั้ง Mac และ Windows, ไม่ต้องมี Word)
// ใช้: node format-docx.mjs <in.docx> <out.docx> <config.json> [replacements.tsv] [report.txt]
// ทำตาราง (รวมตารางหัว/เนื้อหา ฯลฯ) ด้วย fix-tables.mjs ก่อนไฟล์นี้
// ข้อจำกัด: ไม่มีระบบจัดหน้าของ Word -> เรื่องที่ต้องรู้ "หน้า" (หน้าว่าง, ตารางสูงเกินหน้า, เลขหน้าสารบัญ)
//           ให้ Word อัปเดตเองตอนเปิดไฟล์ (updateFields) และตรวจด้วยตาจาก PDF
import fs from 'node:fs';
import { loadDocx, saveDocx, parseXml, serialize, els, el, all, attr, setAttr, mk, remove, insertBefore, textOf, ensure } from './lib/docx.mjs';

const [, , inF, outF, cfgF, repF, reportF] = process.argv;
if (!inF || !outF || !cfgF) { console.error('usage: node format-docx.mjs in.docx out.docx config.json [replacements.tsv] [report.txt]'); process.exit(1); }
const cfg = JSON.parse(fs.readFileSync(cfgF, 'utf8'));
const log = []; const L = m => { console.log(m); log.push(m); };
const CODE = new RegExp(cfg.code_fonts || 'Consol|Courier|Mono', 'i');
const DIST = /^(thaiDistribute|distribute|both)$/;
const doc = loadDocx(inF);

// ---------- 1. plain-text level edits in every word/*.xml ----------
for (const e of doc.entries) {
  if (!/^word\/.*\.xml$/.test(e.name)) continue;
  let s = e.data.toString('utf8'); const before = s;
  for (const [from, to] of Object.entries(cfg.font_replace || {})) s = s.split(`"${from}"`).join(`"${to}"`);
  if (cfg.thin_border_to_black) {
    // faint borders (0.125/0.25 pt, or a light colour) look "missing" -> black; keep colours in keep_border_colors (e.g. code boxes)
    const bsz = Math.round((cfg.table_border_pt || 0.5) * 8); const keepC = (cfg.keep_border_colors || []).map(c => c.toUpperCase());
    s = s.replace(/<w:(top|left|bottom|right|insideH|insideV|start|end) w:val="single" w:sz="(\d+)" w:space="(\d+)" w:color="([0-9A-Fa-f]{6})"/g, (m, side, sz, sp, col) => {
      const C = col.toUpperCase(); const r = parseInt(C.slice(0, 2), 16), g = parseInt(C.slice(2, 4), 16), b = parseInt(C.slice(4, 6), 16);
      const light = 0.299 * r + 0.587 * g + 0.114 * b > 140;
      if (keepC.includes(C) || (!light && +sz >= 4)) return m;
      return `<w:${side} w:val="single" w:sz="${Math.max(+sz, bsz)}" w:space="${sp}" w:color="000000"`;
    });
  }
  if (cfg.no_italic) s = s.replace(/<w:i(?:Cs)?(?: w:val="(?:1|true|on)")?\/>/g, '');
  if (s !== before) e.data = Buffer.from(s, 'utf8');
}
L('fonts / thin borders / italics handled in XML');

// ---------- styles: which styles are headings, which are distributed ----------
const stylesXml = doc.get('word/styles.xml') || '';
const sOutline = {}, sJc = {}, sBase = {}; let defStyle = null;
for (const m of stylesXml.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
  const id = (m[1].match(/w:styleId="([^"]*)"/) || [])[1]; if (!id) continue;
  if (/w:type="paragraph"/.test(m[1]) && /w:default="1"/.test(m[1])) defStyle = id;
  const o = (m[2].match(/<w:outlineLvl w:val="(\d+)"/) || [])[1]; if (o !== undefined) sOutline[id] = +o;
  const j = (m[2].match(/<w:jc w:val="([^"]*)"/) || [])[1]; if (j) sJc[id] = j;
  const b = (m[2].match(/<w:basedOn w:val="([^"]*)"/) || [])[1]; if (b) sBase[id] = b;
}
const resolve = (map, id) => { for (let k = 0; id && k < 10; k++, id = sBase[id]) if (map[id] !== undefined) return map[id]; return undefined; };
// headings keep with next + keep lines (in styles.xml)
if (stylesXml) {
  const ns = stylesXml.replace(/(<w:style\b[^>]*w:type="paragraph"[^>]*>[\s\S]*?<w:pPr>)([\s\S]*?)(<\/w:pPr>)/g, (all0, a, body, c) => {
    if (!/<w:outlineLvl w:val="\d"/.test(body)) return all0;
    let b = body; if (!/<w:keepNext\/>/.test(b)) b = '<w:keepNext/>' + b; if (!/<w:keepLines\/>/.test(b)) b = b.replace('<w:keepNext/>', '<w:keepNext/><w:keepLines/>');
    return a + b + c;
  });
  doc.set('word/styles.xml', ns);
}

// ---------- document.xml ----------
const xml = doc.get('word/document.xml'); const root = parseXml(xml);
if (serialize(root) !== xml) { console.error('XML round-trip mismatch — stop'); process.exit(2); }
const body = el(el(root, 'w:document'), 'w:body');
const pStyleOf = p => { const pPr = el(p, 'w:pPr'); const ps = pPr && el(pPr, 'w:pStyle'); return ps ? attr(ps, 'w:val') : defStyle; };
const outlineOf = p => { const pPr = el(p, 'w:pPr'); const o = pPr && el(pPr, 'w:outlineLvl'); return o ? +attr(o, 'w:val') : resolve(sOutline, pStyleOf(p)); };
const jcOf = p => { const pPr = el(p, 'w:pPr'); const j = pPr && el(pPr, 'w:jc'); return j ? attr(j, 'w:val') : resolve(sJc, pStyleOf(p)); };
const hasDrawing = n => all(n, 'w:drawing').length + all(n, 'w:pict').length > 0;
const isEmptyP = p => p.name === 'w:p' && !textOf(p).trim() && !hasDrawing(p) && !all(p, 'w:sectPr').length && !all(p, 'w:br').some(b => attr(b, 'w:type') === 'page') && !all(p, 'w:fldChar').length;
const PPR_ORDER = ['w:pStyle', 'w:keepNext', 'w:keepLines', 'w:pageBreakBefore', 'w:framePr', 'w:widowControl', 'w:numPr', 'w:suppressLineNumbers', 'w:pBdr', 'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap', 'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN', 'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind', 'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc', 'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap', 'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr', 'w:pPrChange'];
function setPPr(p, name, attrs = []) {
  const pPr = ensure(p, 'w:pPr'); let c = el(pPr, name);
  if (c) { attrs.forEach(([k, v]) => setAttr(c, k, v)); return; }
  c = mk(name, attrs); c.parent = pPr; const my = PPR_ORDER.indexOf(name);
  const idx = pPr.children.findIndex(x => x.type === 'el' && PPR_ORDER.indexOf(x.name) > my);
  if (idx < 0) pPr.children.push(c); else pPr.children.splice(idx, 0, c);
}
const bodyKids = () => els(body);
let firstSect = bodyKids().findIndex(n => n.name === 'w:p' && all(n, 'w:sectPr').length);
const coverEnd = firstSect < 0 ? 0 : firstSect;

// 2. replacements (works across runs inside one paragraph)
if (repF && fs.existsSync(repF)) {
  const pairs = fs.readFileSync(repF, 'utf8').replace(/^﻿/, '').split(/\r?\n/).filter(l => l.trim() && !l.startsWith('#')).map(l => l.split('\t')).filter(p => p.length >= 2);
  const paras = all(root, 'w:p'); let n = 0; const miss = [];
  for (const [from, to] of pairs) {
    let hit = false;
    for (const p of paras) {
      const ts = all(p, 'w:t', [], 'w:p'); if (!ts.length) continue;
      const parts = ts.map(t => t.children.map(c => c.text || '').join('')); let full = parts.join('');
      let pos = full.indexOf(from);
      while (pos >= 0) {
        hit = true; let acc = 0, done = false;
        for (let k = 0; k < ts.length; k++) {
          const s = acc, e = acc + parts[k].length; acc = e;
          if (e <= pos || s >= pos + from.length) continue;
          const a = Math.max(0, pos - s), b = Math.min(parts[k].length, pos + from.length - s);
          parts[k] = parts[k].slice(0, a) + (done ? '' : to) + parts[k].slice(b); done = true;
        }
        ts.forEach((t, k) => { t.children = [{ type: 'text', text: parts[k] }]; if (/^\s|\s$/.test(parts[k])) setAttr(t, 'xml:space', 'preserve'); });
        full = parts.join(''); pos = full.indexOf(from, pos + to.length);
      }
    }
    if (hit) n++; else miss.push(from);
  }
  L(`typo replacements applied: ${n} (not found: ${miss.length})`);
}

// 3. stray manual line breaks inside distributed/justified paragraphs
if (cfg.remove_stray_line_breaks) {
  let n = 0;
  for (const p of all(root, 'w:p')) {
    if (!DIST.test(jcOf(p) || '')) continue;
    const runs = all(p, 'w:r', [], 'w:p');
    runs.forEach((r, ri) => {
      for (const br of els(r, 'w:br')) {
        const t = attr(br, 'w:type'); if (t && t !== 'textWrapping') continue;
        const rf = el(el(r, 'w:rPr') || { children: [] }, 'w:rFonts'); if (rf && CODE.test(attr(rf, 'w:ascii') || '')) continue;
        // next visible char
        const idx = r.children.indexOf(br); let nxt = '';
        for (const c of r.children.slice(idx + 1)) if (c.name === 'w:t') { nxt = textOf({ children: [c] }) || c.children.map(x => x.text).join(''); break; }
        if (!nxt) for (const r2 of runs.slice(ri + 1)) { const tt = all(r2, 'w:t').map(t => t.children.map(x => x.text).join('')).join(''); if (tt) { nxt = tt; break; } }
        if (nxt && !/^\s/.test(nxt)) { remove(br); n++; }
      }
    });
  }
  L(`manual line breaks removed: ${n}`);
}

// 4. empty paragraph runs (after the cover), empty bullets, cover bullet dot, duplicated paragraphs
{
  let collapsed = 0, eb = 0, cov = 0, dup = 0;
  const kids = bodyKids();
  for (let i = coverEnd + 1; i < kids.length; i++) {
    if (isEmptyP(kids[i]) && i > 0 && isEmptyP(kids[i - 1]) && cfg.max_empty_lines <= 1) { remove(kids[i]); collapsed++; }
  }
  if (cfg.remove_empty_bullets) for (const p of bodyKids()) { if (p.name === 'w:p' && el(el(p, 'w:pPr') || { children: [] }, 'w:numPr') && !textOf(p).trim() && !hasDrawing(p)) { remove(p); eb++; } }
  if (cfg.fix_cover_bullet_dot) bodyKids().slice(0, coverEnd + 1).forEach(p => { const pPr = el(p, 'w:pPr'); const np = pPr && el(pPr, 'w:numPr'); if (p.name === 'w:p' && hasDrawing(p) && (np || /bullet|list/i.test(pStyleOf(p) || ''))) { if (np) remove(np); const ps = el(pPr, 'w:pStyle'); if (ps) remove(ps); setPPr(p, 'w:jc', [['w:val', 'center']]); cov++; } });
  if (cfg.remove_duplicate_paragraphs) { const k2 = bodyKids(); for (let i = 0; i < k2.length; i++) { const p = k2[i]; if (p.name !== 'w:p') continue; const t = textOf(p).trim(); if (t.length < 80) continue; for (let j = Math.max(0, i - 8); j < i; j++) if (k2[j].name === 'w:p' && k2[j].parent && textOf(k2[j]).trim() === t) { remove(p); dup++; break; } } }
  L(`empty paragraph runs collapsed=${collapsed} empty bullets removed=${eb} cover bullet dots removed=${cov} duplicated paragraphs removed=${dup}`);
}

// 5. table captions "ตาราง N ..." -> "<label> N ..."
{
  let n = 0; const label = cfg.table_caption_label || 'ตารางที่';
  for (const p of bodyKids()) {
    if (p.name !== 'w:p') continue; const full = textOf(p).trimStart();
    if (!/^ตาราง\s*\d/.test(full) || full.startsWith(label)) continue;
    const t0 = all(p, 'w:t', [], 'w:p').find(t => /ตาราง/.test(t.children.map(c => c.text).join(''))); if (!t0) continue;
    const s = t0.children.map(c => c.text).join('');
    t0.children = [{ type: 'text', text: s.replace(/^(\s*)ตาราง(?!ที่)\s*/, `$1${label} `) }]; setAttr(t0, 'xml:space', 'preserve'); n++;
  }
  L(`table captions relabelled: ${n}`);
}

// 6. tables: header rows repeat, rows do not split, keep rules, borders, padding, leading spaces
{
  let hdr = 0, keep = 0, bord = 0, pad = 0, lead = 0;
  const shaded = tr => all(tr, 'w:shd').some(s => { const f = (attr(s, 'w:fill') || '').toUpperCase(); return f && f !== 'AUTO' && f !== 'FFFFFF'; });
  const addTrPr = (tr, name) => { const trPr = ensure(tr, 'w:trPr', false); if (!el(trPr, name)) { const n = mk(name); n.parent = trPr; trPr.children.push(n); } if (tr.children.indexOf(trPr) !== tr.children.findIndex(c => c.type === 'el' && c.name !== 'w:tblPrEx')) { tr.children.splice(tr.children.indexOf(trPr), 1); const i = tr.children.findIndex(c => c.type === 'el' && c.name === 'w:tc'); tr.children.splice(i < 0 ? tr.children.length : i, 0, trPr); } };
  for (const tbl of all(root, 'w:tbl')) {
    const trs = els(tbl, 'w:tr'); if (!trs.length) continue;
    const ncols = (el(tbl, 'w:tblGrid') ? els(el(tbl, 'w:tblGrid'), 'w:gridCol').length : 1);
    // header rows
    let h = 0;
    if (trs.length > 1 && shaded(trs[0]) && !shaded(trs[1])) h = 1;
    if (/^(Table Name|ชื่อตาราง)/.test(textOf(trs[0]).trim())) { h = 1; trs.slice(0, 4).forEach((tr, i) => { const t = textOf(el(tr, 'w:tc') || tr).trim(); if (/^(NO|No\.?|ลำดับ|#)$/.test(t) || /^Table /.test(t)) h = i + 1; }); }
    for (let i = 0; i < h; i++) { addTrPr(trs[i], 'w:tblHeader'); }
    if (h) hdr++;
    const len = textOf(tbl).length;
    trs.forEach(tr => { if (textOf(tr).length < 1800) addTrPr(tr, 'w:cantSplit'); });
    const picRow = trs.map(tr => hasDrawing(tr));
    const small = !picRow.some(Boolean) && trs.length <= (cfg.keep_whole_table_max_rows || 8) && len < (cfg.keep_whole_table_max_chars || 2000);
    const lim = small ? trs.length - 1 : Math.max(h, 1) + (cfg.keep_header_plus_rows || 2) - 1;
    // never chain a row to a following picture row (screenshot rows are too tall to share a page with the chain)
    trs.slice(0, lim).forEach((tr, i) => { if (picRow[i + 1]) return; all(tr, 'w:p', [], 'w:tbl').forEach(p => setPPr(p, 'w:keepNext')); });
    if (small) keep++;
    // caption / title paragraph right before the table
    const prev = tbl.parent.children[tbl.parent.children.indexOf(tbl) - 1];
    if (prev && prev.name === 'w:p' && textOf(prev).trim().length && textOf(prev).length < 250) setPPr(prev, 'w:keepNext');
    // padding
    const tblPr = ensure(tbl, 'w:tblPr');
    const mar = el(tblPr, 'w:tblCellMar'); const want = String(Math.round((cfg.table_cell_padding_pt || 5.4) * 20));
    const cur = mar && el(mar, 'w:left'); if (!cur || (+attr(cur, 'w:w') || 0) < 80) {
      const m2 = mar || (() => { const n = mk('w:tblCellMar'); n.parent = tblPr; tblPr.children.push(n); return n; })();
      for (const side of ['w:left', 'w:right']) { let s = el(m2, side); if (!s) { s = mk(side); s.parent = m2; m2.children.push(s); } setAttr(s, 'w:w', want); setAttr(s, 'w:type', 'dxa'); }
      pad++;
    }
    // borders for borderless multi-column tables
    if (ncols >= 2 && !el(tblPr, 'w:tblBorders') && !el(tblPr, 'w:tblStyle') && !all(tbl, 'w:tcBorders').length) {
      const sz = String(Math.round((cfg.table_border_pt || 0.5) * 8));
      const b = mk('w:tblBorders', [], ['w:top', 'w:left', 'w:bottom', 'w:right', 'w:insideH', 'w:insideV'].map(n => mk(n, [['w:val', 'single'], ['w:sz', sz], ['w:space', '0'], ['w:color', '000000']])));
      b.parent = tblPr; const order = ['w:tblStyle', 'w:tblpPr', 'w:tblOverlap', 'w:bidiVisual', 'w:tblStyleRowBandSize', 'w:tblStyleColBandSize', 'w:tblW', 'w:jc', 'w:tblCellSpacing', 'w:tblInd'];
      let idx = 0; tblPr.children.forEach((c, i) => { if (c.type === 'el' && order.includes(c.name)) idx = i + 1; }); tblPr.children.splice(idx, 0, b); bord++;
    }
    // leading spaces in cells
    if (ncols >= 2) for (const tc of all(tbl, 'w:tc')) { const t = all(tc, 'w:t')[0]; if (!t) continue; const s = t.children.map(c => c.text).join(''); if (/^[  ]+/.test(s)) { t.children = [{ type: 'text', text: s.replace(/^[  ]+/, '') }]; lead++; } }
  }
  L(`tables: header repeat=${hdr} kept whole=${keep} borders added=${bord} padding fixed=${pad} cells leading spaces=${lead}`);
}

// 7. figures: picture keeps with caption; heading + short intro keep with picture; caption in same paragraph -> split
{
  let chained = 0, split = 0;
  const kids = bodyKids();
  for (const p of kids) {
    if (p.name !== 'w:p' || !hasDrawing(p)) continue;
    // split "รูปที่ ..." caption that sits in the picture paragraph
    const runs = els(p, 'w:r'); const di = runs.findIndex(r => hasDrawing(r));
    const tail = runs.slice(di + 1); const tailText = tail.map(r => textOf(r)).join('');
    if (di >= 0 && /^\s*รูปที่/.test(tailText)) {
      const np = mk('w:p'); const pPr = el(p, 'w:pPr'); if (pPr) { const c = parseXml(serialize(pPr)).children[0]; c.parent = np; np.children.push(c); }
      tail.forEach(r => { remove(r); r.parent = np; np.children.push(r); });
      const i = p.parent.children.indexOf(p); p.parent.children.splice(i + 1, 0, np); np.parent = p.parent; setPPr(np, 'w:jc', [['w:val', 'center']]); split++;
    }
    setPPr(p, 'w:keepNext');
    let acc = 0, i = p.parent.children.indexOf(p) - 1;
    for (let k = 0; k < 10 && i >= 0; k++, i--) {
      const q = p.parent.children[i]; if (q.type !== 'el') { k--; continue; }
      if (q.name !== 'w:p' || hasDrawing(q) || /^(รูปที่|ตารางที่)/.test(textOf(q).trim())) break;
      acc += textOf(q).length; if (acc > (cfg.figure_intro_max_chars || 900)) break;
      setPPr(q, 'w:keepNext'); setPPr(q, 'w:keepLines');
      if (outlineOf(q) !== undefined && outlineOf(q) < 9) { chained++; break; }
    }
  }
  L(`figure captions split=${split} figures chained to heading=${chained}`);
}

// 8. chapter (outline level 0) and front-matter headings start on a new page
if (cfg.chapter_new_page) {
  let n = 0; const fronts = new Set(cfg.front_headings || []);
  for (const p of bodyKids()) {
    if (p.name !== 'w:p') continue; const t = textOf(p).trim();
    if (!(outlineOf(p) === 0 || fronts.has(t))) continue;
    if (bodyKids().indexOf(p) <= coverEnd && !fronts.has(t)) continue;
    // drop empty / page-break-only paragraphs right before, and a page break inside the heading itself
    let prev = p.parent.children[p.parent.children.indexOf(p) - 1];
    while (prev && prev.name === 'w:p' && !all(prev, 'w:sectPr').length && !textOf(prev).trim() && !hasDrawing(prev)) { const pp = prev.parent.children[prev.parent.children.indexOf(prev) - 1]; remove(prev); prev = pp; }
    for (const br of all(p, 'w:br')) if (attr(br, 'w:type') === 'page') remove(br);
    setPPr(p, 'w:pageBreakBefore'); n++;
  }
  L(`chapter/front headings on new page: ${n}`);
}

doc.set('word/document.xml', serialize(root));

// 9. settings: don't stretch lines ending with Shift+Enter; ask Word to refresh TOC/fields on open
{
  let s = doc.get('word/settings.xml') || '';
  if (s && !/<w:updateFields/.test(s)) s = s.replace(/(<w:settings\b[^>]*>)/, '$1<w:updateFields w:val="true"/>');
  if (s && !/<w:doNotExpandShiftReturn/.test(s)) s = /<w:compat>/.test(s) ? s.replace('<w:compat>', '<w:compat><w:doNotExpandShiftReturn/>') : s.replace('</w:settings>', '<w:compat><w:doNotExpandShiftReturn/></w:compat></w:settings>');
  doc.set('word/settings.xml', s);
}

// 10. report: red text, italic, highlight, comments, tracked changes (with nearest heading)
{
  const out = []; let head = '(ก่อนหัวข้อแรก)';
  const red = c => { if (!c) return false; c = c.toUpperCase(); if (c === 'AUTO') return false; const r = parseInt(c.slice(0, 2), 16), g = parseInt(c.slice(2, 4), 16), b = parseInt(c.slice(4, 6), 16); return r > 150 && g < 90 && b < 90; };
  for (const p of all(root, 'w:p')) {
    const t = textOf(p).trim(); if (outlineOf(p) !== undefined && outlineOf(p) < 9 && t) head = t.slice(0, 60);
    for (const r of all(p, 'w:r', [], 'w:p')) {
      const rPr = el(r, 'w:rPr'); if (!rPr) continue; const c = el(rPr, 'w:color');
      const tx = textOf(r).trim(); if (!tx) continue;
      if (c && red(attr(c, 'w:val'))) out.push(`สีแดง | ${head} | ${tx.slice(0, 80)}`);
      if (el(rPr, 'w:highlight')) out.push(`ไฮไลต์ | ${head} | ${tx.slice(0, 80)}`);
      if (el(rPr, 'w:i') || el(rPr, 'w:iCs')) out.push(`ตัวเอียง | ${head} | ${tx.slice(0, 80)}`);
    }
  }
  const nIns = all(root, 'w:ins').length + all(root, 'w:del').length; if (nIns) out.push(`Track changes ค้างอยู่ ${nIns} จุด`);
  if (doc.entries.find(e => e.name === 'word/comments.xml')) out.push('มีคอมเมนต์ในไฟล์ (word/comments.xml)');
  L(`scan: ${out.length} items (red/highlight/italic/track changes)`);
  if (reportF) fs.writeFileSync(reportF, '﻿' + [...log, '', ...out].join('\r\n'), 'utf8');
}

saveDocx(doc, outF);
L(`saved ${outF}`);
