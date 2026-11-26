// thaiwrap.mjs — จำลองการตัดบรรทัดภาษาไทยในช่องตาราง เพื่อบอกว่า "กระจายแบบไทย" จะทำให้ตัวอักษรถ่างไหม
// ใช้ความกว้างตัวอักษรจริงจากไฟล์ฟอนต์ TH SarabunPSK + ตัดคำไทยด้วย Intl.Segmenter (ICU) แบบเดียวกับที่ Word ตัดคำ
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const FONT_FILES = ['THSarabun.ttf', 'THSarabunNew.ttf', 'THSarabunPSK.ttf'];
function findFont() {
  const dirs = [path.join(process.env.LOCALAPPDATA || '', 'Microsoft', 'Windows', 'Fonts'), 'C:\\Windows\\Fonts',
    path.join(os.homedir(), 'Library', 'Fonts'), '/Library/Fonts', '/usr/share/fonts', '/usr/share/fonts/truetype'];
  for (const d of dirs) for (const f of FONT_FILES) { const p = path.join(d, f); if (fs.existsSync(p)) return p; }
  return null;
}
// minimal TrueType reader: unitsPerEm + advance widths via cmap format 4
function loadMetrics(file) {
  const b = fs.readFileSync(file);
  const n = b.readUInt16BE(4); const tables = {};
  for (let i = 0; i < n; i++) { const o = 12 + i * 16; tables[b.toString('ascii', o, o + 4)] = b.readUInt32BE(o + 8); }
  const upem = b.readUInt16BE(tables.head + 18);
  const nHM = b.readUInt16BE(tables.hhea + 34);
  const adv = gid => b.readUInt16BE(tables.hmtx + 4 * Math.min(gid, nHM - 1));
  const cm = tables.cmap; const nt = b.readUInt16BE(cm + 2); let sub = null;
  for (let i = 0; i < nt; i++) { const pid = b.readUInt16BE(cm + 4 + i * 8), eid = b.readUInt16BE(cm + 6 + i * 8), off = b.readUInt32BE(cm + 8 + i * 8); if (pid === 3 && eid === 1 && b.readUInt16BE(cm + off) === 4) sub = cm + off; }
  if (!sub) return null;
  const segX2 = b.readUInt16BE(sub + 6), ends = sub + 14, starts = ends + segX2 + 2, deltas = starts + segX2, ranges = deltas + segX2;
  const cache = new Map();
  function gid(c) {
    for (let i = 0; i < segX2 / 2; i++) {
      const end = b.readUInt16BE(ends + i * 2); if (c > end) continue;
      const start = b.readUInt16BE(starts + i * 2); if (c < start) return 0;
      const delta = b.readInt16BE(deltas + i * 2), ro = b.readUInt16BE(ranges + i * 2);
      if (!ro) return (c + delta) & 0xffff;
      const g = b.readUInt16BE(ranges + i * 2 + ro + (c - start) * 2); return g ? (g + delta) & 0xffff : 0;
    }
    return 0;
  }
  return ch => { const c = ch.codePointAt(0); if (cache.has(c)) return cache.get(c); const w = adv(gid(c)) / upem; cache.set(c, w); return w; };
}
const fontFile = findFont();
const width = fontFile ? loadMetrics(fontFile) : null;
export const available = !!width;
const MARK = /[\u0E31\u0E34-\u0E3A\u0E47-\u0E4E\u200B]/;      // Thai above/below marks + zero-width break: no advance
const charW = ch => (MARK.test(ch) ? 0 : (width ? width(ch) : 0.5));
const textW = s => [...s].reduce((a, ch) => a + charW(ch), 0);
const seg = new Intl.Segmenter('th', { granularity: 'word' });

// lines a text breaks into at a given width (em units), Word-like greedy fill; returns [{used, spread}] for every line
// Word's Thai breaker keeps compounds ICU splits ("ใบ|หน้า", "ค้น|หา"): a space-free chunk up to ~3 em is one unbreakable unit
const ATOMIC_EM = 3.0;
function splitParts(text) {
  const out = [];
  for (const piece0 of text.split(/(\s+)/)) {
    if (!piece0) continue;
    // U+200B (Word "No-Width Optional Break") = an allowed break inside a name
    const subs = /​/.test(piece0) ? piece0.split(/(?<=​)/) : [piece0];
    for (const piece of subs) {
    if (/^\s+$/.test(piece) || textW(piece) <= ATOMIC_EM || /^[\x21-\x7E​]+$/.test(piece)) { out.push(piece); continue; }
    for (const x of seg.segment(piece)) out.push(x.segment);
    }
  }
  return out;
}
export function wrap(text, widthEm) {
  const parts = splitParts(text);
  const lines = []; let cur = '', curW = 0;
  const push = () => { const t = cur.replace(/\s+$/, ''); lines.push({ text: t, used: textW(t) }); cur = ''; curW = 0; };
  for (const p of parts) {
    const w = textW(p);
    if (curW + w <= widthEm || /^\s+$/.test(p)) { cur += p; curW += w; continue; }
    if (cur.trim()) push();
    if (w <= widthEm) { cur = p.replace(/^\s+/, ''); curW = textW(cur); continue; }
    for (const ch of p) { const cw = charW(ch); if (curW + cw > widthEm && cur) push(); cur += ch; curW += cw; }   // longer than a line
  }
  if (cur.trim()) push();
  return lines;
}
// biggest per-character gap (em) Thai Distributed would add on any justified line, and the widest single word (em)
export function maxGap(text, widthEm) {
  const lines = wrap(text, widthEm); let worst = 0;
  for (let i = 0; i < lines.length - 1; i++) {
    const l = lines[i]; const slack = widthEm - l.used; if (slack <= 0) continue;
    if (!/[฀-๿]/.test(l.text)) {
      // English-only line: Word widens only the word gaps; "severity   level   not" looks odd once each gap grows by
      // more than ~0.35 em -> scale so that 0.35 em extra per gap == 0.12 (same threshold as Thai letters)
      const spaces = (l.text.match(/ /g) || []).length; if (!spaces) continue;
      worst = Math.max(worst, slack / spaces * 0.34); continue;
    }
    const n = (l.text.match(/[ก-ะาำเ-ๆ๏-๛\s]/g) || []).length;
    worst = Math.max(worst, n ? slack / n : slack);
  }
  return worst;
}
export function chunkWidth(text) { return textW(text.trim()); }
// widest piece a column must hold on one line: Latin/number tokens ("API-001", "Method", "VARCHAR(255)") are kept whole
// (splitting them at "-" breaks codes into "API-/001"); Thai is measured per word. +12% for bold header text.
export function longestUnit(text) {
  let m = 0;
  for (const tok of text.split(/\s+/)) {
    if (!tok) continue;
    if (/[฀-๿]/.test(tok)) for (const x of seg.segment(tok)) m = Math.max(m, textW(x.segment));
    else for (const part of tok.split('​')) m = Math.max(m, textW(part));
  }
  return m * 1.12;
}
export function longestWord(text) { return Math.max(0, ...[...seg.segment(text)].map(x => textW(x.segment.trim()))); }
// would Thai Distributed visibly stretch any line (all lines except the last are justified)?
// gap per stretchable character (Thai base letters + spaces) above ~0.11 em looks like "แ จ้ ง"
export function stretches(text, widthEm, maxGapEm = 0.11) {
  const lines = wrap(text, widthEm);
  for (let i = 0; i < lines.length - 1; i++) {
    const l = lines[i]; const slack = widthEm - l.used; if (slack <= 0) continue;
    const n = (l.text.match(/[\u0E01-\u0E30\u0E32\u0E33\u0E40-\u0E46\u0E4F-\u0E5B\s]/g) || []).length;
    if (!n || slack / n > maxGapEm) return true;
  }
  return false;
}
