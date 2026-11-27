// ตารางสั้น (≤ keep_whole_table_max_rows แถว และ < keep_whole_table_max_chars ตัวอักษร ใน config.json) อยู่หน้าเดียวทั้งตาราง: แถวไม่แยกหน้า + ผูกทุกแถวกับแถวถัดไป (ยกเว้นแถวสุดท้าย)
// ข้ามตารางที่ครอบสารบัญ (ฟิลด์ TOC) และตารางที่มีรูป — ไม่งั้นสารบัญ/รูปทั้งก้อนหลุดไปหน้าถัดไป (27-11-2026)
import { loadDocx, saveDocx } from '../lib/docx.mjs';
import fs from 'node:fs'; import path from 'node:path'; import { fileURLToPath } from 'node:url';
const cfg = (() => { try { return JSON.parse(fs.readFileSync(path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..', 'config.json'), 'utf8').replace(/^﻿/, '')); } catch { return {}; } })();
const MAX_ROWS = cfg.keep_whole_table_max_rows || 8, MAX_CHARS = cfg.keep_whole_table_max_chars || 2000;
const addKeepNext = (p) => {
  if (/<w:keepNext\/>/.test(p)) return p;
  if (/<w:pPr>/.test(p)) return p.replace(/<w:pPr>(\s*<w:pStyle [^>]*\/>)?/, (m) => m + '<w:keepNext/>');
  return p.replace(/^(<w:p(?: [^>]*)?>)/, '$1<w:pPr><w:keepNext/></w:pPr>');
};
for (const f of process.argv.slice(2)) {
  const d = loadDocx(f); let x = d.get('word/document.xml'); let n = 0, skipped = 0;
  // only top-level tables (no nested): split on <w:tbl> boundaries
  x = x.replace(/<w:tbl>(?:(?!<w:tbl>)[\s\S])*?<\/w:tbl>/g, (t) => {
    const rows = t.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) || [];
    if (rows.length < 2 || rows.length > MAX_ROWS) return t;
    if ([...t.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].reduce((a, m) => a + m[1].length, 0) >= MAX_CHARS) return t;
    if (/instrText[^>]*>\s*TOC\b/.test(t) || /<w:drawing/.test(t)) { skipped++; return t; }
    n++;
    let i = 0;
    return t.replace(/<w:tr[ >][\s\S]*?<\/w:tr>/g, (r) => {
      i++;
      if (!/<w:cantSplit\/>/.test(r)) r = /<w:trPr>/.test(r) ? r.replace('<w:trPr>', '<w:trPr><w:cantSplit/>') : r.replace(/^(<w:tr(?: [^>]*)?>)/, '$1<w:trPr><w:cantSplit/></w:trPr>');
      if (i < rows.length) r = r.replace(/<w:p[ >][\s\S]*?<\/w:p>/g, addKeepNext);
      return r;
    });
  });
  d.set('word/document.xml', x); saveDocx(d, f);
  console.log(`${f.split(/[\\/]/).pop()}: short tables kept ${n}, skipped (TOC/picture) ${skipped}`);
}
