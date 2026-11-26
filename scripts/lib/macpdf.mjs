// ตัวช่วยบน Mac: ส่งออก PDF ด้วย Word (scripts/mac/word.sh) และอ่านข้อความแต่ละหน้า (scripts/mac/pdftool.sh)
import { execFileSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const MAC = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'mac');
const run = (sh, args) => execFileSync('/bin/bash', [path.join(MAC, sh), ...args], { encoding: 'utf8', maxBuffer: 1 << 28 });

// อัปเดตสารบัญ + ส่งออก PDF -> { pages, note }
export function wordExport(docx, pdf) {
  const out = run('word.sh', pdf ? [docx, pdf] : [docx]);
  return { pages: +((out.match(/pages (\d+)/) || [])[1] || 0), note: (out.match(/^note: .*$/m) || [''])[0], out };
}
export function pdfToPng(pdf, outDir, width = 1100, from = 1, to = 99999) { return run('pdftool.sh', ['png', pdf, outDir, String(width), String(from), String(to)]); }
export function pdfCheck(pdf, outDir) { return run('pdftool.sh', ['check', pdf, outDir]); }
export function pdfPages(pdf) { return JSON.parse(run('pdftool.sh', ['text', pdf])); }

// ตัวค้นหาตำแหน่งข้อความ -> เลขหน้า (เทียบแบบตัดช่องว่าง/อักขระล่องหน)
const N = s => s.replace(/[\s​⁠ ­]+/g, '');
export function pageIndex(pages) {
  const starts = []; let all = '';
  pages.forEach(t => { starts.push(all.length); all += N(t); });
  const pageAt = pos => { let lo = 0, hi = starts.length - 1; while (lo < hi) { const m = (lo + hi + 1) >> 1; if (starts[m] <= pos) lo = m; else hi = m - 1; } return lo + 1; };
  return {
    all, starts, pageAt,
    // หา key ตั้งแต่ตำแหน่ง from (ตัวอักษรใน all) — คืน { pos, page } หรือ null
    find(key, from = 0, maxLen = 24) { const k = N(key).slice(0, maxLen); if (k.length < 2) return null; const p = all.indexOf(k, from); return p < 0 ? null : { pos: p, page: pageAt(p), len: k.length }; },
  };
}
