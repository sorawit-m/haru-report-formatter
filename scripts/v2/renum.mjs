// ลดเลขบทที่พิมพ์เอง (นอกตาราง) หลังลบบทนำ: "2.1.1 x" -> "1.1.1 x", "บทที่ 2" -> "บทที่ 1", "ตารางที่ 2-1" -> "ตารางที่ 1-1"
// --dry แสดงอย่างเดียว; --split "ข้อความหัวบท" = ตั้งแต่ย่อหน้านี้ไป ใช้ map แยก (--map2 "3:4")
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const a = process.argv.slice(2); const f = a[0]; const dry = a.includes('--dry');
const split = a.includes('--split') ? a[a.indexOf('--split') + 1] : null;
const map2 = a.includes('--map2') ? Object.fromEntries(a[a.indexOf('--map2') + 1].split(',').map(s => s.split(':').map(Number))) : null;
const d = loadDocx(f); let x = d.get('word/document.xml');
let depth = 0, after = false, n = 0; const out = [];
const PAT = [/^(\s*)(\d+)(?=(?:\.\d+)+\.?\s)/, /^(\s*(?:บทที่)\s*)(\d+)(?=[\s.\-–]|$)/];
const toc = /<w:pStyle w:val="(?:TOC\d|toc\d|1[3-9]|2[3-9])"/; // TOC lines are rebuilt by Word anyway
x = x.replace(/<w:tbl>|<\/w:tbl>|<w:p[ >][\s\S]*?<\/w:p>/g, p => {
  if (p === '<w:tbl>') { depth++; return p; } if (p === '</w:tbl>') { depth--; return p; }
  if (depth || /<w:instrText[^>]*>\s*(?:TOC|PAGEREF)/.test(p)) return p;
  const txt = [...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(m => m[1]).join('');
  if (split && !after && txt.includes(split)) after = true;
  const mp = after && map2 ? map2 : null;
  for (const re of PAT) {
    const m = txt.match(re); if (!m) continue;
    const old = +m[2]; const nw = mp ? (mp[old] ?? old) : (old >= 2 ? old - 1 : old); if (nw === old) return p;
    // the number must sit inside the first non-empty w:t
    let done = false;
    p = p.replace(/(<w:t(?: [^>]*)?>)([^<]*)(<\/w:t>)/, (all, o, t, c) => { if (t.trim() === '' ) return all; const mm = t.match(re); if (!mm) return all; done = true; return o + t.replace(re, `$1${nw}`) + c; });
    n++; out.push(`${done ? '' : '!! SKIP(split run) '}${txt.slice(0, 70)}  ->  ${old}->${nw}`);
    return p;
  }
  return p;
});
console.log(out.join('\n')); console.log(`total ${n}`);
if (!dry) { d.set('word/document.xml', x); saveDocx(d, f); }
