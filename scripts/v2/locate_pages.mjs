// แทน locate_pages.ps1 บน Mac: เติมเลขหน้าให้ตัวอย่างในรายงานการเปลี่ยนแปลง (changes.json จาก change_report.mjs)
// อ่านข้อความแต่ละหน้าจาก PDF ที่ส่งออกแล้ว ; เริ่มหาหลังหน้าสารบัญ (ไม่งั้นหัวข้อไปเจอในสารบัญ)
// usage: node locate_pages.mjs <file.pdf> <changes.json>
import fs from 'node:fs';
import { pdfPages, pageIndex } from '../lib/macpdf.mjs';
const [pdf, json] = process.argv.slice(2);
const cats = JSON.parse(fs.readFileSync(json, 'utf8'));
const pages = pdfPages(pdf); const idx = pageIndex(pages);
// หน้าสารบัญ: ในช่วง 15 หน้าแรก หน้าที่มีคำว่า "สารบัญ" หรือมีบรรทัดลงท้ายด้วยเลขหน้าเกินครึ่ง
let tocEnd = 0; pages.slice(0, 15).forEach((t, i) => { const lines = t.split('\n').filter(l => l.trim()); const nums = lines.filter(l => /\d+\s*$/.test(l)).length; if (/สารบัญ/.test(t) || (lines.length > 8 && nums / lines.length > 0.6)) tocEnd = i + 1; });
const from = tocEnd < idx.starts.length ? idx.starts[tocEnd] : 0;
const printedOf = n => { const ls = (pages[n - 1] || '').split('\n').map(s => s.trim()).filter(Boolean); for (const l of [ls[ls.length - 1], ls[0]]) { const m = (l || '').match(/^(?:หน้า\s*)?(\d{1,4})$/); if (m) return m[1]; } return ''; };
let found = 0;
for (const c of cats) for (const e of c.examples || []) {
  const key = String(e.find || ''); if (!key) continue;
  const h = idx.find(key, from) || idx.find(key, 0); if (!h) continue;
  e.page = h.page; e.printed = printedOf(h.page); found++;
}
fs.writeFileSync(json, JSON.stringify(cats, null, 1), 'utf8');
console.log(`pages located: ${found}`);
