// บอกชนิดเล่ม: api (กล่องโค้ด/ตาราง Field) | testlist (ตารางทดสอบ 9 คอลัมน์ + รูปประกอบ) | testresult (ตารางทดสอบ + ผลที่พบ) | general
import { loadDocx } from '../lib/docx.mjs';
const x = loadDocx(process.argv[2]).get('word/document.xml');
const norm = s => s.replace(/[\s​]+/g, '').toLowerCase();
let code = 0, field = 0, test = 0, nine = 0;
for (const m of x.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)) {
  const t = m[0]; const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length;
  if (g === 1) { code++; continue; }
  const r0 = (t.match(/<w:tr\b[\s\S]*?<\/w:tr>/) || [''])[0]; const h = norm([...r0.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('|'));
  if (/^field\|type/.test(h)) field++; if (/^testid\|caseid/.test(h)) test++;
  else if ((g === 9 || g === 10) && /^[a-z]+_?\d{3}/.test(h.replace(/\|/g, ''))) nine++;   // ตารางข้อมูลทดสอบที่แยกจากหัวตาราง (หัวมีตารางเดียว)
}
if (test >= 1 && test < 5 && test + nine >= 3) test = 5;
const txt = x.replace(/<[^>]+>/g, '');
let kind = 'general';
if (test >= 5) kind = /ผลที่พบ|Actual Result/.test(txt) ? 'testresult' : 'testlist';
else if (code >= 10 || field >= 10) kind = 'api';
console.log(JSON.stringify({ kind, codeBoxes: code, fieldTables: field, testTables: test }));
