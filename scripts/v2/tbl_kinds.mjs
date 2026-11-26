// สรุปชนิดตาราง (ตามหัวตาราง): จำนวน, ความกว้างคอลัมน์ (cm) แบบที่พบ, ความกว้างรวมเทียบความกว้างข้อความ
import { loadDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const x = loadDocx(f).get('word/document.xml');
const sp = [...x.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map(m => m[0]).pop();
const TW = +sp.match(/<w:pgSz [^>]*w:w="(\d+)"/)[1] - +sp.match(/w:left="(-?\d+)"/)[1] - +sp.match(/w:right="(-?\d+)"/)[1];
const kinds = {};
for (const m of x.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)) {
  const t = m[0]; const r0 = (t.match(/<w:tr\b[\s\S]*?<\/w:tr>/) || [''])[0];
  const h = [...r0.matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(c => [...c[1].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').replace(/\u200B/g, '').trim().slice(0, 14)).join('|');
  const g = [...(t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].matchAll(/w:w="(\d+)"/g)].map(a => +a[1]);
  const ind = +((t.match(/<w:tblInd w:w="(-?\d+)"/) || [])[1] || 0);
  const key = h; (kinds[key] ||= { n: 0, w: {} }).n++;
  const ws = g.map(v => (v / 567).toFixed(1)).join('/') + ` =${(g.reduce((a, b) => a + b, 0) / 567).toFixed(1)}${ind ? ' ind' + (ind / 567).toFixed(1) : ''}`;
  kinds[key].w[ws] = (kinds[key].w[ws] || 0) + 1;
}
console.log(`text width ${(TW / 567).toFixed(2)} cm`);
for (const [k, v] of Object.entries(kinds).sort((a, b) => b[1].n - a[1].n)) console.log(`${v.n}x [${k.slice(0, 70)}]  ${Object.entries(v.w).map(([w, c]) => `${w}${c > 1 ? ' x' + c : ''}`).join(' ; ').slice(0, 200)}`);
