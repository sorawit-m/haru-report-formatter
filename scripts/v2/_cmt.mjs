import { loadDocx } from '../lib/docx.mjs';
for (const f of process.argv.slice(2)) { const x = loadDocx(f).get('word/document.xml'); const txt = [...x.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').replace(/\u200B/g, '');
  const tby = (txt.match(/Test(ed)? ?by|Approve(d)? ?by/gi) || []).length;
  const capNo = (txt.match(/(รูปที่|ตารางที่)\d/g) || []).length, capSp = (txt.match(/(รูปที่|ตารางที่) \d/g) || []).length;
  console.log(f.split('\\').pop(), { testByApproveBy: tby, captionNoSpace: capNo, captionWithSpace: capSp, tables: (x.match(/<w:tbl>/g) || []).length, images: (x.match(/<w:drawing>/g) || []).length });
  const ex = txt.match(/.{0,30}(Test(ed)? ?by|Approve(d)? ?by).{0,40}/i); if (ex) console.log('   e.g.', ex[0]); }
