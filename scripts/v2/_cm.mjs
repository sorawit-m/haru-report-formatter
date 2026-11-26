import { loadDocx } from '../lib/docx.mjs';
for (const f of process.argv.slice(2)) { const x = loadDocx(f).get('word/document.xml'); const txt = [...x.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').replace(/\u200B/g, '');
  const by = (txt.match(/Test(?:ed)? ?by|Approve(?:d)? ?by/gi) || []).length; const cap = (txt.match(/(?:รูปที่|ตารางที่)\d/g) || []).length; const cap2 = (txt.match(/(?:รูปที่|ตารางที่) \d/g) || []).length;
  console.log(f.split('\\').pop().slice(0, 45), { testedBy_approvedBy: by, captionNoSpace: cap, captionWithSpace: cap2, images: (x.match(/<w:drawing>/g) || []).length }); }
