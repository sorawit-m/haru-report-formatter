// แทน add_pagenum.ps1 (ไม่ต้องเปิด Word): ทุกรายงานต้องมีเลขหน้า — ถ้าไม่มีฟิลด์ PAGE ใน footer ใดเลย
// ใส่เลขหน้าชิดขวา 16pt TH SarabunPSK ที่ section สุดท้าย (เนื้อหา) เริ่ม 1 ; ปก/สารบัญไม่มีเลข ; รันฟอนต์เดียว (one_font) ก่อนตามต้นฉบับ
// สารบัญอัปเดตตอนส่งออก PDF (scripts/mac/word.sh)
// usage: node add_pagenum.mjs <file.docx>
import { execFileSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const f = process.argv[2]; const HERE = path.dirname(fileURLToPath(import.meta.url));
execFileSync(process.execPath, [path.join(HERE, 'one_font.mjs'), f], { stdio: 'ignore' });
const d = loadDocx(f);
const footers = d.entries.filter(e => /^word\/footer\d*\.xml$/.test(e.name));
const hasPage = footers.some(e => /<w:fldSimple [^>]*w:instr="\s*PAGE\b|<w:instrText[^>]*>\s*PAGE\b/.test(d.get(e.name)));
if (hasPage) { console.log('page numbers already present'); process.exit(0); }

let x = d.get('word/document.xml');
const nSect = (x.match(/<w:sectPr\b/g) || []).length;
// footer ใหม่
let k = 1; while (d.get(`word/footer${k}.xml`) !== null) k++;
const part = `footer${k}.xml`;
const root = (x.match(/<w:document\b[^>]*>/) || [''])[0];
const ns = [...root.matchAll(/ (xmlns:\w+)="([^"]+)"/g)].map(m => ` ${m[1]}="${m[2]}"`).join('') || ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const F = 'TH SarabunPSK', rPr = `<w:rPr><w:rFonts w:ascii="${F}" w:eastAsia="${F}" w:hAnsi="${F}" w:cs="${F}"/><w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>`;
d.set(`word/${part}`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:ftr${ns}><w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="right"/>${rPr}</w:pPr>` +
  `<w:r>${rPr}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rPr}<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r>${rPr}<w:fldChar w:fldCharType="separate"/></w:r><w:r>${rPr}<w:t>1</w:t></w:r><w:r>${rPr}<w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>`);
// relationship + content type
let rels = d.get('word/_rels/document.xml.rels'); let id = 1; while (rels.includes(`Id="rId${id}"`)) id++;
rels = rels.replace('</Relationships>', `<Relationship Id="rId${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="${part}"/></Relationships>`);
d.set('word/_rels/document.xml.rels', rels);
let ct = d.get('[Content_Types].xml');
if (!ct.includes(`/word/${part}`)) ct = ct.replace('</Types>', `<Override PartName="/word/${part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>`);
d.set('[Content_Types].xml', ct);
// section สุดท้าย = sectPr ตัวสุดท้ายใน body
const at = x.lastIndexOf('<w:sectPr');
const end = x.indexOf('</w:sectPr>', at) + '</w:sectPr>'.length;
let sp = x.slice(at, end);
sp = sp.replace(/<w:footerReference [^>]*w:type="default"[^>]*\/>/, '');
// footerReference ต้องอยู่หลัง headerReference และก่อนลูกอื่น
const open = sp.match(/^<w:sectPr\b[^>]*>/)[0]; let body = sp.slice(open.length);
const lastRef = [...body.matchAll(/<w:(?:header|footer)Reference [^>]*\/>/g)].pop();
const ref = `<w:footerReference w:type="default" r:id="rId${id}"/>`;
body = lastRef ? body.slice(0, lastRef.index + lastRef[0].length) + ref + body.slice(lastRef.index + lastRef[0].length) : ref + body;
// เริ่มเลข 1 ที่ section นี้ (ถ้ามีหลาย section)
if (nSect > 1) {
  if (/<w:pgNumType\b/.test(body)) body = body.replace(/<w:pgNumType\b([^>]*?)\/>/, (m, a) => `<w:pgNumType${a.replace(/ w:start="[^"]*"/, '')} w:start="1"/>`);
  else {
    // ลำดับใน sectPr: … pgMar, paperSrc, pgBorders, lnNumType, pgNumType, cols, formProt, vAlign, noEndnote, titlePg, textDirection, bidi, rtlGutter, docGrid …
    const after = body.search(/<w:(?:cols|formProt|vAlign|noEndnote|titlePg|textDirection|bidi|rtlGutter|docGrid|printerSettings|sectPrChange)\b/);
    const pos = after >= 0 ? after : body.lastIndexOf('</w:sectPr>');
    body = body.slice(0, pos) + '<w:pgNumType w:start="1"/>' + body.slice(pos);
  }
}
x = x.slice(0, at) + open + body + x.slice(end);
if (!/xmlns:r=/.test(root)) x = x.replace(root, root.replace('<w:document', '<w:document xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'));
d.set('word/document.xml', x); saveDocx(d, f);
console.log(`page numbers added to section ${nSect} (start 1)`);
