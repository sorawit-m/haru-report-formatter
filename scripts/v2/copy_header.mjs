// คัดลอกหัวกระดาษ (โลโก้ + ข้อความโครงการ + เส้น) จากเล่มต้นแบบ ไปแทนหัวกระดาษทุกอันในเล่มปลายทางที่มีข้อความโครงการ
// usage: node copy_header.mjs <source.docx> <target.docx>
import { loadDocx, saveDocx } from '../lib/docx.mjs';
const [srcF, dstF] = process.argv.slice(2);
const S = loadDocx(srcF), T = loadDocx(dstF);
const rels = (doc, part) => { const p = part.replace(/^word\//, 'word/_rels/') + '.rels'; return { path: p, xml: doc.get(p) || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>' }; };
// source: the default header of the last (body) section
const sdoc = S.get('word/document.xml'); const lastSect = [...sdoc.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].pop()[0];
const rid = (lastSect.match(/<w:headerReference [^>]*w:type="default"[^>]*r:id="([^"]+)"/) || lastSect.match(/<w:headerReference [^>]*r:id="([^"]+)"[^>]*w:type="default"/))[1];
const srels = S.get('word/_rels/document.xml.rels'); const sTarget = srels.match(new RegExp(`<Relationship [^>]*Id="${rid}"[^>]*Target="([^"]+)"`))[1];
const sPart = 'word/' + sTarget; const sXml = S.get(sPart); const sR = rels(S, sPart);
console.log('source header:', sPart);
// media used by the source header
const media = [...sR.xml.matchAll(/<Relationship [^>]*Id="([^"]+)"[^>]*Target="(media\/[^"]+)"/g)].map(m => ({ id: m[1], target: m[2] }));
// ใส่รูปแบบของสไตล์ (เช่น "หัวกระดาษ1": ชิดขวา 14pt) ลงในย่อหน้าตรงๆ แล้วเอา pStyle ออก
// เพราะเล่มปลายทางมักไม่มีสไตล์นั้น และ Word จะตัด pStyle ที่หาไม่เจอทิ้งตอนบันทึก -> หัวกระดาษกลายเป็นตัวใหญ่ไม่ชิดขวา (เจอในเล่ม 6 อัตโนมัติ)
const sStyles = S.get('word/styles.xml');
const styleOf = id => (sStyles.match(new RegExp(`<w:style\\b[^>]*w:styleId="${id}"[\\s\\S]*?</w:style>`)) || [])[0] || '';
const inner = (x, tag) => ((x.match(new RegExp(`<w:${tag}>([\\s\\S]*?)</w:${tag}>`)) || [])[1] || '');
const tagsOf = s => [...s.matchAll(/<w:([A-Za-z]+)\b/g)].map(m => m[1]);
function inlineStyles(xml) {
  return xml.replace(/<w:p\b[^>]*>[\s\S]*?<\/w:p>/g, p => {
    const id = (p.match(/<w:pStyle w:val="([^"]+)"\/>/) || [])[1]; if (!id) return p;
    const st = styleOf(id); if (!st) return p;
    const sp = inner(st, 'pPr').replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, ''), sr = inner(st, 'rPr');
    p = p.replace(/<w:pStyle w:val="[^"]+"\/>/, '');
    // pPr: เติมเฉพาะแท็กที่ย่อหน้ายังไม่มี
    const have = tagsOf(inner(p, 'pPr'));
    const addP = [...sp.matchAll(/<w:([A-Za-z]+)\b[^>]*?(?:\/>|>[\s\S]*?<\/w:\1>)/g)].filter(m => !have.includes(m[1])).map(m => m[0]).join('');
    p = /<w:pPr>/.test(p) ? p.replace('<w:pPr>', '<w:pPr>' + addP) : p.replace(/^(<w:p\b[^>]*>)/, `$1<w:pPr>${addP}</w:pPr>`);
    // rPr ของแต่ละ run: เติมเฉพาะแท็กที่ยังไม่มี
    const rTags = [...sr.matchAll(/<w:([A-Za-z]+)\b[^>]*?(?:\/>|>[\s\S]*?<\/w:\1>)/g)].map(m => ({ t: m[1], x: m[0] })).filter(o => !/^(rFonts|lang|ligatures)$/.test(o.t));
    p = p.replace(/<w:r\b(?: [^>]*)?>(?:<w:rPr>([\s\S]*?)<\/w:rPr>)?/g, (r, ri = '') => { const h = tagsOf(ri); const add = rTags.filter(o => !h.includes(o.t)).map(o => o.x).join(''); return r.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, '').replace(/>$/, `><w:rPr>${ri}${add}</w:rPr>`); });
    return p;
  });
}
let n = 0;
for (const e of T.entries.filter(e => /^word\/header\d*\.xml$/.test(e.name))) {
  const x = T.get(e.name); const txt = [...x.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('');
  if (!/โครงการ/.test(txt)) continue;
  // --raw: คัดลอกตามต้นแบบตรงๆ (เหมือนเล่ม 6 ที่ผ่าน — ความสูงหัวกระดาษเท่าต้นแบบ หน้าออกมาเท่ากัน)
  let out = process.argv.includes('--raw') ? sXml : inlineStyles(sXml); const tR = rels(T, e.name); let tRels = tR.xml;
  for (const m of media) {
    const name = 'media/hdrlogo_' + m.target.split('/').pop();
    if (!T.entries.find(z => z.name === 'word/' + name)) { const src = S.entries.find(z => z.name === 'word/' + m.target); T.entries.push({ name: 'word/' + name, data: src.data }); }
    const newId = 'rIdHdr' + m.id.replace(/\D/g, '');
    tRels = tRels.replace(new RegExp(`<Relationship [^>]*Id="${newId}"[^>]*/>`), '');
    tRels = tRels.replace('</Relationships>', `<Relationship Id="${newId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${name}"/></Relationships>`);
    out = out.split(`"${m.id}"`).join(`"${newId}"`);
  }
  T.set(e.name, out); T.set(tR.path, tRels); n++;
}
// --raw: หัวกระดาษอ้างสไตล์ของต้นแบบ (หัวกระดาษ1 / header / Table Grid) — ถ้าเล่มปลายทางไม่มี ให้คัดลอกสไตล์ไปด้วย
// ไม่งั้น Word ตัดการอ้างสไตล์ที่หาไม่เจอทิ้งตอนบันทึก หัวกระดาษกลายเป็นตัวปกติ สูงขึ้น หน้าเพิ่ม (เล่ม 6: 95 แทน 93)
// จับคู่สไตล์ด้วย "ชื่อ" ไม่ใช่ id (Word ไทยตั้ง id เป็น a/a0/af… ส่วนเล่มต้นฉบับอังกฤษใช้ Normal/Header… — id ชนกันคนละความหมาย)
//  - ชื่อที่เล่มมีอยู่แล้ว (header, Table Grid): ใช้นิยามของต้นแบบ แต่คง id ของเล่ม
//  - ชื่อที่เล่มไม่มี (หัวกระดาษ1, หัวกระดาษ อักขระ): เพิ่มด้วย id ใหม่ที่ไม่ซ้ำ
//  - basedOn/link/next แปลงตามชื่อ (Normal ของต้นแบบ -> Normal ของเล่ม)
if (process.argv.includes('--raw') && n) {
  let tSt = T.get('word/styles.xml'); const log = [];
  const all = s => [...s.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>[\s\S]*?<\/w:style>/g)].map(m => ({ id: m[1], name: (m[0].match(/<w:name w:val="([^"]+)"\/>/) || [])[1] || '', xml: m[0] }));
  const sAll = all(sStyles); const sById = Object.fromEntries(sAll.map(s => [s.id, s]));
  const tByName = () => Object.fromEntries(all(tSt).map(s => [s.name.toLowerCase(), s]));
  const tIds = () => new Set(all(tSt).map(s => s.id));
  const map = {}; // template id -> target id
  const need = [...new Set([...sXml.matchAll(/<w:(?:pStyle|rStyle|tblStyle) w:val="([^"]+)"\/>/g)].map(m => m[1]))];
  const resolve = (sid, depth = 0) => {
    if (map[sid] || depth > 8) return map[sid]; const s = sById[sid]; if (!s) return null;
    const tn = tByName()[s.name.toLowerCase()];
    const base = /^(normal|default paragraph font|normal table|no list)$/i.test(s.name);
    if (tn && base) { map[sid] = tn.id; return tn.id; }   // สไตล์พื้นฐาน: ใช้ของเล่มตามเดิม
    let id = tn ? tn.id : sid; if (!tn) { let k = 0; const used = tIds(); while (used.has(id)) id = 'HDR' + sid + (k++ || ''); }
    map[sid] = id;
    let xml = s.xml.replace(/w:styleId="[^"]+"/, `w:styleId="${id}"`).replace(/<w:rsid [^>]*\/>/g, '');
    for (const m of [...xml.matchAll(/<w:(basedOn|link|next) w:val="([^"]+)"\/>/g)]) { const r = resolve(m[2], depth + 1); xml = r ? xml.replace(m[0], `<w:${m[1]} w:val="${r}"/>`) : xml.replace(m[0], ''); }
    tSt = tn ? tSt.replace(tn.xml, xml) : tSt.replace('</w:styles>', xml + '</w:styles>');
    log.push(`${s.name}: ${tn ? 'แทนนิยาม ' + id : 'เพิ่ม ' + id}`); return id;
  };
  need.forEach(id => resolve(id));
  T.set('word/styles.xml', tSt);
  for (const e of T.entries.filter(e => /^word\/header\d*\.xml$/.test(e.name))) { let h = T.get(e.name); if (!/โครงการ/.test(h)) continue; h = h.replace(/<w:(pStyle|rStyle|tblStyle) w:val="([^"]+)"\/>/g, (a, t, v) => map[v] ? `<w:${t} w:val="${map[v]}"/>` : a); T.set(e.name, h); }
  console.log('header styles:', log.join(' · ') || 'none');
}
// make sure png/jpeg/emf content types exist
let ct = T.get('[Content_Types].xml');
for (const [ext, type] of [['png', 'image/png'], ['jpeg', 'image/jpeg'], ['jpg', 'image/jpeg'], ['emf', 'image/x-emf']]) if (!new RegExp(`Extension="${ext}"`, 'i').test(ct)) ct = ct.replace('<Default ', `<Default Extension="${ext}" ContentType="${type}"/><Default `);
T.set('[Content_Types].xml', ct);
saveDocx(T, dstF); console.log('headers replaced:', n);
