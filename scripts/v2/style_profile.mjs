// วิเคราะห์ "รูปแบบบ้าน" ของเล่ม (อ่านอย่างเดียว): หน้ากระดาษ/ขอบ, สไตล์หลัก, ปก, สารบัญ, หัว/ท้ายกระดาษ, ตาราง, คำบรรยาย, ระยะบรรทัด
import { loadDocx } from '../lib/docx.mjs';
const cm = tw => (tw / 567).toFixed(2);
for (const f of process.argv.slice(2)) {
  const d = loadDocx(f); const x = d.get('word/document.xml'); const st = d.get('word/styles.xml') || '';
  const R = {};
  // page setup (distinct)
  const secs = [...x.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map(m => m[0]);
  const ps = {}; for (const s of secs) { const W = +(s.match(/<w:pgSz [^>]*w:w="(\d+)"/) || [])[1], H = +(s.match(/<w:pgSz [^>]*w:h="(\d+)"/) || [])[1]; const g = k => cm(+((s.match(new RegExp(`<w:pgMar [^>]*w:${k}="(-?\\d+)"`)) || [])[1] || 0)); const k = `${W > H ? 'แนวนอน' : 'แนวตั้ง'} ${cm(W)}x${cm(H)} | บน ${g('top')} ล่าง ${g('bottom')} ซ้าย ${g('left')} ขวา ${g('right')} หัว ${g('header')} ท้าย ${g('footer')}`; ps[k] = (ps[k] || 0) + 1; }
  R['หน้ากระดาษ/ขอบ (ซม.)'] = Object.entries(ps).map(([k, v]) => `${v}x ${k}`).join('\n      ');
  // styles
  const sty = {}; for (const m of st.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[\s\S]*?<\/w:style>/g)) sty[m[1]] = m[0];
  const def = (st.match(/<w:rPrDefault>[\s\S]*?<\/w:rPrDefault>/) || [''])[0];
  const desc = s => { if (!s) return '-'; const nm = (s.match(/<w:name w:val="([^"]+)"/) || [])[1]; const font = (s.match(/w:ascii="([^"]+)"/) || s.match(/w:cs="([^"]+)"/) || [])[1] || ''; const sz = (s.match(/<w:sz w:val="(\d+)"/) || [])[1]; const b = /<w:b\/>/.test(s) ? 'หนา' : ''; const jc = (s.match(/<w:jc w:val="([^"]+)"/) || [])[1] || ''; const col = (s.match(/<w:color w:val="([^"]+)"/) || [])[1] || ''; const sp = s.match(/<w:spacing ([^>]*)\/>/); const ind = (s.match(/<w:ind ([^>]*)\/>/) || [])[1] || ''; const pbb = /<w:pageBreakBefore\/>/.test(s) ? 'ขึ้นหน้าใหม่' : ''; const num = /<w:numPr>/.test(s) ? 'มีเลขอัตโนมัติ' : ''; return [nm, font, sz ? sz / 2 + 'pt' : '', b, jc, col && col !== 'auto' ? '#' + col : '', sp ? sp[1].replace(/w:/g, '').replace(/"/g, '') : '', ind.replace(/w:/g, '').replace(/"/g, ''), pbb, num].filter(Boolean).join(' · '); };
  const byName = n => Object.values(sty).find(s => new RegExp(`<w:name w:val="${n}"`, 'i').test(s));
  R['ค่าเริ่มต้นเอกสาร'] = desc('<w:name w:val="defaults"/>' + def);
  for (const n of ['Normal', 'heading 1', 'heading 2', 'heading 3', 'Title', 'TOC Heading', 'toc 1', 'toc 2', 'toc 3', 'caption', 'table of figures', 'header', 'footer', 'List Paragraph'])
    { const s = byName(n); if (s) R['สไตล์ ' + n] = desc(s); }
  // style usage in body
  const use = {}; for (const m of x.matchAll(/<w:pStyle w:val="([^"]+)"/g)) use[m[1]] = (use[m[1]] || 0) + 1;
  R['สไตล์ที่ใช้มากสุด'] = Object.entries(use).sort((a, b) => b[1] - a[1]).slice(0, 8).map(([k, v]) => `${((sty[k] || '').match(/<w:name w:val="([^"]+)"/) || [, k])[1]}:${v}`).join(', ');
  // direct formatting of body text
  const szs = {}, fonts = {}, jcs = {}; for (const m of x.matchAll(/<w:sz w:val="(\d+)"/g)) szs[m[1] / 2] = (szs[m[1] / 2] || 0) + 1;
  for (const m of x.matchAll(/w:cs="([^"]+)"/g)) fonts[m[1]] = (fonts[m[1]] || 0) + 1; for (const m of x.matchAll(/<w:jc w:val="([^"]+)"/g)) jcs[m[1]] = (jcs[m[1]] || 0) + 1;
  R['ขนาดตัวอักษรที่พบ (pt:จำนวน)'] = Object.entries(szs).sort((a, b) => b[1] - a[1]).slice(0, 7).map(([k, v]) => `${k}:${v}`).join(', ');
  R['ฟอนต์ไทยที่พบ'] = Object.entries(fonts).sort((a, b) => b[1] - a[1]).slice(0, 5).map(([k, v]) => `${k}:${v}`).join(', ');
  R['การจัดย่อหน้าที่ตั้งเอง'] = Object.entries(jcs).sort((a, b) => b[1] - a[1]).map(([k, v]) => `${k}:${v}`).join(', ');
  // cover
  const b0 = x.indexOf('<w:body>'); const s0 = x.indexOf('<w:sectPr', b0); const cover = s0 > 0 ? x.slice(b0, x.indexOf('</w:p>', s0)) : '';
  const cl = [...cover.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map(p => { const t = [...p[0].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').trim(); const sz = (p[0].match(/<w:sz w:val="(\d+)"/) || [])[1]; const b = /<w:b\/>/.test(p[0]); return t ? `${t.slice(0, 50)} [${sz ? sz / 2 : '?'}pt${b ? ' หนา' : ''}]` : ''; }).filter(Boolean);
  R['ปก'] = (cl.length ? cl.join('\n      ') : '(ปกไม่ได้แยก section)') + (/<w:drawing>|<v:shape/.test(cover) ? '\n      มีรูป/โลโก้บนปก' : '');
  // TOC
  R['สารบัญ'] = [/TOC \\o/.test(x) ? 'สารบัญหัวข้อ' : '', /TOC \\[^"]*\\c "?(Figure|รูป)/.test(x) ? 'สารบัญรูป' : '', /TOC \\[^"]*\\c "?(Table|ตาราง)/.test(x) ? 'สารบัญตาราง' : ''].filter(Boolean).join(', ') || 'ไม่มี';
  // header/footer
  const hf = d.entries.filter(e => /^word\/(header|footer)\d*\.xml$/.test(e.name)).map(e => { const h = d.get(e.name); const t = [...h.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').trim(); return `${e.name.replace('word/', '')}: ${t.slice(0, 70) || '(ไม่มีข้อความ)'}${/<w:drawing>/.test(h) ? ' +รูป' : ''}${/PAGE/.test(h) ? ' +เลขหน้า' : ''}${/<w:tbl>/.test(h) ? ' (เป็นตาราง)' : ''}`; });
  R['หัว/ท้ายกระดาษ'] = hf.join('\n      ');
  // tables
  const tb = [...x.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)].map(m => m[0]); const fills = {}, borders = {}, hdrRepeat = [0, 0];
  for (const t of tb) { const r0 = (t.match(/<w:tr\b[\s\S]*?<\/w:tr>/) || [''])[0]; const fl = (r0.match(/w:fill="([0-9A-Fa-f]{6})"/) || [])[1]; if (fl) fills[fl] = (fills[fl] || 0) + 1; const bd = (t.match(/<w:top w:val="([^"]+)" w:sz="(\d+)"/) || []); if (bd[1]) borders[bd[1] + ' ' + bd[2] / 8 + 'pt'] = (borders[bd[1] + ' ' + bd[2] / 8 + 'pt'] || 0) + 1; hdrRepeat[/<w:tblHeader\/>/.test(t) ? 0 : 1]++; }
  R['ตาราง'] = `${tb.length} ตาราง | สีหัวตาราง: ${Object.entries(fills).sort((a, b) => b[1] - a[1]).slice(0, 4).map(([k, v]) => '#' + k + ':' + v).join(', ') || '-'} | เส้น: ${Object.entries(borders).slice(0, 3).map(([k, v]) => k + ':' + v).join(', ') || '-'} | หัวซ้ำ ${hdrRepeat[0]}/ไม่ซ้ำ ${hdrRepeat[1]}`;
  // captions
  const txt = [...x.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map(p => [...p[0].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').trim());
  const caps = txt.filter(t => /^(ตารางที่|รูปที่|ภาพที่|แผนภาพที่)\s*\d/.test(t));
  R['คำบรรยาย (ตัวอย่าง)'] = `${caps.length} รายการ: ` + caps.slice(0, 3).map(c => '"' + c.slice(0, 50) + '"').join(' / ');
  R['รูปภาพ'] = `${(x.match(/<w:drawing>/g) || []).length} รูป`;
  // heading numbering samples
  const heads = []; for (const m of x.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)) { const sid = (m[0].match(/<w:pStyle w:val="([^"]+)"/) || [])[1]; const s = sty[sid] || ''; if (/<w:name w:val="(heading [12]|หัวเรื่อง [12])"/i.test(s) || /<w:outlineLvl w:val="[01]"/.test(s)) { const t = [...m[0].matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').trim(); if (t) heads.push(t.slice(0, 45)); } }
  R['ตัวอย่างหัวข้อ'] = heads.slice(0, 6).join(' / ');
  console.log('\n################ ' + f.split('\\').pop()); for (const [k, v] of Object.entries(R)) console.log(`  ▸ ${k}: ${v}`);
}
