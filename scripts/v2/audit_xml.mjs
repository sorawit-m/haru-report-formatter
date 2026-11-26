// ตรวจแบบอ่านอย่างเดียว (ไม่แก้ไฟล์) ตามรายการปัญหาเดิม — พิมพ์จำนวนที่พบต่อหัวข้อ
import { loadDocx } from '../lib/docx.mjs';
for (const f of process.argv.slice(2)) {
  const d = loadDocx(f); const x = d.get('word/document.xml'); const st = d.get('word/styles.xml'); const R = {};
  const b = x.indexOf('<w:body>'); const ce = x.indexOf('</w:p>', x.indexOf('<w:sectPr', b)) + 6; const body = x.slice(ce);
  const shaded = tc => /<w:shd [^>]*w:fill="(?!auto|FFFFFF|ffffff)[0-9A-Fa-f]{6}"/.test(tc);
  // fonts
  const fonts = {}; for (const e of d.entries) { if (!/^word\/(document|styles|header\d*|footer\d*)\.xml$/.test(e.name)) continue; for (const m of d.get(e.name).matchAll(/w:ascii="([^"]+)"/g)) fonts[m[1]] = (fonts[m[1]] || 0) + 1; }
  R['fonts other than TH SarabunPSK'] = Object.entries(fonts).filter(([k]) => !/TH SarabunPSK|Symbol|Wingdings/.test(k)).map(([k, v]) => `${k}:${v}`).join(' ') || 0;
  // page numbers
  R['page number field in footers'] = d.entries.filter(e => /footer\d*\.xml$/.test(e.name)).some(e => /PAGE/.test(d.get(e.name))) ? 'yes' : 'NO';
  // tables
  let tables = 0, noRepeat = 0, hdrNotBold = 0, vCenter = 0, brInCell = 0, multiParaDesc = 0, indentInCell = 0, negInd = 0, charSpacing = 0;
  for (const m of body.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)) {
    const t = m[0]; const g = ((t.match(/<w:tblGrid>[\s\S]*?<\/w:tblGrid>/) || [''])[0].match(/<w:gridCol\b/g) || []).length; if (g < 2) continue; tables++;
    const rows = [...t.matchAll(/<w:tr\b[\s\S]*?<\/w:tr>/g)].map(r => r[0]);
    const tcs0 = [...rows[0].matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(c => c[1]);
    if (tcs0.length > 1 && tcs0.every(shaded)) { if (!/<w:tblHeader\/>/.test(rows[0]) && !/<w:drawing/.test(t) && rows.length > 6) noRepeat++; for (const tc of tcs0) for (const r of tc.matchAll(/<w:r(?: [^>]*)?>[\s\S]*?<\/w:r>/g)) if (/<w:t[ >][^<]*[^\s<]/.test(r[0]) && !/<w:b\/>/.test(r[0])) { hdrNotBold++; break; } }
    if (/<w:tblInd w:w="-\d+"/.test(t)) negInd++;
    for (const r of rows.slice(1)) { const tcs = [...r.matchAll(/<w:tc>([\s\S]*?)<\/w:tc>/g)].map(c => c[1]); if (tcs.every(shaded)) continue;
      for (const tc of tcs) { if (/<w:vAlign w:val="(center|bottom)"/.test(tc) && g >= 7) vCenter++; if (/<w:br\/>/.test(tc)) brInCell++;
        const ps = [...tc.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map(p => p[0]).filter(p => /<w:t[ >]/.test(p));
        if (g >= 7 && ps.length > 1 && !/<w:numPr>|<w:drawing|w:hanging/.test(tc) && ps.every(p => !/^\s*\d+\./.test([...p.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('')))) multiParaDesc++;
        for (const p of ps) if (/<w:ind [^>]*w:firstLine="[1-9]/.test(p)) indentInCell++; } }
  }
  for (const m of body.matchAll(/<w:rPr>[\s\S]*?<\/w:rPr>/g)) if (/<w:spacing w:val="-?[1-9]/.test(m[0])) charSpacing++;
  Object.assign(R, { 'tables (2+ cols)': tables, 'long tables without repeating header': noRepeat, 'header rows not bold': hdrNotBold, 'wide-table cells not top-aligned': vCenter, 'manual line breaks in cells': brInCell, 'cells split into several lines (9-col)': multiParaDesc, 'first-line indent in cells': indentInCell, 'tables into left margin': negInd, 'squeezed/stretched character spacing runs': charSpacing });
  // captions / typos
  const txt = [...body.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>/g)].map(a => a[1]).join('').replace(/​/g, '');
  R['"รูปที่10" without space'] = (txt.match(/(รูปที่|ตารางที่)\d/g) || []).length;
  R['"InPut"'] = (txt.match(/InPut/g) || []).length;
  R['Test by / Approve by'] = (txt.match(/Test(ed)? ?by|Approve(d)? ?by/gi) || []).length;
  R['intro chapter (บทนำ heading)'] = /<w:pStyle w:val="[^"]+"\/>(?:(?!<\/w:p>)[\s\S])*?>บทนำ</.test(body) ? 'present' : 0;
  // headings / TOC
  R['TOC field'] = /TOC \\/.test(x) ? 'yes' : 'NO';
  R['TOC italic style'] = [...st.matchAll(/<w:style\b[^>]*>[\s\S]*?<\/w:style>/g)].some(m => /<w:name w:val="toc \d"\/>/.test(m[0]) && /<w:i\/>/.test(m[0])) ? 'check' : 0;
  // header (page header) logo size / text
  const hdrs = d.entries.filter(e => /header\d*\.xml$/.test(e.name)).map(e => d.get(e.name)).filter(h => /โครงการ/.test(h));
  R['page header variants'] = [...new Set(hdrs.map(h => (h.match(/<wp:extent cx="(\d+)"/) || [])[1]))].map(v => v ? (v / 360000).toFixed(1) + 'cm logo' : 'no logo').join(', ');
  console.log('### ' + f.split('\\').pop()); for (const [k, v] of Object.entries(R)) console.log(`  ${k.padEnd(44)} ${v}`);
}
