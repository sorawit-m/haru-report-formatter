// ตรวจความสม่ำเสมอของเนื้อหา (รายงานอย่างเดียว ไม่แก้) — จากงานที่ผู้ใช้แก้เอง 27-11-2026
//  1) ช่องว่างติดกัน 2 ตัวขึ้นไปในเนื้อหา/ตาราง (มักเกิดจากลบสัญลักษณ์ เช่น "→" ออกแล้วเหลือช่องว่าง)
//  2) ตัวคั่นที่ใช้ไม่เหมือนกันในคอลัมน์เดียวกันของตาราง (บางช่องใช้ "→" บางช่องใช้ช่องว่างคู่ / ">" / "-")
//  3) การอ้างข้อ TOR เดียวกันหลายแบบ เช่น "5.5 ลำดับที่ 6" กับ "5.5" (รวมปก หัว/ท้ายกระดาษ)
// ผลเป็นบรรทัด "ไฟล์ส่วนไหน · ข้อความ · เหตุผล" ให้คนเลือกแบบเดียวกันทั้งเล่ม
import { loadDocx } from '../lib/docx.mjs';
const textOf = (xml) => [...xml.matchAll(/<w:t(?: [^>]*)?>([^<]*)<\/w:t>|<w:tab\/>/g)].map((m) => (m[1] === undefined ? '\t' : m[1])).join('').replace(/[​‌‍⁠]/g, '');
const isCode = (p) => /Consolas|Courier|Code/i.test(p);
const out = [];
for (const f of process.argv.slice(2)) {
  const d = loadDocx(f);
  const parts = ['word/document.xml', ...d.entries.map((e) => e.name).filter((n) => /word\/(header|footer)\d+\.xml/.test(n))];
  const refs = new Map(); // base clause -> Map(variant -> [where])
  let dbl = 0;
  for (const part of parts) {
    const x = d.get(part); if (!x) continue;
    const where = part.includes('header') ? 'หัวกระดาษ' : part.includes('footer') ? 'ท้ายกระดาษ' : 'เนื้อหา';
    // 1) double spaces
    for (const m of x.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)) {
      if (isCode(m[0])) continue;
      const t = textOf(m[0]);
      const sp = /\S( {2,})\S/.exec(t);
      if (sp && dbl < 40) {
        // ข้อความรอบ ๆ จุดที่มีช่องว่างเกิน (~25 ตัวแต่ละข้าง) + แยกสาเหตุ: 2 ตัว = มักลบสัญลักษณ์ / 3 ตัวขึ้นไป = ใช้ช่องว่างจัดตำแหน่ง
        const gap = sp[1], at = sp.index + 1;
        const ctx = ((at > 25 ? '…' : '') + t.slice(Math.max(0, at - 25), at + gap.length + 25) + (at + gap.length + 25 < t.length ? '…' : '')).trim();
        const why = gap.length >= 3 ? `ใช้ช่องว่าง ${gap.length} ตัวจัดตำแหน่ง ควรใช้ Tab หรือตาราง` : 'อาจลบสัญลักษณ์แล้วเหลือช่องว่าง';
        const line = `ช่องว่างเกิน · ${where} · "${ctx}" · ${why}`;
        const prev = out.findIndex(o => o.split(' (พบ ')[0] === line);
        if (prev >= 0) { const n = +(out[prev].match(/\(พบ (\d+) ที่\)$/)?.[1] || 1) + 1; out[prev] = `${line} (พบ ${n} ที่)`; } else { dbl++; out.push(line); }
      }
      // 3) TOR clause variants: "<n.n...> ลำดับที่ <k>" vs bare "<n.n...>" next to TOR wording
      for (const r of t.matchAll(/(?:TOR[^0-9]{0,12})?\b(\d+\.\d+(?:\.\d+)*)(\s*ลำดับที่\s*\d+)?/g)) {
        const base = r[1]; if (!/TOR|ข้อ|^\d+\.\d+$/.test(r[0]) ) continue;
        // "ลำดับที่ N" ที่อยู่ไกลออกไปในประโยคเดียวกัน (เช่น "ข้อ 5.5 … ซึ่งรายการลำดับที่ 6") นับเป็นแบบมีลำดับเหมือนกัน
        const sub = r[2] || (t.slice(r.index + r[0].length, r.index + r[0].length + 120).match(/^[^0-9]*?(ลำดับที่\s*\d+)/) || [])[1];
        const v = (base + (sub ? ' ' + sub.trim() : '')).replace(/\s+/g, ' ');
        if (!refs.has(base)) refs.set(base, new Map());
        const mm = refs.get(base); if (!mm.has(v)) mm.set(v, []); mm.get(v).push(where);
      }
    }
    // 2) separators per table column
    if (part === 'word/document.xml') {
      let ti = 0;
      for (const tb of x.matchAll(/<w:tbl>(?:(?!<w:tbl>)[\s\S])*?<\/w:tbl>/g)) {
        ti++;
        const cols = new Map();
        for (const tr of tb[0].matchAll(/<w:tr[ >][\s\S]*?<\/w:tr>/g)) {
          let ci = 0;
          for (const tc of tr[0].matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)) {
            ci++; const t = textOf(tc[0]);
            const kind = /→/.test(t) ? '→' : /\S {2,}\S/.test(t) ? 'ช่องว่างคู่' : / > /.test(t) ? '>' : null;
            if (!kind) continue;
            if (!cols.has(ci)) cols.set(ci, new Map()); const c = cols.get(ci); if (!c.has(kind)) c.set(kind, []); c.get(kind).push(t.trim().slice(0, 40));
          }
        }
        for (const [ci, c] of cols) if (c.size > 1) out.push(`ตัวคั่นไม่เหมือนกัน · ตารางลำดับที่ ${ti} คอลัมน์ ${ci} · ${[...c].map(([k, v]) => `${k} ${v.length} ช่อง (เช่น "${v[0]}")`).join(' / ')} · ควรใช้แบบเดียว`);
      }
    }
  }
  for (const [base, mm] of refs) {
    const withSub = [...mm.keys()].filter((k) => k !== base);
    if (withSub.length && mm.has(base)) out.push(`อ้างข้อ TOR ไม่เหมือนกัน · ${base} · ${[...mm].map(([k, w]) => `"${k}" ${w.length} ที่ (${[...new Set(w)].join(', ')})`).join(' / ')} · เลือกแบบเดียวทั้งเล่ม รวมปกและบทนำ`);
  }
}
console.log(out.length ? out.join('\n') : 'ไม่พบ (ช่องว่างคู่ / ตัวคั่นในตาราง / การอ้างข้อ TOR สม่ำเสมอ)');
