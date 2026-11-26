// วงตรวจ-แก้หลังจัดเล่ม (Trained by Haru) — ใช้ได้ทั้ง Mac และ Windows (แทน review_loop.ps1)
// ผู้ตรวจพบหน้าไหน -> ผู้แก้ (Claude Opus ผ่าน `claude -p`) แก้หน้านั้นทันที -> ส่งออก PDF ใหม่ -> ผู้ตรวจอ่านซ้ำ -> วนจนไม่เจอ (สูงสุด maxRounds รอบ)
// อ่าน/เขียน <state>/state.json ให้แอปฮารุแสดงผลสด ; `claude` ใช้บัญชีที่ล็อกอินในเครื่อง
// usage: node review_loop.mjs --state <dir> [--max-rounds 3]
import fs from 'node:fs'; import os from 'node:os'; import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
const S = path.dirname(fileURLToPath(import.meta.url)); const WIN = process.platform === 'win32';
const argv = process.argv.slice(2); const get = k => { const i = argv.indexOf('--' + k); return i >= 0 ? argv[i + 1] : undefined; };
const State = path.resolve(get('state') || ''); const MaxRounds = +(get('max-rounds') || 3);
if (!get('state')) { console.error('usage: node review_loop.mjs --state <dir> [--max-rounds 3]'); process.exit(1); }
const sj = path.join(State, 'state.json');
const run = (cmd, args, input) => { const r = spawnSync(cmd, args, { encoding: 'utf8', input, maxBuffer: 1 << 28 }); return ((r.stdout || '') + (r.stderr || '') + (r.error ? r.error.message : '')).trim(); };
const ps1 = (p, ...a) => run('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', p, ...a.map(String)]);
const mac = (sh, ...a) => run('/bin/bash', [path.join(S, '..', 'mac', sh), ...a.map(String)]);
const exportPdf = (docx, pdf, out) => WIN ? ps1(path.join(S, 'export_check.ps1'), '-File', docx, '-Out', out) : mac('word.sh', docx, pdf);
const checkPdf = (pdf, out) => WIN ? ps1(path.join(S, 'pdf_check.ps1'), '-Pdf', pdf, '-Out', out) : mac('pdftool.sh', 'check', pdf, out);
const pngPage = (pdf, dir, n) => WIN ? ps1(path.join(S, '..', 'windows', 'pdf-to-png.ps1'), '-Pdf', pdf, '-OutDir', dir, '-Width', 1100, '-From', n, '-To', n) : mac('pdftool.sh', 'png', pdf, dir, 1100, n, n);
const NowMs = () => Date.now(); const hhmm = () => new Date().toTimeString().slice(0, 5);
let ST = JSON.parse(fs.readFileSync(sj, 'utf8').replace(/^﻿/, ''));
function Save() { const j = JSON.stringify(ST); fs.writeFileSync(sj + '.tmp', j, 'utf8'); fs.renameSync(sj + '.tmp', sj); fs.writeFileSync(path.join(State, 'state.js'), `window.HARU_STATE&&window.HARU_STATE(${j});`, 'utf8'); }
function Say(from, to, text, extra = {}) { ST.feed = [...(ST.feed || []), { from, to, text, t: hhmm(), ts: NowMs(), ...extra }]; Save(); }
const docx = ST.files.docx, pdf = ST.files.pdf; let total = +ST.pages.total; const q = Math.max(1, Math.ceil(total / 4));
const shotDir = path.join(State, 'shots'); fs.mkdirSync(path.join(shotDir, 'after'), { recursive: true });
ST.phases = [...ST.phases, { name: 'ตรวจ-แก้วนจนผ่าน', status: 'active', sub: 'เริ่ม', frac: 0 }]; ST.status = 'running'; ST.ended = null; Save();
const lab = { 'blank': 'หน้าว่าง', 'heading/table header at page end': 'หัวข้อหรือหัวตารางค้างท้ายหน้า', 'past right margin': 'มีของล้นขอบขวา', 'starts lower than normal': 'เนื้อหาเริ่มต่ำกว่าหน้าอื่น' };
const renderHint = WIN ? `scripts/windows/pdf-to-png.ps1 -Pdf "${pdf}" -OutDir <โฟลเดอร์> -From N -To N` : `bash scripts/mac/pdftool.sh png "${pdf}" <โฟลเดอร์> 1100 N N`;
for (let round = 1; round <= MaxRounds; round++) {
  const issues = Object.entries(ST.pages.map).filter(([, v]) => v === 'issue').map(([k]) => +k).sort((a, b) => a - b);
  if (!issues.length) break;
  ST.round = round + 1; const ph = ST.phases[ST.phases.length - 1]; ph.sub = `รอบ ${round} · แก้ ${issues.length} หน้า`;
  ST.agents.fix = { on: true, state: `แก้รอบ ${round}`, task: 'เปิดไฟล์ Word', p: 0 }; Save();
  Say('fix', 'rev', `รับ ${issues.length} หน้า เริ่มแก้ทีละหน้าเลย`);
  let i = 0;
  for (const n of issues) {
    i++; const why = String(ST.pages.notes[n] || ''); const lane = Math.min(3, Math.floor((n - 1) / q));
    Object.assign(ST.agents.fix, { task: `หน้า ${n}: ${why}`, p: Math.round(100 * (i - 1) / issues.length) }); ph.frac = (i - 1) / issues.length * 0.7; Save();
    const prompt = `คุณคือผู้แก้ของทีมฮารุ (สกิล report-formatter ที่ ~/.claude/skills/report-formatter) แก้ไฟล์ "${docx}"
ปัญหาที่ผู้ตรวจพบ: หน้า ${n} ของ PDF "${pdf}" — ${why}
ทำ: ดูภาพหน้า ${n} (${renderHint.replace(/N/g, n)}) หาว่าเกิดจากอะไรในไฟล์ Word แล้วแก้ให้หาย โดยใช้วิธีตามบทเรียนของสกิล (keep_heads, fit_tables, cells_top ฯลฯ)
ห้าม: เปลี่ยนเนื้อหา/คำ, ใส่เครดิตในไฟล์, แตะต้นฉบับ, ส่งออก PDF (ระบบส่งออกเองหลังแก้ครบ)
ตอบบรรทัดเดียวภาษาไทย ว่าแก้อะไร (ถ้าแก้ไม่ได้ ให้บอกเหตุผลสั้นๆ)`;
    const out = run('claude', ['-p', '--model', 'opus', '--output-format', 'json', '--permission-mode', 'acceptEdits', '--allowedTools', WIN ? 'Read,Glob,Grep,Edit,Write,Skill,Bash,PowerShell' : 'Read,Glob,Grep,Edit,Write,Skill,Bash'], prompt);
    let res = ''; try { res = String(JSON.parse(out).result || '').trim(); } catch { res = out.replace(/\s+/g, ' ').trim(); }
    if (!res) res = 'แก้แล้ว'; if (res.length > 200) res = res.slice(0, 199) + '…';
    ST.pages.map[n] = 'fixed'; (ST.shots[n] ||= {}).fix = res;
    Say('fix', 'rev', res, { tag: 'fixed', pages: [n], lane });
  }
  // ส่งออกใหม่ + ผู้ตรวจอ่านซ้ำ
  const ph2 = ST.phases[ST.phases.length - 1];
  Object.assign(ST.agents.fix, { task: 'ส่งออก PDF ใหม่ให้ผู้ตรวจ', p: 100 }); Object.assign(ph2, { sub: `รอบ ${round} · ส่งออก PDF`, frac: 0.75 }); Save();
  const chk = fs.mkdtempSync(path.join(os.tmpdir(), 'rf_loop_'));
  exportPdf(docx, pdf, chk);
  ST.agents.fix = { on: false, state: 'รอผลตรวจซ้ำ', task: `แก้รอบ ${round} ครบ ${issues.length} หน้า`, p: 100 };
  Object.assign(ST.agents.rev, { on: true, state: `ตรวจซ้ำรอบ ${round}`, task: 'อ่านซ้ำหน้าที่แก้ และทุกหน้าที่อาจเลื่อน' }); ph2.sub = `รอบ ${round} · ตรวจซ้ำ`; Save();
  Say('fix', 'rev', 'แก้ครบแล้ว PDF ใหม่มาแล้ว ช่วยตรวจซ้ำหน่อย', { act: 'walk' });
  const pc = checkPdf(pdf, chk);
  const nt = +((pc.match(/pages:\s*(\d+)/) || [])[1] || 0); if (nt) { total = nt; ST.pages.total = total; }
  const flag = new Map();
  for (const [k, v] of Object.entries(lab)) { const m = pc.match(new RegExp(k.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&') + ':\\s*([0-9,]*)')); if (!m) continue; for (const p of m[1].split(',').filter(Boolean)) { if (!flag.has(p)) flag.set(p, []); flag.get(p).push(v); } }
  for (const n of issues) {
    const lane = Math.min(3, Math.floor((n - 1) / q)); const afterDir = path.join(shotDir, 'after');
    pngPage(pdf, afterDir, n);
    const raw = path.join(afterDir, `p${String(n).padStart(3, '0')}.png`), png = path.join(afterDir, `p${String(n).padStart(3, '0')}_r${round}.png`);
    if (fs.existsSync(raw)) fs.renameSync(raw, png);
    (ST.shots[n] ||= {}).after = png;
    if (flag.has(String(n))) { ST.pages.map[n] = 'issue'; ST.pages.notes[n] = flag.get(String(n)).join(' · '); Say('rev', 'fix', 'ยังไม่หาย: ' + ST.pages.notes[n], { tag: 'issue', pages: [n], lane }); }
    else { ST.pages.map[n] = 'pass'; Say('rev', 'fix', `หน้า ${n} ผ่านแล้ว`, { tag: 'ok', pages: [n], lane }); }
  }
  // หน้าใหม่ที่เพิ่งมีปัญหา (เช่นเนื้อหาเลื่อน)
  for (const [p, whys] of flag) {
    const n = +p; if (issues.includes(n) || ['issue', 'pass'].includes(ST.pages.map[p])) continue;
    const lane = Math.min(3, Math.floor((n - 1) / q)); pngPage(pdf, shotDir, n);
    ST.pages.map[p] = 'issue'; ST.pages.notes[p] = whys.join(' · '); ST.shots[p] = { before: path.join(shotDir, `p${String(n).padStart(3, '0')}.png`), issue: whys.join(' · ') };
    Say('rev', 'fix', whys.join(' · '), { tag: 'issue', pages: [n], lane });
  }
  ST.agents.rev.on = false; ph2.frac = Math.min(0.95, round / MaxRounds); Save();
  fs.rmSync(chk, { recursive: true, force: true });
}
// สรุปใหม่
const left = Object.entries(ST.pages.map).filter(([, v]) => v === 'issue').map(([k]) => +k).sort((a, b) => a - b);
const fixedN = Object.values(ST.pages.map).filter(v => v === 'pass').length;
const ph = ST.phases[ST.phases.length - 1]; ph.status = 'done'; ph.sub = left.length ? `เหลือ ${left.length} หน้า` : 'ผ่านทุกหน้า';
if (ST.summary) {
  ST.summary.checks = left.map(p => ({ p, t: String(ST.pages.notes[p] || '') }));
  ST.summary.verdict = left.length ? `ดู ${left.length} หน้า` : 'ผ่าน';
  ST.summary.headline = left.length ? `แก้ไป ${fixedN} หน้า เหลือ ${left.length} หน้าที่ควรเปิดดู` : `แก้ครบ ${fixedN} หน้า ตรวจซ้ำผ่านทุกหน้า`;
}
ST.agents.fix = { on: false, state: 'เสร็จ', task: `แก้ไป ${fixedN} หน้า`, p: 100 };
Say('rev', 'prep', left.length ? `จบวงตรวจ-แก้ เหลือ ${left.length} หน้าที่ต้องให้คนตัดสิน` : 'ตรวจซ้ำผ่านทุกหน้าแล้ว!', { tag: left.length ? 'issue' : 'ok' });
ST.status = 'done'; ST.ended = NowMs(); Save();
console.log(`loop done: fixed ${fixedN}, left ${left.length}`);
