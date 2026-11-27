// จัดเล่มอัตโนมัติทั้งเล่ม (ใช้ได้ทั้ง Mac และ Windows) — ลำดับขั้นเหมือน auto.ps1 ทุกขั้น
// โยนไฟล์ .docx มา -> ได้ .docx + .pdf ที่จัดแล้ว + ไฟล์สรุปการปรับ (ภาษาไทย) ; ไม่แตะต้นฉบับ (สำรองไว้ในโฟลเดอร์ "สำรองต้นฉบับ")
// usage: node auto.mjs --in <file.docx> [--out-dir <dir>] [--remove-intro] [--keep-intro] [--state <dir>] [--open] [--loop]
//   --state <โฟลเดอร์> : เขียน state.json + state.js + ภาพหน้าที่ต้องเช็ค ให้แอปฮารุแสดงความคืบหน้า
//   --open : เปิดหน้าติดตามงาน (ห้อง 3D + สรุปผล) ในเบราว์เซอร์ ; --loop : จบแล้วเข้าวงตรวจ-แก้ (Opus) จนไม่เจอปัญหา
// Windows: ขั้นที่ต้องใช้ Word เรียก .ps1 เดิม ; Mac: ใช้ .mjs ที่แก้ XML ตรง + scripts/mac/word.sh (ส่งออก PDF) + scripts/mac/pdftool.sh (ตรวจภาพหน้า)
import fs from 'node:fs'; import os from 'node:os'; import path from 'node:path';
import { execFileSync, spawnSync, spawn } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const S = path.dirname(fileURLToPath(import.meta.url));
const A = path.join(S, '..', '..', 'assets');
const WIN = process.platform === 'win32';
const argv = process.argv.slice(2); const opt = {};
for (let i = 0; i < argv.length; i++) { const a = argv[i]; if (!a.startsWith('--')) { if (!opt.in) opt.in = a; continue; } const k = a.slice(2); if (['in', 'out-dir', 'state'].includes(k)) opt[k] = argv[++i]; else opt[k] = true; }
if (!opt.in) { console.error('usage: node auto.mjs --in <file.docx> [--state <dir>] [--open] [--loop] [--remove-intro]'); process.exit(1); }
const src = path.resolve(opt.in); const OutDir = opt['out-dir'] ? path.resolve(opt['out-dir']) : path.dirname(src);
const name = path.basename(src, path.extname(src));
const bk = path.join(path.dirname(src), 'สำรองต้นฉบับ'); fs.mkdirSync(bk, { recursive: true });
if (!fs.existsSync(path.join(bk, path.basename(src)))) fs.copyFileSync(src, path.join(bk, path.basename(src)));
const f = path.join(OutDir, name + ' (จัดเล่ม).docx'); fs.copyFileSync(src, f);
const pdfOf = p => p.replace(/\.docx$/i, '.pdf');
const log = [];

// ---- คำสั่งย่อย
const node = (script, ...args) => run(process.execPath, [path.join(S, script), ...args.map(String)]);
const ps1 = (script, ...args) => run('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(S, script), ...args.map(String)]);
const mac = (script, ...args) => run('/bin/bash', [path.join(S, '..', 'mac', script), ...args.map(String)]);
function run(cmd, args) { const r = spawnSync(cmd, args, { encoding: 'utf8', maxBuffer: 1 << 28 }); return ((r.stdout || '') + (r.stderr || '') + (r.error ? String(r.error.message) : '')).trim(); }

// ---- สถานะสำหรับแอปฮารุ
const NowMs = () => Date.now(); const hhmm = () => new Date().toTimeString().slice(0, 5); const hms = () => new Date().toTimeString().slice(0, 8);
const State = opt.state ? path.resolve(opt.state) : ''; const EXPECT = 26; let ST = null;
if (State) {
  fs.mkdirSync(State, { recursive: true });
  const lanes = [0, 1, 2, 3].map(i => ({ name: 'กขคง'[i] + ' · ช่วง ' + (i + 1), p: 0 }));
  ST = {
    title: path.basename(src), kind: '', status: 'running', realStart: NowMs(), ended: null, round: 0, job: path.dirname(src),
    phases: [{ name: 'จัดเล่มอัตโนมัติ', status: 'active', sub: 'กำลังเริ่ม', frac: 0 }, { name: 'ตรวจทุกหน้า', status: '', sub: '', frac: 0 }, { name: 'สรุปผล', status: '', sub: '', frac: 0 }],
    agents: { prep: { on: true, state: 'กำลังทำ', task: 'เปิดไฟล์และดูชนิดเล่ม', p: 0 }, rev: { on: false, state: 'รอคิว', task: '', lanes }, fix: { on: false, state: 'รอคำสั่ง', task: 'พิมพ์สั่งในแชทได้เลย' }, sum: { on: false, state: 'รอคิว', task: '' } },
    feed: [], steps: [], pages: { total: 0, map: {}, notes: {} }, shots: {}, summary: null,
    files: { original: src, docx: f, pdf: pdfOf(f), summary: path.join(OutDir, name + ' (จัดเล่ม) - สรุปการปรับ.txt'), backup: bk },
  };
}
function SaveState() { if (!ST) return; const j = JSON.stringify(ST); const tmp = path.join(State, 'state.tmp'); fs.writeFileSync(tmp, j, 'utf8'); fs.renameSync(tmp, path.join(State, 'state.json')); fs.writeFileSync(path.join(State, 'state.js'), `window.HARU_STATE&&window.HARU_STATE(${j});`, 'utf8'); }
if (ST) {
  // ลงทะเบียนงานให้แอปฮารุเห็น (เอกสาร/Haru งานจัดเล่ม/_jobs.txt)
  const reg = path.join(os.homedir(), WIN ? 'Documents' : 'Documents', 'Haru งานจัดเล่ม'); fs.mkdirSync(reg, { recursive: true });
  const rf = path.join(reg, '_jobs.txt'); const have = fs.existsSync(rf) ? fs.readFileSync(rf, 'utf8').replace(/^﻿/, '').split(/\r?\n/) : [];
  if (!have.includes(path.dirname(src))) fs.appendFileSync(rf, path.dirname(src) + (WIN ? '\r\n' : '\n'), 'utf8');
  if (opt.open) { const ui = path.join(A, 'haru_app.html'); if (fs.existsSync(ui)) {
    const page = '<!doctype html><html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head><body>' + fs.readFileSync(ui, 'utf8') + '</body></html>';
    const idx = path.join(State, 'index.html'); fs.writeFileSync(idx, page, 'utf8'); SaveState();
    spawn(WIN ? 'cmd' : 'open', WIN ? ['/c', 'start', '', idx] : [idx], { detached: true, stdio: 'ignore' }).unref(); } }
}
function Say(from, to, text, extra = {}) { if (!ST) return; ST.feed.push({ from, to, text, t: hhmm(), ts: NowMs(), ...extra }); SaveState(); }
const Short = o => { o = o.replace(/\s+/g, ' ').trim(); return o.length > 90 ? o.slice(0, 89) + '…' : o; };
function Step(label, fn) {
  let rec = null;
  if (ST) { rec = { name: label, status: 'run', detail: '', t: '' }; ST.steps.push(rec); const fr = Math.min(0.95, ST.steps.length / EXPECT); Object.assign(ST.phases[0], { sub: label, frac: fr }); Object.assign(ST.agents.prep, { task: label, p: Math.round(100 * fr) }); SaveState(); }
  let o; try { o = [fn()].flat().filter(Boolean).join('\n').trim(); } catch (e) { o = 'error: ' + e.message; }
  log.push({ l: label, o }); console.log(`[${label}] ${o.replace(/\s+/g, ' ')}`);
  if (rec) { Object.assign(rec, { status: 'done', detail: Short(o), t: hms() }); SaveState(); }
}

SaveState(); Say('sys', '', 'รับไฟล์ ' + path.basename(src) + ' · เริ่มจัดเล่ม');
const kind = JSON.parse(node('detect_kind.mjs', f));
if (ST) { ST.kind = { api: 'เอกสาร API', testlist: 'รายการทดสอบระบบ', testresult: 'รายงานผลการทดสอบ', general: 'เล่มรายงานทั่วไป' }[kind.kind] || kind.kind; Say('prep', '', 'ได้เล่มใหม่มาแล้ว เป็น' + ST.kind + ' เริ่มจัดเลย!'); }
console.log(`ชนิดเล่ม: ${kind.kind} (กล่องโค้ด ${kind.codeBoxes}, ตาราง Field ${kind.fieldTables}, ตารางทดสอบ ${kind.testTables})`);
const isTest = kind.kind.startsWith('test');

// ---- 0) ก่อนจัด: เฉพาะเล่มทดสอบ
if (kind.kind === 'testlist') Step('เอาแถว Test by/Approve by ออก, เว้นวรรคหลัง รูปที่', () => node('test_report_pre.mjs', f));
if (kind.kind === 'testresult') {
  Step('แก้คำ InPut -> Input', () => run(process.execPath, ['--input-type=module', '-e', `import('${new URL('../lib/docx.mjs', import.meta.url).href}').then(m=>{const d=m.loadDocx(process.argv[1]);let x=d.get('word/document.xml');let n=0;x=x.replace(/InPut/g,()=>{n++;return 'Input'});d.set('word/document.xml',x);m.saveDocx(d,process.argv[1]);console.log('fixed',n)})`, f]));
  Step('เอาแถว Test by/Approve by ออก, เว้นวรรคหลัง รูปที่', () => node('test_report_pre.mjs', f));
}

// ---- 1) ปก บทนำ เลขบท
Step('ขนาดตัวอักษรปก (ชื่อเรื่อง 24 / โครงการ 18)', () => node('cover_size_fix.mjs', f));
if (opt['remove-intro'] && !opt['keep-intro']) { Step('เอาบทนำออก', () => WIN ? ps1('remove_intro.ps1', '-File', f) : node('remove_intro.mjs', f)); Step('เลขบท/หัวข้อที่พิมพ์เอง', () => node('renum.mjs', f)); }
Step('ช่องตารางที่ใช้สไตล์หัวข้อ (ไม่ให้เข้าสารบัญ)', () => node('fix_heading_cells.mjs', f));
Step('ชื่อสารบัญไม่มีเลขนำหน้า', () => node('toc_title_nonum.mjs', f));

// ---- 2) ข้อความ
Step('รวมบรรทัดที่ขาดในช่อง + ชิดซ้าย-ขวา + กันกล่องโค้ด', () => node('text_flow.mjs', f));
Step('จุดตัดคำอังกฤษยาว', () => [node('more_breaks.mjs', f), node('camel_breaks.mjs', f)]);
Step('ฟอนต์เดียวทั้งเล่ม', () => node('one_font.mjs', f));
Step('เนื้อหา 16 pt', () => node('content_size.mjs', f, 16));
Step('ตัวหนังสือปกสีดำ', () => node('cover_black.mjs', f));

// ---- 3) ตาราง
Step('ตัดคอลัมน์ล่องหนในตาราง', () => node('grid_trim.mjs', f));
if (kind.kind === 'api') Step('ความกว้างตาราง Field เท่ากันทุกตาราง', () => node('set_widths2.mjs', f, 'Field|Type|Request|คำอธิบาย=5.0,3.4,2.2,5.4', 'Field|Type / ตำแหน่ง|Request|คำอธิบาย=5.0,3.4,2.2,5.4', 'Field|Type|Required|คำอธิบาย=5.0,3.4,2.2,5.4', 'Field|Type|คำอธิบาย=5.0,3.5,7.5', 'รหัส|Method|Endpoint|คำอธิบาย=2.0,1.9,5.6,6.5', 'HTTP Status|ความหมาย=3.0,13.0'));
if (kind.kind === 'testlist') Step('ความกว้างตารางทดสอบเท่ากันทุกตาราง', () => node('dense_fix.mjs', f, '3.3,1.2,2.6,3.0,1.9,4.8,2.6,3.6,1.4'));
if (kind.kind === 'testresult') { Step('แยกตารางทดสอบเป็นรายกรณี', () => node('split_cases.mjs', f)); Step('ความกว้างตารางทดสอบเท่ากันทุกตาราง', () => node('set_widths2.mjs', f, 'Test ID|Case ID|Feature|Description|Category|Step|Input Data|Expected Result|Test Result=2.2,1.2,3.0,3.6,1.9,4.35,3.6,3.1,1.45')); }
Step('ตารางไม่ล้ำขอบซ้าย/ขวา', () => node('fit_tables.mjs', f));
if (isTest) {
  Step('ขั้นตอน (Step) เป็นข้อ 1. 2. 3.', () => node('steps_list.mjs', f, 'Step'));
  Step('รวมข้อความบรรยายที่ขาดเป็นท่อน', () => node('merge_desc_cols.mjs', f));
  Step('แยกบรรทัด User/Password', () => [node('label_split.mjs', f), node('br_to_para.mjs', f)]);
  Step('ช่องตารางชิดบน', () => node('cells_top.mjs', f));
}
Step('หัวตารางตัวหนา + ไม่มีย่อหน้าในตาราง', () => node('table_cleanup.mjs', f));

// ---- 4) คำไทย ชิดซ้าย-ขวา
Step('จุดตัดระหว่างคำไทย (ไม่ตัดกลางคำ) + ไม่บีบตัวอักษร', () => node('thai_words.mjs', f));
if (kind.kind === 'testresult') Step('รวมคำที่ถูกตัดกลางคำ (ถูกต้อง, ผู้ใช้)', () => node('tag_thai.mjs', f, '--no-compound'));
Step('ย่อหน้าอังกฤษชิดซ้าย-ขวาแบบไม่ถ่างตัวอักษร', () => [node('latin_justify.mjs', f), node('mixed_back.mjs', f)]);
if (isTest) Step('ข้อความไทยในตารางทดสอบชิดซ้าย-ขวาแบบไทย', () => node('thai_justify_tables.mjs', f));

// ---- 5) การตัดหน้า
Step('บรรทัดว่างหลังตาราง / ตัวคั่นระหว่างตาราง', () => [node('drop_after_table.mjs', f), node('thin_separators.mjs', f)]);
Step('หัวข้อ/หัวตารางติดกับแถวแรก', () => node('keep_heads.mjs', f));
Step('ตารางสั้นอยู่หน้าเดียว (ไม่รวมตารางสารบัญ)', () => node('keep_short_tables.mjs', f));
Step('หัวตารางซ้ำเมื่อข้ามหน้า', () => node('repeat_headers.mjs', f));
if (kind.kind === 'testlist') Step('ป้ายกับรูปอยู่หน้าเดียวกัน, ไม่ซ้ำหัวตารางเหนือรูป', () => [node('pic_rows.mjs', f), node('no_repeat_wide.mjs', f)]);
if (kind.kind === 'testresult') Step('ผลที่พบติดกับรูป', () => node('case_layout6.mjs', f));

// ---- 6) หัวกระดาษ + Word
const tpl = path.join(A, 'header_template.docx');
const hk = fs.existsSync(tpl) ? node('header_kind.mjs', f, tpl).trim() : 'own';
if (hk === 'old') Step('หัวกระดาษโครงการแบบเก่า -> แบบมาตรฐาน', () => [node('copy_header.mjs', tpl, f, '--raw'), node('fit_header_width.mjs', f)]);
else Step('คงหัวกระดาษเดิมของเล่ม ปรับความกว้างให้พอดีหน้า', () => node('fit_header_width.mjs', f));
const tf = kind.kind === 'testlist' ? 13 : kind.kind === 'testresult' ? 12 : 0, th = kind.kind === 'testresult' ? 14 : 0, pt = kind.kind === 'testresult' ? 13 : 0;
Step('สารบัญ 16 pt / ปก / ตัวอักษรตาราง', () => WIN ? ps1('word_finish.ps1', '-File', f, '-TestFont', tf, '-TestHead', th, '-PicText', pt) : node('word_finish.mjs', f, tf, th, pt));
if (kind.kind === 'testlist') Step('รูปพอดีช่อง', () => WIN ? ps1('fit_images.ps1', '-File', f) : node('fit_images.mjs', f, 'fit'));
if (kind.kind === 'testresult') Step('ขยายรูปให้ชัด / หลายรูปเรียงบรรทัดเดียว', () => WIN ? [ps1('enlarge_images.ps1', '-File', f, '-MaxHeightShare', 0.68), ps1('fit_multi_images.ps1', '-File', f)] : [node('fit_images.mjs', f, 'enlarge', 0.68), node('fit_images.mjs', f, 'multi')]);
if (kind.codeBoxes > 0) Step('หัวกล่องโค้ดไม่ค้างท้ายหน้า', () => WIN ? ps1('box_strand_fix.ps1', '-File', f) : node('box_strand_fix.mjs', f));
Step('หัวข้อ/สารบัญไม่มีตัวล่องหน', () => node('clean_heads.mjs', f));
Step('เลขหน้า + ฟอนต์เดียว (ขั้นสุดท้าย)', () => WIN ? ps1('add_pagenum.ps1', '-File', f) : node('add_pagenum.mjs', f));

// ---- 7) ตรวจ
if (ST) {
  Object.assign(ST.phases[0], { status: 'done', sub: `${ST.steps.length} ขั้นตอน` }); Object.assign(ST.phases[1], { status: 'active', sub: 'ส่งออก PDF' }); ST.round = 1;
  ST.agents.prep = { on: false, state: 'คุมงาน', task: 'จัดเล่มครบแล้ว รอผลตรวจ', p: 100 };
  Object.assign(ST.agents.rev, { on: true, state: 'อ่านทุกหน้า', task: 'ส่งออก PDF แล้วอ่านทุกหน้าแบบภาพ' });
  Say('prep', 'rev', 'จัดเล่มเสร็จแล้ว ฝากอ่านทุกหน้าหน่อยนะ', { act: 'walk' });
}
const chk = fs.mkdtempSync(path.join(os.tmpdir(), 'rf_check_'));
const exp = WIN ? ps1('export_check.ps1', '-File', f, '-Out', chk) : mac('word.sh', f, pdfOf(f));
const wordNote = (exp.match(/^note: .*$/m) || [''])[0];
if (ST) { Object.assign(ST.phases[1], { sub: 'ตรวจในไฟล์ Word', frac: 0.3 }); ST.agents.rev.lanes.forEach(l => l.p = 30); SaveState(); }
const aud = node('audit_xml.mjs', f);
if (ST) { Object.assign(ST.phases[1], { sub: 'อ่านภาพทุกหน้า', frac: 0.55 }); ST.agents.rev.lanes.forEach(l => l.p = 60); SaveState(); }
const pc = WIN ? ps1('pdf_check.ps1', '-Pdf', pdfOf(f), '-Out', chk) : mac('pdftool.sh', 'check', pdfOf(f), chk);
const pngPage = (pdf, dir, n, n2 = n, w = 1100) => WIN ? run('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(S, '..', 'windows', 'pdf-to-png.ps1'), '-Pdf', pdf, '-OutDir', dir, '-Width', String(w), '-From', String(n), '-To', String(n2)]) : mac('pdftool.sh', 'png', pdf, dir, w, n, n2);
if (ST) {
  // หน้าที่ต้องเช็ค + ภาพแคป (ภาพอยู่ใน <State>/shots — แอปลบเมื่อกด "จบงาน")
  const total = +((pc.match(/pages:\s*(\d+)/) || [])[1] || 0);
  ST.pages.total = total; for (let i = 1; i <= total; i++) ST.pages.map[String(i)] = 'ok';
  const lab = { 'blank': 'หน้าว่าง', 'heading/table header at page end': 'หัวข้อหรือหัวตารางค้างท้ายหน้า', 'past right margin': 'มีของล้นขอบขวา', 'big white gap mid-page': 'ช่องว่างใหญ่กลางหน้า', 'starts lower than normal': 'เนื้อหาเริ่มต่ำกว่าหน้าอื่น', 'less than 25% used': 'ใช้พื้นที่หน้าไม่ถึง 1 ใน 4' };
  const flag = new Map();
  for (const [k, v] of Object.entries(lab)) {
    if (kind.kind === 'testresult' && /gap|25%/.test(k)) continue;   // เล่มผลทดสอบ: รูปใหญ่ขึ้นหน้าใหม่ = ปกติ
    const m = pc.match(new RegExp(k.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&') + ':\\s*([0-9,]*)')); if (!m) continue;
    for (const p of m[1].split(',').filter(Boolean)) { if (!flag.has(p)) flag.set(p, []); flag.get(p).push(v); }
  }
  const shotDir = path.join(State, 'shots'); fs.mkdirSync(shotDir, { recursive: true });
  const q = Math.max(1, Math.ceil(total / 4)); let k = 0;
  for (const p of [...flag.keys()].sort((a, b) => a - b)) {
    k++; if (k > 150) { Say('rev', 'prep', `ยังมีอีก ${flag.size - 150} หน้าที่ควรดู (ดูรายการเต็มในไฟล์สรุป)`); break; }
    const n = +p, why = flag.get(p).join(' · '), lane = Math.min(3, Math.floor((n - 1) / q));
    pngPage(pdfOf(f), shotDir, n);
    const png = path.join(shotDir, `p${String(n).padStart(3, '0')}.png`);
    ST.pages.map[p] = 'issue'; ST.pages.notes[p] = why; ST.shots[p] = { before: png, issue: why };
    ST.agents.rev.task = `หน้า ${n}: ${why}`; ST.agents.rev.lanes[lane].at = n;
    Say('rev', 'prep', why, { tag: 'issue', pages: [n], lane });
  }
  ST.agents.rev.lanes.forEach(l => { l.p = 100; delete l.at; });
  // ภาพย่อทุกหน้า (หน้าปกติก็เปิดดูได้ในหน้าเว็บ) — เล่มเกิน 400 หน้าข้ามไว้ เพราะไฟล์ภาพเยอะเกิน
  ST.thumbs = {};
  if (total && total <= 400) {
    const thumbDir = path.join(shotDir, 'all'); fs.mkdirSync(thumbDir, { recursive: true });
    pngPage(pdfOf(f), thumbDir, 1, total, 700);
    for (let i = 1; i <= total; i++) { const tp = path.join(thumbDir, `p${String(i).padStart(3, '0')}.png`); if (fs.existsSync(tp)) ST.thumbs[String(i)] = tp; }
  }
  const nf = flag.size;
  ST.agents.rev = { on: false, state: 'ส่งผลแล้ว', task: nf ? `อ่านครบ ${total} หน้า มี ${nf} หน้าที่ควรเปิดดู` : `อ่านครบ ${total} หน้า ไม่พบจุดผิดปกติ`, lanes: ST.agents.rev.lanes };
  Object.assign(ST.phases[1], { status: 'done', sub: nf ? `ควรดู ${nf} หน้า` : 'ไม่พบปัญหา' }); Object.assign(ST.phases[2], { status: 'active', sub: 'กำลังเขียน' });
  Say('rev', 'sum', nf ? `อ่านครบ ${total} หน้า มี ${nf} หน้าที่ควรเปิดดู ฝากสรุปด้วย` : `อ่านครบ ${total} หน้า เรียบร้อยทุกหน้า ฝากสรุปด้วย`, { act: 'walk', lane: 0 });
  ST.agents.sum = { on: true, state: 'กำลังเขียน', task: 'เทียบกับต้นฉบับ หาว่าเปลี่ยนอะไรไปบ้าง พร้อมเลขหน้า', p: 30 }; SaveState();
}

// ---- 8) สรุป
const sum = [];
sum.push(`สรุปการจัดเล่ม: ${name}`, `ไฟล์ผลลัพธ์: ${f} (+ .pdf)`, `ต้นฉบับสำรองไว้ที่: ${bk}`, `ชนิดเล่ม: ${kind.kind}`, '', 'สิ่งที่ปรับ:');
const skipped = [];
for (const e of log) {
  const o = e.o.replace(/\s+/g, ' ').trim();
  const nums = [...o.matchAll(/(?<![\w.])(\d+)(?![\w.])/g)].map(m => +m[1]);
  if (/no intro|already present/.test(o) || (nums.length && nums.reduce((a, b) => a + b, 0) === 0)) { skipped.push(e.l); continue; }
  const pairs = [...o.matchAll(/"?([A-Za-z][A-Za-z ]+?)"?\s*[:=]\s*(\d+)/g)].filter(m => +m[2] > 0).map(m => `${m[1].trim()} ${m[2]}`);
  sum.push(` - ${e.l}  (${pairs.length ? pairs.slice(0, 4).join(', ') : o.slice(0, 90)})`);
}
if (skipped.length) sum.push('', 'ตรวจแล้วไม่ต้องแก้: ' + skipped.join(' / '));
sum.push('', 'ผลตรวจทุกหน้า:', kind.kind === 'testresult' ? pc.split(/\r?\n/).filter(l => !/^big white gap/.test(l)).join('\n') : pc.trim(), '', 'ผลตรวจในไฟล์:', aud.trim());
if (wordNote) sum.push('', wordNote.replace(/^note: /, 'หมายเหตุ: '));
let cc = ''; try { cc = String(node('check_consistency.mjs', f) ?? '').trim(); sum.push('', 'ควรเช็คเนื้อหา (ฮารุรายงานอย่างเดียว ไม่แก้เอง):', cc); } catch (e) { sum.push('', 'ควรเช็คเนื้อหา: ตรวจไม่ได้ (' + e.message + ')'); }
const sf = path.join(OutDir, name + ' (จัดเล่ม) - สรุปการปรับ.txt'); fs.writeFileSync(sf, '﻿' + sum.join(os.EOL) + os.EOL, 'utf8');
console.log(''); console.log(sum.join('\n'));

// ---- 9) ผลสรุปสำหรับแอปฮารุ
if (ST) {
  const cj = path.join(State, 'changes.json');
  node('change_report.mjs', src, f, cj);
  Object.assign(ST.agents.sum, { task: 'หาเลขหน้าของแต่ละจุดที่เปลี่ยน', p: 70 }); SaveState();
  if (WIN) ps1('locate_pages.ps1', '-File', f, '-Json', cj); else node('locate_pages.mjs', pdfOf(f), cj);
  let cats = []; try { cats = [].concat(JSON.parse(fs.readFileSync(cj, 'utf8').replace(/^﻿/, '')) ?? []); /* PowerShell เขียนรายการ 1 หมวดเป็นวัตถุเดี่ยว ไม่ใช่ array */ } catch {}
  let auto = 0; const changes = [];
  for (const c of [...cats].sort((a, b) => b.count - a.count)) {
    auto += +c.count;
    changes.push({ label: c.label, count: +c.count, ex: (c.examples || []).slice(0, 6).map(e => ({ p: e.page ? +e.page : null, b: String(e.before ?? ''), a: String(e.after ?? '') })) });
  }
  const checks = Object.keys(ST.shots).sort((a, b) => a - b).map(p => ({ p: +p, t: ST.shots[p].issue }));
  const nf = checks.length, mins = Math.max(1, Math.round((NowMs() - ST.realStart) / 60000));
  ST.summary = {
    headline: nf ? `จัดเล่มเสร็จแล้ว มี ${nf} หน้าที่ควรเปิดดู` : 'จัดเล่มเสร็จแล้ว พร้อมส่ง',
    sub: 'เนื้อหาเหมือนต้นฉบับ จัดใหม่ให้ระเบียบ ชิดซ้าย-ขวา ฟอนต์เดียว ตารางเท่ากัน หัวข้อไม่ค้างท้ายหน้า',
    verdict: nf ? `ดู ${nf} หน้า` : 'ผ่าน',
    stats: [{ v: `${ST.pages.total}`, l: 'หน้า' }, { v: `${nf}`, l: 'หน้าที่ควรเปิดดู' }, { v: `${ST.steps.length}`, l: 'ขั้นตอนที่ทำ' }, { v: `${auto}`, l: 'จุดที่ปรับอัตโนมัติ' }, { v: `${mins} นาที`, l: 'เวลาที่ใช้' }],
    changes, checks,
    content: cc.split(/\r?\n/).filter(l => l.includes(' · ')),
    files: [{ ext: 'docx', name: path.basename(f), path: f }, { ext: 'pdf', name: path.basename(pdfOf(f)), path: pdfOf(f) }, { ext: 'txt', name: path.basename(sf), path: sf }],
  };
  ST.phases.forEach(ph => ph.status = 'done'); ST.phases[2].sub = nf ? `ดู ${nf} หน้า` : 'พร้อมส่ง';
  ST.agents.sum = { on: false, state: 'เสร็จ', task: 'ส่งสรุปให้ฮารุแล้ว', p: 100 };
  Say('sum', 'prep', nf ? `สรุปเสร็จแล้ว มี ${nf} หน้าที่ควรเปิดดู` : 'สรุปเสร็จแล้ว เล่มนี้เรียบร้อย', { act: 'walk' });
  ST.status = 'done'; ST.ended = NowMs(); SaveState();
  if (opt.loop && nf) console.log(WIN ? ps1('review_loop.ps1', '-State', State) : node('review_loop.mjs', '--state', State));
}
fs.rmSync(chk, { recursive: true, force: true });
