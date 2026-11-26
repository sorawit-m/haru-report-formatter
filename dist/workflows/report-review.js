export const meta = {
  name: 'report-review',
  description: 'ตรวจ-แก้เล่มรายงาน: ผู้ตรวจอ่าน PDF ทุกหน้า ↔ ผู้แก้แก้ Word วนจนผ่าน (Trained by Haru)',
  whenToUse: 'หลังจัดเล่มด้วย report-formatter แล้ว ต้องการให้ผู้ตรวจกับผู้แก้คุยกันจนเล่มไม่มีข้อผิดพลาด — args: { file: "<path .docx>", maxRounds?: 4 }',
  phases: [
    { title: 'เตรียมภาพหน้า', detail: 'ส่งออก PDF และแปลงทุกหน้าเป็นภาพ' },
    { title: 'ผู้ตรวจอ่าน PDF', detail: 'แบ่งหน้าให้ผู้ตรวจหลายคนดูภาพจริงทุกหน้า' },
    { title: 'ผู้แก้แก้ Word', detail: 'แก้ทุกจุดที่ผู้ตรวจพบ แล้วส่งออกใหม่' },
    { title: 'สรุปผล', detail: 'สรุปสิ่งที่คุยกันแก้ และหน้าที่ควรดู' },
  ],
}

const FILE = args && args.file
if (!FILE) throw new Error('ต้องส่ง args.file = path ของไฟล์ .docx')
const MAX_ROUNDS = (args && args.maxRounds) || 4
const CHUNK = 25
const V = '~/.claude/skills/report-formatter/scripts/v2'

const PREP = { type: 'object', properties: { pdf: { type: 'string' }, pngDir: { type: 'string' }, pages: { type: 'number' }, autoCheck: { type: 'string' } }, required: ['pdf', 'pngDir', 'pages'] }
const ISSUES = { type: 'object', properties: { issues: { type: 'array', items: { type: 'object', properties: { page: { type: 'number' }, where: { type: 'string' }, problem: { type: 'string' }, severity: { type: 'string', enum: ['major', 'minor'] } }, required: ['page', 'where', 'problem', 'severity'] } } }, required: ['issues'] }
const FIXED = { type: 'object', properties: { fixed: { type: 'array', items: { type: 'string' } }, notFixed: { type: 'array', items: { type: 'string' } }, changedPages: { type: 'array', items: { type: 'number' } } }, required: ['fixed', 'notFixed'] }

log('🤖 เริ่มตรวจ-แก้เล่ม: ' + FILE + '  (Trained by Haru)')
const history = []
let lastPages = null

for (let round = 1; round <= MAX_ROUNDS; round++) {
  phase('เตรียมภาพหน้า')
  const prep = await agent(
    `รอบที่ ${round}: เตรียมภาพหน้าของเล่ม "${FILE}" (Windows, PowerShell)
1) ส่งออก PDF และตรวจอัตโนมัติ: powershell -NoProfile -ExecutionPolicy Bypass -File "${V}/export_check.ps1" -File "${FILE}" -Out "<โฟลเดอร์ชั่วคราวใหม่>\\r${round}" (สคริปต์นี้แปลงทุกหน้าเป็น p001.png… ไว้ในโฟลเดอร์ -Out ด้วย)
2) แปลงหน้าเป็นภาพความละเอียดสูงสำหรับอ่าน: powershell -File "~/.claude/skills/report-formatter/scripts/windows/pdf-to-png.ps1" -Pdf <pdf> -OutDir <โฟลเดอร์>\\hi -Width 1400
ตอบ path ของ PDF, โฟลเดอร์ภาพละเอียด (hi), จำนวนหน้า, และผลตรวจอัตโนมัติ (ข้อความที่สคริปต์พิมพ์)`,
    { label: `เตรียมภาพ รอบ ${round}`, phase: 'เตรียมภาพหน้า', schema: PREP })
  if (!prep) throw new Error('เตรียมภาพไม่สำเร็จ')
  log(`📄 รอบ ${round}: ${prep.pages} หน้า — ${String(prep.autoCheck || '').replace(/\s+/g, ' ').slice(0, 160)}`)

  phase('ผู้ตรวจอ่าน PDF')
  const ranges = []
  for (let s = 1; s <= prep.pages; s += CHUNK) ranges.push([s, Math.min(prep.pages, s + CHUNK - 1)])
  const focus = round > 1 && lastPages && lastPages.length ? ` ให้ความสำคัญกับหน้าที่ผู้แก้เพิ่งแก้: ${lastPages.join(', ')}` : ''
  const found = await parallel(ranges.map(([a, b]) => () => agent(
    `ตรวจเล่ม "${FILE}" หน้า PDF ${a}–${b} จากภาพใน "${prep.pngDir}" (ไฟล์ p${String(a).padStart(3, '0')}.png ถึง p${String(b).padStart(3, '0')}.png) ดูทุกหน้าจริง ซูมจุดที่ข้อความเล็ก.${focus}
ผลตรวจอัตโนมัติของทั้งเล่ม (ใช้เป็นเบาะแส): ${prep.autoCheck || '-'}
รายงานทุกปัญหาพร้อมเลขหน้า PDF ตามรายการตรวจของคุณ ถ้าช่วงนี้ไม่มีปัญหาให้ส่ง issues เป็นอาร์เรย์ว่าง`,
    { label: `ผู้ตรวจ หน้า ${a}-${b}`, phase: 'ผู้ตรวจอ่าน PDF', agentType: 'report-reviewer', schema: ISSUES })))
  const issues = found.filter(Boolean).flatMap(r => r.issues)
  const major = issues.filter(i => i.severity === 'major')
  log(`🔍 ผู้ตรวจรอบ ${round}: พบ ${issues.length} จุด (สำคัญ ${major.length})`)
  history.push({ round, pages: prep.pages, issues })
  if (!issues.length) { log('✅ ผู้ตรวจให้ผ่าน — ไม่พบปัญหา'); break }
  if (round === MAX_ROUNDS) { log(`⚠️ ครบ ${MAX_ROUNDS} รอบ ยังเหลือ ${issues.length} จุด — ส่งให้ผู้ใช้ตัดสินใจ`); break }

  phase('ผู้แก้แก้ Word')
  const list = issues.map((i, k) => `${k + 1}. [หน้า ${i.page}] ${i.where}: ${i.problem} (${i.severity})`).join('\n')
  const fix = await agent(
    `รอบที่ ${round}: แก้ไฟล์ "${FILE}" ตามรายการจากผู้ตรวจด้านล่าง แก้ที่ต้นเหตุและทุกจุดแบบเดียวกันทั้งเล่ม จากนั้นส่งออก PDF ใหม่ และเขียนบันทึกรอบนี้ต่อท้ายไฟล์ห้องคุยข้างเล่ม
ถ้าจุดไหนแก้แล้วจะแย่ลงหรือขัดกฎ ให้ใส่ใน notFixed พร้อมเหตุผล
รายการปัญหา:
${list}`,
    { label: `ผู้แก้ รอบ ${round}`, phase: 'ผู้แก้แก้ Word', agentType: 'report-fixer', schema: FIXED })
  if (fix) { history[history.length - 1].fix = fix; lastPages = fix.changedPages || null; log(`🛠️ ผู้แก้รอบ ${round}: แก้ ${fix.fixed.length} จุด, ไม่แก้ ${fix.notFixed.length} จุด`) }
}

phase('สรุปผล')
const summary = await agent(
  `เขียนสรุปภาษาไทยสำหรับผู้ใช้ จากการตรวจ-แก้เล่ม "${FILE}" (ข้อมูลรอบต่อรอบด้านล่าง) ให้มี: จำนวนรอบ, ปัญหาที่ผู้ตรวจพบแต่ละรอบ, สิ่งที่ผู้แก้แก้, สิ่งที่เหลือ/ตัดสินใจไม่แก้พร้อมเหตุผล, และหน้าที่ควรเปิดดู
บันทึกสรุปนี้ลงไฟล์ "<ชื่อเล่ม> - สรุปตรวจแก้.md" ข้างไฟล์เล่ม ปิดท้ายด้วยบรรทัด "ตรวจ-แก้โดยทีมเอเจนต์ · Trained by Haru" แล้วตอบเนื้อหาสรุปกลับมา
ข้อมูล: ${JSON.stringify(history).slice(0, 60000)}`,
  { label: 'สรุปผล', phase: 'สรุปผล' })
return summary
