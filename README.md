# ฮารุ · Haru Report Formatter

จัดเล่มรายงานภาษาไทย (.docx) อัตโนมัติด้วย Claude Code — ฟอนต์เดียวทั้งเล่ม, ตารางขอบเท่ากัน หัวตารางซ้ำทุกหน้า,
ตัดคำไทยไม่ขาดกลางคำ, หัวข้อไม่ค้างท้ายหน้า, ลบหน้าว่าง, สารบัญถูกต้อง แล้วส่งออก PDF พร้อมสรุปภาษาไทย

## ติดตั้งแบบข้อความเดียว

เปิด Claude Code (แอป Claude แท็บ Code หรือคำสั่ง `claude`) แล้ววาง:

```
ติดตั้งสกิลจาก https://github.com/sorawit-m/haru-report-formatter.git ให้หน่อย
```

เสร็จแล้วพิมพ์ **"ฮารุ"** หรือ **"จัดเล่ม"** แล้วแนบไฟล์ .docx ได้เลย (อยากเอาบทนำออก พิมพ์ "เอาบทนำออก" ด้วย)

## สำหรับ Claude: วิธีติดตั้ง

ทำตามนี้ทีละข้อ ตอบผู้ใช้เป็นภาษาไทย

1. ดาวน์โหลด repo นี้ไปที่โฟลเดอร์ชั่วคราว (`git clone`)
2. **Windows:** รัน `powershell -ExecutionPolicy Bypass -File <repo>\install.ps1`
   ตัวติดตั้งคัดลอกสกิลไป `~/.claude/skills/report-formatter` + เอเจนต์ `dist/agents` + workflow `dist/workflows` แล้วตรวจเครื่องมือ
   **Mac:** ถ้ายังไม่มี `install.sh` ให้คัดลอกเองแบบเดียวกัน (repo → `~/.claude/skills/report-formatter` ยกเว้น `.git` `dist`,
   `dist/agents/*` → `~/.claude/agents/`, `dist/workflows/*` → `~/.claude/workflows/`) และบอกผู้ใช้ว่าบน Mac ตอนนี้ใช้ได้ชุดเก่า
   (`scripts/mac/run.sh`) จนกว่างานใน `MAC_PORT.md` จะเสร็จ
3. ถ้าขาดเครื่องมือ ให้**ขออนุญาตผู้ใช้ก่อน**แล้วติดตั้ง:
   - Node.js — Windows: `winget install --id OpenJS.NodeJS.LTS -e` · Mac: `brew install node`
   - Microsoft Word (Windows) — ติดตั้งให้ไม่ได้ (ต้องมีไลเซนส์) ให้แจ้งผู้ใช้ ; Mac ใช้ LibreOffice แทนได้ (`brew install --cask libreoffice`)
   - ฟอนต์ TH SarabunPSK — แจ้งผู้ใช้ให้ติดตั้งถ้ายังไม่มี
4. บอกผู้ใช้ว่าติดตั้งเสร็จ และเริ่มใช้ได้ด้วยการพิมพ์ "ฮารุ" หรือ "จัดเล่ม" — สกิลจะโหลดใน**แชทใหม่**

## ใช้ได้บนเครื่องไหน

| เครื่อง | ต้องมี | ความสามารถ |
| --- | --- | --- |
| Windows | Word, Node.js, ฟอนต์ TH SarabunPSK | ครบทุกขั้น + วงตรวจ-แก้ + แอปฮารุ |
| Mac | Node.js, Word for Mac หรือ LibreOffice | ชุดเก่า ~90% (กำลังทำให้ครบ ดู `MAC_PORT.md`) |

รายละเอียดการใช้งาน: [`README-ติดตั้งและวิธีใช้.md`](README-ติดตั้งและวิธีใช้.md) · กฎการจัดเล่มทั้งหมด: [`SKILL.md`](SKILL.md)

— Report Formatter · Trained by Haru
