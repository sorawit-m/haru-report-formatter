# ใบงาน: ทำให้ฮารุจัดเล่มบน Mac ได้เท่า Windows

> สำหรับ Claude Code ที่เปิดบนเครื่อง Mac — อ่านไฟล์นี้ แล้วทำตามจนถึง "เสร็จเมื่อ"
> ตอบผู้ใช้เป็นภาษาไทย

## สถานะตอนนี้
- ชุดจัดเล่มหลัก (v2) อยู่ที่ `scripts/v2/` สั่งงานผ่าน `scripts/v2/auto.ps1` (PowerShell, Windows เท่านั้น)
- ขั้นตอนส่วนใหญ่ (~45 ขั้น) เป็น Node `.mjs` แก้ XML ใน .docx ตรง ๆ → **ใช้บน Mac ได้เลย**
- ขั้นที่สั่ง Microsoft Word ผ่าน COM (Windows เท่านั้น) ต้องทำใหม่:
  - `remove_intro.ps1` ลบบท "บทนำ" (ใช้เฉพาะเมื่อสั่ง `-RemoveIntro`)
  - `word_finish.ps1` ขนาดสารบัญ/ปก/ตัวอักษรตาราง
  - `fit_images.ps1`, `fit_multi_images.ps1`, `enlarge_images.ps1` ย่อ/ขยายรูปให้พอดีช่อง
  - `box_strand_fix.ps1` หัวกล่องโค้ดไม่ค้างท้ายหน้า
  - `add_pagenum.ps1` เลขหน้า + อัปเดตสารบัญ
  - `export_check.ps1` ส่งออก PDF + ตรวจหน้า
  - `pdf_check.ps1`, `locate_pages.ps1` ตรวจภาพหน้า PDF (ใช้ `scripts/windows/pdf-to-png.ps1`)
  - `review_loop.ps1` วงตรวจ-แก้ด้วย `claude -p`
- ของเดิมสำหรับ Mac: `scripts/mac/run.sh`, `scripts/mac/export-pdf.sh` (ชุดเก่า v1 ได้ราว 90%)

## สิ่งที่ต้องทำ
1. **ติดตั้งเครื่องมือ** (ถามผู้ใช้ก่อนติดตั้งแต่ละตัว): Node.js (`brew install node`), poppler (`brew install poppler` → `pdftoppm`, `pdftotext`),
   LibreOffice ถ้าไม่มี Word for Mac, ฟอนต์ TH SarabunPSK
2. **ตัวสั่งงานข้ามระบบ** `scripts/v2/auto.mjs` — ลำดับขั้นเหมือน `auto.ps1` ทุกขั้น (อ่าน `auto.ps1` เป็นต้นแบบ)
   รับ `--in`, `--state`, `--open`, `--loop`, `--remove-intro` ; เขียน `_haru/state.json` + `state.js` แบบเดียวกัน ;
   บน Windows เรียก `.ps1` เดิม, บน Mac เรียกตัวใหม่ข้อ 3
3. **แทนขั้นที่ใช้ Word COM** — ทำเป็น Node แก้ XML ให้มากที่สุด (ใช้ได้ทั้งสองระบบ):
   - รูปพอดีช่อง: คำนวณจาก `w:tcW` / `w:gridCol` แล้วตั้ง `wp:extent` + `a:ext` (EMU) ตามสัดส่วน ; จำกัดความสูงตามพื้นที่หน้า
   - เลขหน้า: ใส่ field `PAGE` ใน footer ; สารบัญ: ตั้ง `<w:updateFields w:val="true"/>` ใน settings.xml
   - ส่งออก PDF: Word for Mac ผ่าน AppleScript (`osascript`) ถ้ามี — อัปเดตสารบัญก่อนส่งออก ; ไม่มีใช้ `soffice --headless --convert-to pdf`
   - ภาพหน้า: `pdftoppm -png -r 60` ; ตรวจหน้าว่าง/หัวข้อค้างท้ายหน้า: port ตรรกะจาก `pdf_check.ps1` (ใช้ภาพ + `pdftotext -layout`)
   - ลบบทนำ: ทำใน XML (หา heading ระดับ 1 "บทนำ" ถึงก่อน heading ระดับ 1 ถัดไป คงการขึ้น section)
4. **`install.sh`** ติดตั้งสกิล + `dist/agents` + `dist/workflows` ลง `~/.claude/` และตรวจเครื่องมือ (เหมือน `install.ps1`)
5. อัปเดต `SKILL.md` ให้บอกคำสั่งบน Mac และ `README-ติดตั้งและวิธีใช้.md`
6. แอปฮารุ (`assets/haru_app.html`) สั่ง `powershell.exe` — ให้เลือกคำสั่งตามระบบ (`node auto.mjs` บน Mac) ใช้ `os.platform()` จาก desktop API

## กฎที่ห้ามหลุด (จากงานจริง — ละเอียดใน `SKILL.md` และ `reference/rules.md`)
- ห้ามแตะต้นฉบับ: ผลออกเป็น `<ชื่อ> (จัดเล่ม).docx` + `.pdf` คู่กัน
- ข้อความสีแดงห้ามแก้ ; ห้ามแต่งเนื้อหา ; ฟอนต์ TH SarabunPSK ทั้งเล่ม
- คงหัวกระดาษเดิมของแต่ละเล่ม (ไม่มีต้นแบบหัวกระดาษแล้ว)
- เก็บบทนำไว้เสมอ เอาออกเฉพาะเมื่อผู้ใช้สั่ง
- เล่มรายการทดสอบ: แถวข้อมูลกับรูปของกรณีเดียวกันต้องอยู่หน้าเดียวกัน (ย่อรูปถ้าจำเป็น), หัวตารางห้ามตัดกลางคำ
- **ห้ามทำให้ฝั่ง Windows พัง:** อย่าลบ/เปลี่ยนพฤติกรรม `.ps1` เดิม

## เสร็จเมื่อ
- บน Mac: `node scripts/v2/auto.mjs --in <เล่มทดสอบ.docx>` ได้ `(จัดเล่ม).docx` + `.pdf` + สรุปภาษาไทย
- ทดสอบอย่างน้อย 2 เล่ม (เล่มทั่วไป 1 + เล่มรายการทดสอบที่มีรูป 1) — ผู้ใช้ให้ไฟล์เอง ; ดูภาพทุกหน้าแล้วเทียบกับผลบน Windows ถ้ามี
- `install.sh` ติดตั้งบนเครื่องใหม่ได้ ; พิมพ์ "ฮารุ" / "จัดเล่ม" ใน Claude Code บน Mac แล้วใช้งานได้
- commit + push ขึ้น repo นี้ (ไม่ใส่ไฟล์เอกสารลูกค้า)
