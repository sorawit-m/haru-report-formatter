# ใบงาน: ทำให้ฮารุจัดเล่มบน Mac ได้เท่า Windows — ✅ เสร็จแล้ว (26-11-2026)

## ทำอะไรไปแล้ว
- `scripts/v2/auto.mjs` ตัวสั่งงานข้ามระบบ ลำดับขั้นเหมือน `auto.ps1` ทุกขั้น (Windows เรียก .ps1 เดิม, Mac เรียกตัวใหม่) — state.json/state.js/_jobs.txt/สรุปภาษาไทย แบบเดียวกัน
- แทนขั้น Word COM ด้วย Node แก้ XML ตรง (ใช้ได้ทั้งสองระบบ): `word_finish.mjs`, `add_pagenum.mjs`, `fit_images.mjs` (fit/enlarge/multi), `remove_intro.mjs`
- ขั้นที่ต้องรู้เลขหน้า (`box_strand_fix.mjs`, `locate_pages.mjs`) อ่านตำแหน่งจาก PDF ที่ Word for Mac ส่งออก
- `scripts/mac/word.sh` อัปเดตสารบัญ + ส่งออก PDF ผ่าน AppleScript (คัดลอกเข้า sandbox ของ Word ; ปิด Word เฉพาะเมื่อเอกสารที่เปิดทั้งหมดเป็นของฮารุ เพราะ Word for Mac ปิดเอกสารผ่าน AppleScript ไม่ได้)
- `scripts/mac/pdftool.swift` (+ `pdftool.sh`) แทน `pdf-to-png.ps1` / `pdf_check.ps1` ด้วย PDFKit: `png`, `check`, `text`, `sheet`
  - `check` ไม่นับแถวหัว-ท้ายกระดาษที่เหมือนกันทุกหน้า (เดิมเส้นท้ายกระดาษสีฟ้าทำให้ทุกหน้าถูกนับว่า "หัวข้อค้างท้ายหน้า")
- `scripts/v2/review_loop.mjs` วงตรวจ-แก้ข้ามระบบ ; `install.sh` ติดตั้ง + ติดตั้งเครื่องมือที่ขาดให้เอง
- แอปฮารุ (`assets/haru_app.html`) เลือกคำสั่งตามระบบ (Mac: `node auto.mjs` / `review_loop.mjs`, path แบบ `/`)
- แก้บั๊กที่กระทบทั้งสองระบบ: ขั้นปก (cover_size_fix, cover_black, content_size, one_font, text_flow, no_midword, word_finish) ถือ "ทั้ง section แรก" เป็นปก
  → เล่มที่ปก/ทะเบียนเวอร์ชัน/สารบัญอยู่ section เดียวกัน ตัวหนังสือสารบัญกลายเป็น 24pt ; เอกสาร section เดียวคำนวณตำแหน่งผิดจนไฟล์เสียได้
  → ใช้ `coverEnd()` (lib/wordxml.mjs): ปกจบที่ page break แรก หรือจบ section แรก แล้วแต่อันไหนมาก่อน
- ปกเต็มหน้าแล้วบรรทัด page break ล้นเป็นหน้าว่าง → word_finish ย้ายการขึ้นหน้าใหม่ไปที่ย่อหน้าถัดไป

## ทดสอบแล้ว (Mac, Word for Mac 16)
- เล่มทั่วไป 28 หน้า (รายงานส่งมอบงวดงาน): ผ่าน, ตารางล้นขอบขวาหายหมด, ไม่มีหน้าว่าง, สารบัญถูก
- เล่ม API/กล่องโค้ด 56 หน้า (Course Book): ผ่าน ~80 วินาที
- รูปพอดีช่องในเล่ม MA 76 หน้า: ย่อ 17 รูป เค้าโครงไม่เพี้ยน

## ยังไม่ได้ทดสอบกับไฟล์จริง
- เล่มรายการทดสอบ (testlist) และรายงานผลการทดสอบ (testresult) — ไม่มีไฟล์ตัวอย่างในเครื่อง (ขั้นที่ใช้: fit_images enlarge/multi, ตัวอักษรตารางทดสอบใน word_finish)
- วงตรวจ-แก้ `--loop` ต้องมี Claude Code (`claude`) และล็อกอินแล้ว

## กฎที่ห้ามหลุด (จากงานจริง — ละเอียดใน `SKILL.md` และ `reference/rules.md`)
- ห้ามแตะต้นฉบับ: ผลออกเป็น `<ชื่อ> (จัดเล่ม).docx` + `.pdf` คู่กัน
- ข้อความสีแดงห้ามแก้ ; ห้ามแต่งเนื้อหา ; ฟอนต์ TH SarabunPSK ทั้งเล่ม
- คงหัวกระดาษเดิมของแต่ละเล่ม ; เก็บบทนำไว้เสมอ เอาออกเฉพาะเมื่อผู้ใช้สั่ง
- **ห้ามทำให้ฝั่ง Windows พัง:** ไม่ได้ลบ/เปลี่ยนพฤติกรรม `.ps1` เดิม
