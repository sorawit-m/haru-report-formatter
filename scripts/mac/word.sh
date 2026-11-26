#!/bin/bash
# word.sh — ให้ Word for Mac อัปเดตสารบัญ/สารบัญรูป-ตาราง บันทึก และส่งออก PDF (แทนขั้น Word COM ของ Windows)
# ใช้: bash word.sh <file.docx> [out.pdf]
#   พิมพ์ "pages N" (จำนวนหน้าตาม Word) ; ไฟล์ .docx ถูกเขียนทับด้วยฉบับที่อัปเดตสารบัญแล้ว
# Word for Mac เปิดไฟล์นอก sandbox แล้วถามสิทธิ์ทุกไฟล์ -> คัดลอกเข้า ~/Library/Containers/com.microsoft.Word/Data/haru ก่อน
# Word for Mac ปิดเอกสารผ่าน AppleScript ไม่ได้ (error -1708) -> ปิด Word ทั้งโปรแกรมเฉพาะเมื่อทุกเอกสารที่เปิดอยู่เป็นของฮารุ
#   ถ้าผู้ใช้เปิดงานอื่นค้างไว้ ปล่อยเอกสารของฮารุเปิดไว้ (ชื่อขึ้นต้น haru_) แล้วบอกในผล
set -e
export PATH="$HOME/.local/node/bin:/opt/homebrew/bin:/usr/local/bin:$PATH"
IN="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"; PDF="${2:-}"
[ -f "$IN" ] || { echo "ไม่พบไฟล์ $IN" >&2; exit 1; }
[ -d "/Applications/Microsoft Word.app" ] || { echo "ไม่พบ Microsoft Word — ใช้ LibreOffice: soffice --headless --convert-to pdf" >&2; exit 2; }
BOX="$HOME/Library/Containers/com.microsoft.Word/Data/haru"; mkdir -p "$BOX"
ID="haru_$(date +%s)_$$"; W="$BOX/$ID.docx"; WP="$BOX/$ID.pdf"
cp "$IN" "$W"
OUT=$(osascript - "$W" "$WP" "$([ -n "$PDF" ] && echo 1 || echo 0)" <<'EOF'
on run argv
  set w to item 1 of argv
  set wp to item 2 of argv
  set wantPdf to item 3 of argv
  -- แปลง path นอก tell (ใน tell Word คำสั่ง POSIX file กับตัวแปรถูกส่งให้ Word แปลผิด)
  set wHfs to (POSIX file w) as string
  set wpHfs to (POSIX file wp) as string
  set nm to (do shell script "basename " & quoted form of w)
  tell application "Microsoft Word"
    open file name wHfs
    -- ตอน Word เพิ่งเปิด คำสั่ง open ไม่คืนค่าเอกสาร -> หาเอกสารจากชื่อ รอได้ถึง 60 วินาที
    set d to missing value
    repeat 120 times
      try
        set d to document nm
        exit repeat
      end try
      delay 0.5
    end repeat
    if d is missing value then error "Word เปิดไฟล์ไม่สำเร็จ: " & nm
    try
      repeat with i from 1 to (count of tables of contents of d)
        update (table of contents i of d)
      end repeat
    end try
    try
      repeat with i from 1 to (count of tables of figures of d)
        update (table of figures i of d)
      end repeat
    end try
    -- อัปเดตอีกรอบ: เลขหน้าในสารบัญอาจเลื่อนหลังรอบแรก
    try
      repeat with i from 1 to (count of tables of contents of d)
        update page numbers (table of contents i of d)
      end repeat
    end try
    -- สารบัญ/สารบัญรูป-ตาราง: 16pt ไม่เอียง (เหมือน word_finish.ps1 / add_pagenum.ps1 บน Windows)
    try
      repeat with i from 1 to (count of tables of contents of d)
        set fo to font object of text object of (table of contents i of d)
        set font size of fo to 16
        set italic of fo to false
      end repeat
    end try
    try
      repeat with i from 1 to (count of tables of figures of d)
        set fo to font object of text object of (table of figures i of d)
        set font size of fo to 16
        set italic of fo to false
      end repeat
    end try
    set pc to compute statistics d statistic statistic pages
    save d
    if wantPdf is "1" then save as d file name wpHfs file format format PDF
    return "pages " & pc
  end tell
end run
EOF
)
cp "$W" "$IN"
[ -n "$PDF" ] && mv "$WP" "$PDF"
echo "$OUT"
# ปิด Word เฉพาะเมื่อทุกเอกสารที่เปิดเป็นของฮารุ
osascript <<'EOF' 2>/dev/null || true
tell application "Microsoft Word"
  set ns to name of every document
  set mine to true
  repeat with n in ns
    if (n as string) does not start with "haru_" then set mine to false
  end repeat
  if mine then
    quit saving no
  end if
end tell
EOF
sleep 1
if pgrep -x "Microsoft Word" >/dev/null && osascript -e 'tell application "Microsoft Word" to get name of every document' 2>/dev/null | grep -q "$ID"; then
  echo "note: Word เปิดงานอื่นของผู้ใช้อยู่ จึงไม่ได้ปิด Word — เอกสาร $ID.docx ยังเปิดค้าง ปิดเองได้โดยไม่ต้องบันทึก"
else
  rm -f "$W" "$BOX/~\$${ID#haru_}.docx" 2>/dev/null || true
fi
