#!/bin/bash
# export-pdf.sh — ส่งออก PDF บน Mac (ใช้ Word for Mac ถ้ามี ไม่งั้นใช้ LibreOffice) แล้วแปลงหน้าเป็นภาพไว้ตรวจด้วยตา
# ใช้: bash export-pdf.sh <in.docx> <out.pdf> [png-folder]
set -e
IN="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"; OUT="$2"; PNG="$3"
OUTABS="$(cd "$(dirname "$OUT")" && pwd)/$(basename "$OUT")"
if [ -d "/Applications/Microsoft Word.app" ]; then
  # Word for Mac: เปิดไฟล์ -> อัปเดตสารบัญ/ฟิลด์ -> บันทึกเป็น PDF  (ครั้งแรก macOS อาจถามสิทธิ์ให้ Terminal ควบคุม Word: กดอนุญาต)
  osascript <<EOF
tell application "Microsoft Word"
  open POSIX file "$IN"
  set d to active document
  try
    repeat with t in (tables of contents of d)
      update t
    end repeat
  end try
  try
    repeat with f in (tables of figures of d)
      update f
    end repeat
  end try
  save as d file name (POSIX file "$OUTABS" as string) file format format PDF
  close d saving no
end tell
EOF
elif command -v soffice >/dev/null 2>&1 || [ -x "/Applications/LibreOffice.app/Contents/MacOS/soffice" ]; then
  SOFF="$(command -v soffice || echo /Applications/LibreOffice.app/Contents/MacOS/soffice)"
  TMP="$(mktemp -d)"; "$SOFF" --headless --convert-to pdf --outdir "$TMP" "$IN" >/dev/null
  mv "$TMP"/*.pdf "$OUTABS"; rmdir "$TMP"
  echo "หมายเหตุ: ใช้ LibreOffice — การตัดหน้าอาจต่างจาก Word เล็กน้อย และสารบัญจะอัปเดตเมื่อเปิดไฟล์ใน Word"
else
  echo "ไม่พบ Microsoft Word หรือ LibreOffice — ติดตั้ง LibreOffice (ฟรี) แล้วลองใหม่"; exit 1
fi
echo "PDF: $OUTABS"
if [ -n "$PNG" ]; then
  mkdir -p "$PNG"
  if command -v pdftoppm >/dev/null 2>&1; then pdftoppm -r 40 -png "$OUTABS" "$PNG/p"
  else
    # ใช้ตัวแปลงที่มากับ macOS (ทีละหน้า)
    N=$(mdls -raw -name kMDItemNumberOfPages "$OUTABS" 2>/dev/null || echo 0)
    if [ "$N" -gt 0 ] 2>/dev/null; then for ((i=1;i<=N;i++)); do sips -s format png "$OUTABS" --out "$PNG/p$(printf %03d $i).png" >/dev/null 2>&1 || true; done; fi
    echo "แนะนำติดตั้ง poppler (brew install poppler) เพื่อแปลง PDF เป็นภาพทุกหน้าได้ครบ"
  fi
  echo "PNG: $PNG"
fi
