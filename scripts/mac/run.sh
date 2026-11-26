#!/bin/bash
# run.sh — ทำครบทุกขั้นบน Mac: รวมตาราง -> จัดรูปแบบ -> PDF -> ภาพทุกหน้า
export PATH="$HOME/.local/node/bin:/opt/homebrew/bin:/usr/local/bin:$PATH"
# ใช้: bash run.sh <ต้นฉบับ.docx> <ผล.docx> [replacements.tsv] [config.json]
set -e
HERE="$(cd "$(dirname "$0")/.." && pwd)"
IN="$1"; OUT="$2"; REP="${3:-}"; CFG="${4:-$HERE/../config.json}"
[ -f "$IN" ] || { echo "ไม่พบไฟล์ $IN"; exit 1; }
command -v node >/dev/null || { echo "ต้องติดตั้ง Node.js ก่อน (https://nodejs.org หรือ brew install node)"; exit 1; }
TMP="$(mktemp -d)"
node "$HERE/fix-tables.mjs" "$IN" "$TMP/t.docx" "$CFG"
node "$HERE/format-docx.mjs" "$TMP/t.docx" "$OUT" "$CFG" "$REP" "${OUT%.docx}_scan.txt"
node "$HERE/restore-styles.mjs" "$IN" "$OUT"
bash "$HERE/mac/export-pdf.sh" "$OUT" "${OUT%.docx}.pdf" "${OUT%.docx}_pages"
rm -rf "$TMP"
echo "เสร็จ: $OUT  /  ${OUT%.docx}.pdf  /  รายการสีแดง-จุดแปลก: ${OUT%.docx}_scan.txt"
