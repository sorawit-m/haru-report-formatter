#!/bin/bash
# pdftool.sh — เรียก pdftool.swift (คอมไพล์ครั้งแรก/เมื่อแก้ไฟล์ แล้วเก็บไว้ใน ~/.cache/haru)
# ใช้: bash pdftool.sh png|check|text <pdf> …   (ดูหัวไฟล์ pdftool.swift)
set -e
HERE="$(cd "$(dirname "$0")" && pwd)"
BIN="$HOME/.cache/haru/pdftool"
if [ ! -x "$BIN" ] || [ "$HERE/pdftool.swift" -nt "$BIN" ]; then
  mkdir -p "$(dirname "$BIN")"
  swiftc -O -o "$BIN" "$HERE/pdftool.swift" 2>/dev/null || { echo "คอมไพล์ pdftool ไม่ได้ (ต้องมี Xcode Command Line Tools: xcode-select --install)" >&2; exit 1; }
fi
exec "$BIN" "$@"
