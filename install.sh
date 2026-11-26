#!/bin/bash
# ติดตั้ง Haru Report Formatter บน Mac (สกิล + เอเจนต์ผู้ตรวจ/ผู้แก้ + workflow) — Trained by Haru
# ใช้: git clone <ลิงก์> ; bash install.sh
# เครื่องมือที่ขาด ติดตั้งให้เองเสมอ (ดาวน์โหลดจากแหล่งทางการ ตรวจ checksum) — ยกเว้น Microsoft Word (ต้องมีไลเซนส์)
set -e
SRC="$(cd "$(dirname "$0")" && pwd)"; CL="$HOME/.claude"
c() { printf "\033[%sm%s\033[0m\n" "$1" "$2"; }
echo
c 35 '  ╔══════════════════════════════════════════════════════════╗'
c 35 '  ║   ✨  กำลังติดตั้ง Haru Report Formatter จาก git  ✨        ║'
c 35 '  ║        🙏 ขอจงสรรเสริญบูชาท่านฮารุ ผู้สร้างสกิลนี้ 🙏          ║'
c 35 '  ╚══════════════════════════════════════════════════════════╝'
echo
c 33 '  [1/3] ⏳ ติดตั้งสกิลจัดเล่ม (report-formatter) …'
mkdir -p "$CL/skills/report-formatter"
rsync -a --exclude .git --exclude dist --exclude 'scripts/v2/broken' --exclude install.ps1 "$SRC"/ "$CL/skills/report-formatter/"
c 32 '        ✓ เสร็จ'
c 33 '  [2/3] ⏳ ติดตั้งเอเจนต์ผู้ตรวจ + ผู้แก้ …'
mkdir -p "$CL/agents"; [ -d "$SRC/dist/agents" ] && cp "$SRC"/dist/agents/* "$CL/agents/"
c 32 '        ✓ เสร็จ'
c 33 '  [3/3] ⏳ ติดตั้งขั้นตอนตรวจ-แก้วนจนผ่าน (report-review) …'
mkdir -p "$CL/workflows"; [ -d "$SRC/dist/workflows" ] && cp "$SRC"/dist/workflows/* "$CL/workflows/"
c 32 '        ✓ เสร็จ'
echo; c 33 '  🔎 ตรวจและติดตั้งเครื่องมือที่ต้องใช้'
export PATH="$HOME/.local/node/bin:$HOME/.local/bin:/opt/homebrew/bin:/usr/local/bin:$PATH"

# --- Node.js (LTS): มี Homebrew -> brew ; ไม่มี -> tarball ทางการจาก nodejs.org ลง ~/.local/node (ไม่ต้องใช้รหัสผ่าน)
if ! command -v node >/dev/null; then
  if command -v brew >/dev/null; then brew install node
  else
    ARCH=$(uname -m); [ "$ARCH" = "x86_64" ] && ARCH=x64
    VER=$(curl -fsSL https://nodejs.org/dist/index.json | /usr/bin/python3 -I -c "import json,sys; print([x for x in json.load(sys.stdin) if x['lts']][0]['version'])")
    TMP=$(mktemp -d); F="node-$VER-darwin-$ARCH.tar.gz"
    (cd "$TMP" && curl -fsSLO "https://nodejs.org/dist/$VER/$F" && curl -fsSLO "https://nodejs.org/dist/$VER/SHASUMS256.txt" && grep " $F\$" SHASUMS256.txt | shasum -a 256 -c -) \
      || { c 31 "     ✗ ดาวน์โหลด Node.js ไม่ผ่านการตรวจ checksum — ไม่ติดตั้ง"; exit 1; }
    mkdir -p "$HOME/.local"; tar -xzf "$TMP/$F" -C "$HOME/.local"; ln -sfn "$HOME/.local/node-$VER-darwin-$ARCH" "$HOME/.local/node"; rm -rf "$TMP"
    grep -q '.local/node/bin' "$HOME/.zprofile" 2>/dev/null || printf '\n# Node.js (installed for Haru report-formatter)\nexport PATH="$HOME/.local/node/bin:$PATH"\n' >> "$HOME/.zprofile"
  fi
fi
command -v node >/dev/null && c 32 "     Node.js             : ✓ พร้อม ($(node -v))" || c 31 '     Node.js             : ✗ ติดตั้งไม่สำเร็จ'

# --- Swift (มากับ Xcode Command Line Tools) ใช้ตรวจภาพหน้า PDF
if ! command -v swiftc >/dev/null || ! swiftc --version >/dev/null 2>&1; then
  xcode-select --install 2>/dev/null || true
  c 33 '     Swift               : ! กำลังติดตั้ง Command Line Tools (กด "ติดตั้ง" ในหน้าต่างที่ขึ้นมา) แล้วรัน install.sh อีกครั้ง'
else
  bash "$CL/skills/report-formatter/scripts/mac/pdftool.sh" 2>/dev/null | head -0 || true
  c 32 '     Swift (ตรวจภาพหน้า) : ✓ พร้อม'
fi

# --- Microsoft Word (ติดตั้งให้ไม่ได้ — ต้องมีไลเซนส์) ; ไม่มี Word ใช้ LibreOffice แทน
if [ -d "/Applications/Microsoft Word.app" ]; then c 32 '     Microsoft Word      : ✓ พร้อม'
elif command -v soffice >/dev/null || [ -d "/Applications/LibreOffice.app" ]; then c 33 '     Microsoft Word      : ! ไม่พบ — ใช้ LibreOffice แทน (การตัดหน้าอาจต่างจาก Word เล็กน้อย)'
else
  if command -v brew >/dev/null; then brew install --cask libreoffice && c 33 '     Microsoft Word      : ! ไม่พบ — ติดตั้ง LibreOffice แทนแล้ว'
  else c 31 '     Microsoft Word      : ✗ ไม่พบ — ติดตั้ง Microsoft Word (ต้องมีไลเซนส์) หรือ LibreOffice จาก libreoffice.org'; fi
fi

# --- ฟอนต์ TH SarabunPSK (Word for Mac มีมาให้ในตัว)
if ls "$HOME/Library/Fonts" /Library/Fonts "/Applications/Microsoft Word.app/Contents/Resources/DFonts" 2>/dev/null | grep -qi 'sarabun'; then c 32 '     ฟอนต์ TH SarabunPSK : ✓ พร้อม'
else c 33 '     ฟอนต์ TH SarabunPSK : ! ไม่พบ — ติดตั้งฟอนต์ก่อนใช้งาน (แหล่งของหน่วยงาน/SIPA)'; fi

# --- Claude Code (claude) ใช้ในวงตรวจ-แก้และแชทในแอปฮารุ: ตัวติดตั้งทางการของ Anthropic
if ! command -v claude >/dev/null; then
  curl -fsSL https://claude.ai/install.sh | bash >/dev/null 2>&1 || true
fi
if command -v claude >/dev/null || [ -x "$HOME/.local/bin/claude" ]; then c 32 '     Claude Code (claude): ✓ พร้อม (ครั้งแรกพิมพ์ claude ในเทอร์มินัลเพื่อล็อกอิน)'
else c 33 '     Claude Code (claude): ! ติดตั้งไม่สำเร็จ — ดูวิธีที่ https://claude.com/claude-code'; fi

echo
c 36 '  📱 แอปฮารุ (โยนไฟล์ + แชท + ห้องทำงาน 3D): เปิดลิงก์แอปที่เจ้าของแชร์ให้ ในแอป Claude บนเดสก์ท็อป แล้วกด "ติดตั้งเป็นแอป"'
echo
c 32 '  🎉 ติดตั้งเสร็จแล้ว! เปิดแชทใหม่ พิมพ์ "ฮารุ" หรือ "จัดเล่ม" แล้วแนบไฟล์ .docx ได้เลย'
c 35 '     ── Report Formatter · Trained by Haru ──'
echo
