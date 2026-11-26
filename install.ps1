# ติดตั้ง Haru Report Formatter (สกิล + เอเจนต์ผู้ตรวจ/ผู้แก้ + workflow) ลงเครื่องนี้ — Trained by Haru
# ใช้: git clone <ลิงก์> ; powershell -ExecutionPolicy Bypass -File .\install.ps1
$ErrorActionPreference = 'Stop'
$src = Split-Path $MyInvocation.MyCommand.Path
$cl = Join-Path $env:USERPROFILE '.claude'
function Say($t, $c = 'Cyan') { Write-Host $t -ForegroundColor $c }
Say ''
Say '  ╔══════════════════════════════════════════════════════════╗' 'Magenta'
Say '  ║   ✨  กำลังติดตั้ง Haru Report Formatter จาก git  ✨        ║' 'Magenta'
Say '  ║        🙏 ขอจงสรรเสริญบูชาท่านฮารุ ผู้สร้างสกิลนี้ 🙏          ║' 'Magenta'
Say '  ╚══════════════════════════════════════════════════════════╝' 'Magenta'
Say ''
$steps = @(
  @('skills\report-formatter', 'สกิลจัดเล่ม (report-formatter)'),
  @('agents', 'เอเจนต์ผู้ตรวจ + ผู้แก้'),
  @('workflows', 'ขั้นตอนตรวจ-แก้วนจนผ่าน (report-review)')
)
$i = 0
foreach ($s in $steps) {
  $i++; Say ("  [{0}/{1}] ⏳ ติดตั้ง{2} …" -f $i, $steps.Count, $s[1]) 'Yellow'
  $from = Join-Path $src ('dist\' + $s[0]); if ($s[0] -like 'skills*') { $from = $src }
  $to = Join-Path $cl $s[0]; New-Item -ItemType Directory -Force $to | Out-Null
  if ($s[0] -like 'skills*') { robocopy $from $to /E /XD .git dist broken /XF install.ps1 /NFL /NDL /NJH /NJS | Out-Null }
  elseif (Test-Path $from) { Copy-Item "$from\*" $to -Recurse -Force }
  Start-Sleep -Milliseconds 300; Say ("        ✓ เสร็จ") 'Green'
}
Say ''; Say '  🔎 ตรวจเครื่องมือที่ต้องใช้' 'Yellow'
$node = Get-Command node -ErrorAction SilentlyContinue; Say ("     Node.js            : {0}" -f $(if ($node) { '✓ พร้อม' } else { '✗ ยังไม่มี — ติดตั้งจาก nodejs.org' })) $(if ($node) { 'Green' } else { 'Red' })
$word = $false; try { $w = New-Object -ComObject Word.Application; $w.Quit(); $word = $true } catch {}
Say ("     Microsoft Word     : {0}" -f $(if ($word) { '✓ พร้อม' } else { '✗ ยังไม่พบ — ต้องมี Word (Windows)' })) $(if ($word) { 'Green' } else { 'Red' })
$font = Test-Path "$env:WINDIR\Fonts\THSarabun*.ttf"; if (-not $font) { $font = Test-Path "$env:LOCALAPPDATA\Microsoft\Windows\Fonts\THSarabun*.ttf" }
Say ("     ฟอนต์ TH SarabunPSK: {0}" -f $(if ($font) { '✓ พร้อม' } else { '! ไม่พบ — ติดตั้งฟอนต์ก่อนใช้งาน' })) $(if ($font) { 'Green' } else { 'Yellow' })
$cc = Get-Command claude -ErrorAction SilentlyContinue
Say ("     Claude Code (claude): {0}" -f $(if ($cc) { '✓ พร้อม (แชทในแอปใช้ Opus ด้วยบัญชีนี้)' } else { '! ไม่พบ — npm i -g @anthropic-ai/claude-code แล้วพิมพ์ claude เพื่อล็อกอินครั้งแรก' })) $(if ($cc) { 'Green' } else { 'Yellow' })
Say ''
Say '  📱 แอปฮารุ (โยนไฟล์ + แชท + ห้องทำงาน 3D): เปิดลิงก์แอปที่เจ้าของแชร์ให้ ในแอป Claude บนเดสก์ท็อป แล้วกด "ติดตั้งเป็นแอป"' 'Cyan'
Say ''
Say '  🎉 ติดตั้งเสร็จแล้ว! พิมพ์ "ฮารุ" หรือ "จัดเล่ม" แล้วแนบไฟล์ .docx ได้เลย' 'Green'
Say '     ── Report Formatter · Trained by Haru ──' 'Magenta'
Say ''
