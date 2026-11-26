# run.ps1 — ทำครบทุกขั้นบน Windows ในคำสั่งเดียว   *** บันทึกไฟล์นี้เป็น UTF-8 with BOM ***
# ใช้: powershell -NoProfile -ExecutionPolicy Bypass -File run.ps1 -In ต้นฉบับ.docx -Out ผล.docx [-Replacements คำแก้.tsv] [-Config config.json]
# ได้: ผล.docx + ผล.pdf + ผล_scan.txt (รายการสีแดง/จุดแปลก) + ผล_textcheck.txt (เทียบข้อความก่อน/หลัง)
param([string]$In, [string]$Out, [string]$Replacements = "", [string]$Config = "", [string]$Rewrites = "")
# -Rewrites: TSV (ข้อความเดิมทั้งช่อง<TAB>ข้อความใหม่) ที่ Claude เรียบเรียงให้ช่องที่ยังถ่าง — ใช้ก่อนขั้นอื่นทั้งหมด
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent (Split-Path -Parent $here)
if (-not $Config) { $Config = Join-Path $root 'config.json' }
if (-not (Get-Command node -ErrorAction SilentlyContinue)) { throw 'ต้องติดตั้ง Node.js ก่อน (https://nodejs.org)' }
$In = (Resolve-Path -LiteralPath $In).Path
$Out = [IO.Path]::GetFullPath($Out)
$lock = Join-Path (Split-Path $In) ('~$' + (Split-Path $In -Leaf).Substring(2))
if (Test-Path -LiteralPath $lock) { throw "ไฟล์ต้นฉบับเปิดอยู่ใน Word — ปิดก่อนแล้วรันใหม่" }
$tmp = Join-Path $env:TEMP ("rf_" + [guid]::NewGuid().ToString('N').Substring(0, 8)); New-Item -ItemType Directory $tmp | Out-Null
$t1 = Join-Path $tmp 't1.docx'; $pdf = [IO.Path]::ChangeExtension($Out, '.pdf')
function Step($m) { Write-Output ("[{0:HH:mm:ss}] ==> {1}" -f (Get-Date), $m) }
Step "1/6 รวมตารางหัว/เนื้อหา, หัวตารางซ้ำ, ขอบแถวเท่ากัน (XML)"
$ftLog = Join-Path $tmp 'fix-tables.log'
$env:RF_LOG = $ftLog   # fix-tables.mjs also writes its log here (UTF-8) for the review list
if ($Rewrites) { $env:RF_REWRITES = (Resolve-Path -LiteralPath $Rewrites).Path }
$t0 = Join-Path $tmp 't0.docx'
Step "0/6 ตั้งค่าไฟล์ให้เป็นมาตรฐานเดียวกัน (ขอบกระดาษ ฟอนต์ Style หัวข้อ สารบัญ)"
node (Join-Path $root 'scripts\standardize.mjs') $In $t0 $Config
node (Join-Path $root 'scripts\fix-tables.mjs') $t0 $t1 $Config
Remove-Item Env:RF_LOG, Env:RF_REWRITES -ErrorAction SilentlyContinue
if (-not (Test-Path $ftLog)) { New-Item -ItemType File $ftLog | Out-Null }
Step "2/6 จัดรูปแบบด้วย Word (ฟอนต์ ระยะ ตาราง รูป บทขึ้นหน้าใหม่ คำผิด)"
& (Join-Path $here 'format-word.ps1') -In $t1 -Out $Out -Config $Config -Replacements $Replacements | Where-Object { $_ -notmatch 'not found|tables processed' }
Start-Sleep 5
Step "3/6 เอาเลขหัวข้ออัตโนมัติที่งอกขึ้นมาออก"
& (Join-Path $here 'fix-numbering.ps1') -Orig $In -File $Out
Start-Sleep 3
Step "4/6 แก้หัวข้อ/สารบัญ (XML)"
node (Join-Path $root 'scripts\restore-styles.mjs') $In $Out
Step "5/6 ลบหน้าว่าง สร้างสารบัญใหม่ ส่งออก PDF"
& (Join-Path $here 'remove-blank-pages.ps1') -File $Out -Pdf $pdf
Step "6/6 ตรวจข้อความไม่หาย + รายการสีแดง"
$a = Join-Path $tmp 'a.txt'; $b = Join-Path $tmp 'b.txt'
node (Join-Path $root 'scripts\dump-text.mjs') $In $b | Out-Null; node (Join-Path $root 'scripts\dump-text.mjs') $Out $a | Out-Null
$sb = ([IO.File]::ReadAllText($b) -split ' ').Count - 1; $sa = ([IO.File]::ReadAllText($a) -split ' ').Count - 1
$nums = (& (Join-Path $here 'check-numbers.ps1') -Orig $In -File $Out) -join "`r`n"
# pages with a large empty area at the bottom (heading/paragraph chains that jump to the next page) -> list for review
$pg = Join-Path $tmp 'pages'; New-Item -ItemType Directory $pg | Out-Null
& (Join-Path $here 'pdf-to-png.ps1') -Pdf $pdf -OutDir $pg -Width 500 | Out-Null
$gaps = @(& (Join-Path $here 'page-gaps.ps1') -Dir $pg -Min 0.35)
$nums += "`r`nหน้าที่ว่างท้ายหน้า >= 35% (ดูว่าเป็นท้ายบท/ก่อนรูปใหญ่ หรือถูกผูกจนกระโดดหน้า): " + $(if ($gaps.Count) { "`r`n" + ($gaps -join "`r`n") } else { 'ไม่มี' })
$check = "ช่องว่างก่อน=$sb หลัง=$sa (ลดได้เฉพาะที่ตั้งใจแก้ ถ้าลดเกิน ~2% ให้ตรวจคำติดกัน)`r`n$nums"
[IO.File]::WriteAllText(([IO.Path]::ChangeExtension($Out, '') + '_textcheck.txt'), $check, (New-Object Text.UTF8Encoding($true)))
& (Join-Path $here 'scan-word.ps1') -Files @($Out) | Out-File -Encoding utf8 ([IO.Path]::ChangeExtension($Out, '') + '_scan.txt')
# ---- รายการ "ควรเข้าไปเช็ค" พร้อมเลขหน้า: คำที่เติม, คอลัมน์ที่ปรับความกว้าง, ช่องที่อาจยังถ่าง, หน้าที่ว่างมาก ----
$items = New-Object System.Collections.Generic.List[string]
foreach ($l in [IO.File]::ReadAllLines($ftLog, [Text.Encoding]::UTF8)) {
  if ($l -match '^\s*filler: (.+?)  ->  (.+)$') { $items.Add("เติมคำ`t" + $Matches[2] + "`tเดิม: " + $Matches[1]) }
  elseif ($l -match '^\s*rewrite: (.+?)  ->  (.+)$') { $items.Add("เรียบเรียงใหม่`t" + $Matches[2] + "`tเดิม: " + $Matches[1]) }
  elseif ($l -match '^\s*code column narrowed: (\S+) -> (\S+)(?: \| (.+))?$') { if ($Matches[3]) { $items.Add("ปรับความกว้างคอลัมน์`t" + $Matches[3] + "`t" + $Matches[1] + ' -> ' + $Matches[2]) } }
}
$st = Join-Path $tmp 'stretched.tsv'
node (Join-Path $root 'scripts\find-stretched.mjs') $Out 0.12 $st | Out-Null
if (Test-Path $st) { foreach ($l in [IO.File]::ReadAllLines($st, [Text.Encoding]::UTF8)) { $c = $l.TrimStart([char]0xFEFF) -split "`t"; if ($c.Count -ge 4) { $items.Add("อาจยังถ่าง`t" + $c[3] + "`tความถ่าง " + $c[2] + " em") } } }
$it = Join-Path $tmp 'items.tsv'; [IO.File]::WriteAllLines($it, $items, (New-Object Text.UTF8Encoding($false)))
$rv = [IO.Path]::ChangeExtension($Out, '') + '_ควรเช็ค.txt'
& (Join-Path $here 'review-list.ps1') -File $Out -Items $it -Out $rv | Out-Null
if ($gaps.Count) { [IO.File]::AppendAllText($rv, "`r`n`r`nหน้าที่ว่างท้ายหน้ามาก (เลขหน้า PDF — ดูว่าเป็นท้ายบท/ก่อนรูปใหญ่ หรือผิดปกติ):`r`n" + (($gaps | ForEach-Object { '- ' + $_ }) -join "`r`n"), (New-Object Text.UTF8Encoding($false))) }
Write-Output ("รายการควรเช็ค: " + $rv)
Write-Output $check
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
Step "เสร็จ: $Out  /  $pdf"
