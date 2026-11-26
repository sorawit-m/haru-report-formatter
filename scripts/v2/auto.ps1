param([Parameter(Mandatory = $true)][string]$In, [string]$OutDir = '', [switch]$KeepIntro, [switch]$RemoveIntro, [string]$State = '', [switch]$Open, [switch]$Loop)
# -Open : เปิดหน้าติดตามงาน (ห้อง 3D + สรุปผล) ในเบราว์เซอร์ ; -Loop : จบแล้วเข้าวงตรวจ-แก้ (Opus) จนไม่เจอปัญหา
# -State <โฟลเดอร์> : เขียน state.json + ภาพหน้าที่ต้องเช็ค ให้แอปฮารุแสดงความคืบหน้า (ไม่ใส่ก็ทำงานเหมือนเดิม)
# จัดเล่มอัตโนมัติทั้งเล่ม: โยนไฟล์ .docx มา -> ได้ .docx + .pdf ที่จัดแล้ว + ไฟล์สรุปการปรับ (ภาษาไทย)
# ไม่แตะต้นฉบับ (สำรองไว้ในโฟลเดอร์ "สำรองต้นฉบับ") ; ทุกขั้นตอนมาจากเล่มที่ผู้ใช้อนุมัติแล้ว (ดู SKILL.md / report-book-lessons)
$ErrorActionPreference = 'Continue'
$S = Split-Path $MyInvocation.MyCommand.Path
$A = Join-Path (Split-Path (Split-Path $S)) 'assets'
$src = (Resolve-Path -LiteralPath $In).Path
if (-not $OutDir) { $OutDir = Split-Path $src }
$name = [IO.Path]::GetFileNameWithoutExtension($src)
$bk = Join-Path (Split-Path $src) 'สำรองต้นฉบับ'; New-Item -ItemType Directory -Force $bk | Out-Null
if (-not (Test-Path -LiteralPath (Join-Path $bk (Split-Path $src -Leaf)))) { Copy-Item -LiteralPath $src $bk }
$f = Join-Path $OutDir ($name + ' (จัดเล่ม).docx'); Copy-Item -LiteralPath $src $f -Force
$log = New-Object System.Collections.ArrayList
# ---- สถานะสำหรับแอปฮารุ
$ST = $null; $EXPECT = 26
function NowMs { [int64]([DateTimeOffset]::Now.ToUnixTimeMilliseconds()) }
if ($State) {
  New-Item -ItemType Directory -Force $State | Out-Null
  $lanes = @(0..3 | ForEach-Object { [ordered]@{ name = 'กขคง'[$_] + ' · ช่วง ' + ($_ + 1); p = 0 } })
  $ST = [ordered]@{
    title = (Split-Path $src -Leaf); kind = ''; status = 'running'; realStart = (NowMs); ended = $null; round = 0; job = (Split-Path $src)
    phases = @([ordered]@{ name = 'จัดเล่มอัตโนมัติ'; status = 'active'; sub = 'กำลังเริ่ม'; frac = 0 }, [ordered]@{ name = 'ตรวจทุกหน้า'; status = ''; sub = ''; frac = 0 }, [ordered]@{ name = 'สรุปผล'; status = ''; sub = ''; frac = 0 })
    agents = [ordered]@{ prep = [ordered]@{ on = $true; state = 'กำลังทำ'; task = 'เปิดไฟล์และดูชนิดเล่ม'; p = 0 }; rev = [ordered]@{ on = $false; state = 'รอคิว'; task = ''; lanes = $lanes }; fix = [ordered]@{ on = $false; state = 'รอคำสั่ง'; task = 'พิมพ์สั่งในแชทได้เลย' }; sum = [ordered]@{ on = $false; state = 'รอคิว'; task = '' } }
    feed = (New-Object System.Collections.ArrayList); steps = (New-Object System.Collections.ArrayList)
    pages = [ordered]@{ total = 0; map = @{}; notes = @{} }; shots = @{}; summary = $null
    files = [ordered]@{ original = $src; docx = $f; pdf = [IO.Path]::ChangeExtension($f, '.pdf'); summary = (Join-Path $OutDir ($name + ' (จัดเล่ม) - สรุปการปรับ.txt')); backup = $bk }
  }
}
function Save-State { if (-not $ST) { return }; $j = $ST | ConvertTo-Json -Depth 12 -Compress; $tmp = Join-Path $State 'state.tmp'; [IO.File]::WriteAllText($tmp, $j, (New-Object Text.UTF8Encoding $false)); Move-Item -LiteralPath $tmp (Join-Path $State 'state.json') -Force
  [IO.File]::WriteAllText((Join-Path $State 'state.js'), "window.HARU_STATE&&window.HARU_STATE($j);", (New-Object Text.UTF8Encoding $false)) }   # state.js = ให้หน้าเว็บที่เปิดจากไฟล์ (file://) อ่านได้
if ($ST) {
  # ลงทะเบียนงานให้แอปฮารุเห็น (เอกสาร\Haru งานจัดเล่ม\_jobs.txt)
  $reg = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Haru งานจัดเล่ม'; New-Item -ItemType Directory -Force $reg | Out-Null
  $rf = Join-Path $reg '_jobs.txt'; if (-not ((Test-Path $rf) -and ((Get-Content $rf -Encoding UTF8) -contains (Split-Path $src)))) { Add-Content -LiteralPath $rf -Value (Split-Path $src) -Encoding UTF8 }
  # -Open: เปิดหน้าห้องทำงาน + สรุปผลในเบราว์เซอร์ (ใช้ตอนสั่งจากแชท)
  if ($Open) { $ui = Join-Path $A 'haru_app.html'; if (Test-Path $ui) { $page = '<!doctype html><html lang="th"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head><body>' + [IO.File]::ReadAllText($ui, [Text.Encoding]::UTF8) + '</body></html>'; [IO.File]::WriteAllText((Join-Path $State 'index.html'), $page, (New-Object Text.UTF8Encoding $false)); Save-State; Start-Process (Join-Path $State 'index.html') } }
}
function Say([string]$from, [string]$to, [string]$text, [hashtable]$extra = @{}) {
  if (-not $ST) { return }; $m = [ordered]@{ from = $from; to = $to; text = $text; t = (Get-Date -Format 'HH:mm'); ts = (NowMs) }
  foreach ($k in $extra.Keys) { $m[$k] = $extra[$k] }; [void]$ST.feed.Add($m); Save-State
}
function Short([string]$o) { $o = ($o -replace '\s+', ' ').Trim(); if ($o.Length -gt 90) { $o.Substring(0, 89) + '…' } else { $o } }
# หมายเหตุ: ห้ามตั้งตัวแปรชื่อ $s/$f/$a ในฟังก์ชันนี้ — PowerShell ไม่แยกตัวพิมพ์ จะไปทับ $S (ที่อยู่สคริปต์) แล้วทุกขั้นพังเงียบๆ
function Step([string]$label, [scriptblock]$cmd) {
  $stepRec = $null
  if ($ST) { $stepRec = [ordered]@{ name = $label; status = 'run'; detail = ''; t = '' }; [void]$ST.steps.Add($stepRec); $fr = [math]::Min(0.95, $ST.steps.Count / $EXPECT); $ST.phases[0].sub = $label; $ST.phases[0].frac = $fr; $ST.agents.prep.task = $label; $ST.agents.prep.p = [int](100 * $fr); Save-State }
  $o = (& $cmd 2>&1 | Out-String).Trim(); [void]$log.Add(@{ l = $label; o = $o }); Write-Output ("[{0}] {1}" -f $label, ($o -replace '\s+', ' '))
  if ($stepRec) { $stepRec.status = 'done'; $stepRec.detail = (Short $o); $stepRec.t = (Get-Date -Format 'HH:mm:ss'); Save-State }
}
Save-State; Say 'sys' '' ('รับไฟล์ ' + (Split-Path $src -Leaf) + ' · เริ่มจัดเล่ม')
$kind = (node "$S\detect_kind.mjs" $f | ConvertFrom-Json)
if ($ST) { $ST.kind = @{ api = 'เอกสาร API'; testlist = 'รายการทดสอบระบบ'; testresult = 'รายงานผลการทดสอบ'; general = 'เล่มรายงานทั่วไป' }[$kind.kind]; if (-not $ST.kind) { $ST.kind = $kind.kind }; Say 'prep' '' ('ได้เล่มใหม่มาแล้ว เป็น' + $ST.kind + ' เริ่มจัดเลย!') }
Write-Output ("ชนิดเล่ม: {0} (กล่องโค้ด {1}, ตาราง Field {2}, ตารางทดสอบ {3})" -f $kind.kind, $kind.codeBoxes, $kind.fieldTables, $kind.testTables)

# ---- 0) ก่อนจัด: เฉพาะเล่มทดสอบ
if ($kind.kind -eq 'testlist') { Step 'เอาแถว Test by/Approve by ออก, เว้นวรรคหลัง รูปที่' { node "$S\test_report_pre.mjs" $f } }
if ($kind.kind -eq 'testresult') { Step 'แก้คำ InPut -> Input' { node -e "import('file:///$($S -replace '\\','/')/../lib/docx.mjs').then(m=>{const d=m.loadDocx(process.argv[1]);let x=d.get('word/document.xml');let n=0;x=x.replace(/InPut/g,()=>{n++;return 'Input'});d.set('word/document.xml',x);m.saveDocx(d,process.argv[1]);console.log('fixed',n)})" $f }; Step 'เอาแถว Test by/Approve by ออก, เว้นวรรคหลัง รูปที่' { node "$S\test_report_pre.mjs" $f } }

# ---- 1) ปก บทนำ เลขบท
Step 'ขนาดตัวอักษรปก (ชื่อเรื่อง 24 / โครงการ 18)' { node "$S\cover_size_fix.mjs" $f }
if ($RemoveIntro -and -not $KeepIntro) { Step 'เอาบทนำออก' { powershell -NoProfile -ExecutionPolicy Bypass -File "$S\remove_intro.ps1" -File $f }; Step 'เลขบท/หัวข้อที่พิมพ์เอง' { node "$S\renum.mjs" $f } }
Step 'ช่องตารางที่ใช้สไตล์หัวข้อ (ไม่ให้เข้าสารบัญ)' { node "$S\fix_heading_cells.mjs" $f }
Step 'ชื่อสารบัญไม่มีเลขนำหน้า' { node "$S\toc_title_nonum.mjs" $f }

# ---- 2) ข้อความ
Step 'รวมบรรทัดที่ขาดในช่อง + ชิดซ้าย-ขวา + กันกล่องโค้ด' { node "$S\text_flow.mjs" $f }
Step 'จุดตัดคำอังกฤษยาว' { node "$S\more_breaks.mjs" $f; node "$S\camel_breaks.mjs" $f }
Step 'ฟอนต์เดียวทั้งเล่ม' { node "$S\one_font.mjs" $f }
Step 'เนื้อหา 16 pt' { node "$S\content_size.mjs" $f 16 }
Step 'ตัวหนังสือปกสีดำ' { node "$S\cover_black.mjs" $f }

# ---- 3) ตาราง
Step 'ตัดคอลัมน์ล่องหนในตาราง' { node "$S\grid_trim.mjs" $f }
if ($kind.kind -eq 'api') { Step 'ความกว้างตาราง Field เท่ากันทุกตาราง' { node "$S\set_widths2.mjs" $f "Field|Type|Request|คำอธิบาย=5.0,3.4,2.2,5.4" "Field|Type / ตำแหน่ง|Request|คำอธิบาย=5.0,3.4,2.2,5.4" "Field|Type|Required|คำอธิบาย=5.0,3.4,2.2,5.4" "Field|Type|คำอธิบาย=5.0,3.5,7.5" "รหัส|Method|Endpoint|คำอธิบาย=2.0,1.9,5.6,6.5" "HTTP Status|ความหมาย=3.0,13.0" } }
if ($kind.kind -eq 'testlist') { Step 'ความกว้างตารางทดสอบเท่ากันทุกตาราง' { node "$S\dense_fix.mjs" $f "3.3,1.2,2.6,3.0,1.9,4.8,2.6,3.6,1.4" } }
if ($kind.kind -eq 'testresult') { Step 'แยกตารางทดสอบเป็นรายกรณี' { node "$S\split_cases.mjs" $f }; Step 'ความกว้างตารางทดสอบเท่ากันทุกตาราง' { node "$S\set_widths2.mjs" $f "Test ID|Case ID|Feature|Description|Category|Step|Input Data|Expected Result|Test Result=2.2,1.2,3.0,3.6,1.9,4.35,3.6,3.1,1.45" } }
Step 'ตารางไม่ล้ำขอบซ้าย/ขวา' { node "$S\fit_tables.mjs" $f }
if ($kind.kind -like 'test*') {
  Step 'ขั้นตอน (Step) เป็นข้อ 1. 2. 3.' { node "$S\steps_list.mjs" $f Step }
  Step 'รวมข้อความบรรยายที่ขาดเป็นท่อน' { node "$S\merge_desc_cols.mjs" $f }
  Step 'แยกบรรทัด User/Password' { node "$S\label_split.mjs" $f; node "$S\br_to_para.mjs" $f }
  Step 'ช่องตารางชิดบน' { node "$S\cells_top.mjs" $f }
}
Step 'หัวตารางตัวหนา + ไม่มีย่อหน้าในตาราง' { node "$S\table_cleanup.mjs" $f }

# ---- 4) คำไทย ชิดซ้าย-ขวา
Step 'จุดตัดระหว่างคำไทย (ไม่ตัดกลางคำ) + ไม่บีบตัวอักษร' { node "$S\thai_words.mjs" $f }
# เล่มผลทดสอบ: รวมคำที่ถูกตัดแต่รวมกันแล้วเป็นคำเดียวในพจนานุกรม (ถูก|ต้อง -> ถูกต้อง, ผู้|ใช้ -> ผู้ใช้) — แบบเดียวกับเล่ม 6 ที่ผ่าน
if ($kind.kind -eq 'testresult') { Step 'รวมคำที่ถูกตัดกลางคำ (ถูกต้อง, ผู้ใช้)' { node "$S\tag_thai.mjs" $f --no-compound } }
Step 'ย่อหน้าอังกฤษชิดซ้าย-ขวาแบบไม่ถ่างตัวอักษร' { node "$S\latin_justify.mjs" $f; node "$S\mixed_back.mjs" $f }
if ($kind.kind -like 'test*') { Step 'ข้อความไทยในตารางทดสอบชิดซ้าย-ขวาแบบไทย' { node "$S\thai_justify_tables.mjs" $f } }

# ---- 5) การตัดหน้า
Step 'บรรทัดว่างหลังตาราง / ตัวคั่นระหว่างตาราง' { node "$S\drop_after_table.mjs" $f; node "$S\thin_separators.mjs" $f }
Step 'หัวข้อ/หัวตารางติดกับแถวแรก' { node "$S\keep_heads.mjs" $f }
Step 'หัวตารางซ้ำเมื่อข้ามหน้า' { node "$S\repeat_headers.mjs" $f }
if ($kind.kind -eq 'testlist') { Step 'ป้ายกับรูปอยู่หน้าเดียวกัน, ไม่ซ้ำหัวตารางเหนือรูป' { node "$S\pic_rows.mjs" $f; node "$S\no_repeat_wide.mjs" $f } }
if ($kind.kind -eq 'testresult') { Step 'ผลที่พบติดกับรูป' { node "$S\case_layout6.mjs" $f } }

# ---- 6) หัวกระดาษ + Word
# หัวกระดาษ (ผู้ใช้ 2026-11-26 "คงหัวกระดาษเดิม"): คงหัวกระดาษของเล่มไว้ ปรับแค่ความกว้างให้พอดีหน้า
# ยกเว้นหัวกระดาษโครงการเดียวกันแต่เป็นแบบเก่า (สไตล์ "หัวกระดาษ") -> ใช้ต้นแบบแทน ตรงกับเล่ม 4, 5, 6 ที่ผ่าน
$hk = $(if (Test-Path "$A\header_template.docx") { (node "$S\header_kind.mjs" $f "$A\header_template.docx" | Out-String).Trim() } else { 'own' })
if ($hk -eq 'old') { Step 'หัวกระดาษโครงการแบบเก่า -> แบบมาตรฐาน' { node "$S\copy_header.mjs" "$A\header_template.docx" $f --raw; node "$S\fit_header_width.mjs" $f } }
else { Step 'คงหัวกระดาษเดิมของเล่ม ปรับความกว้างให้พอดีหน้า' { node "$S\fit_header_width.mjs" $f } }
# เล่มผลทดสอบ (ตรงกับเล่ม 6 ที่ผ่าน): แถวทดสอบ 12 pt · หัวตาราง 14 pt · ช่อง "ผลที่พบ" ที่มีรูป 13 pt
$tf = $(if ($kind.kind -eq 'testlist') { 13 } elseif ($kind.kind -eq 'testresult') { 12 } else { 0 }); $th = $(if ($kind.kind -eq 'testresult') { 14 } else { 0 }); $pt = $(if ($kind.kind -eq 'testresult') { 13 } else { 0 })
Step 'สารบัญ 16 pt / ปก / ตัวอักษรตาราง' { powershell -NoProfile -ExecutionPolicy Bypass -File "$S\word_finish.ps1" -File $f -TestFont $tf -TestHead $th -PicText $pt }
if ($kind.kind -eq 'testlist') { Step 'รูปพอดีช่อง' { powershell -NoProfile -ExecutionPolicy Bypass -File "$S\fit_images.ps1" -File $f } }
if ($kind.kind -eq 'testresult') { Step 'ขยายรูปให้ชัด / หลายรูปเรียงบรรทัดเดียว' { powershell -NoProfile -ExecutionPolicy Bypass -File "$S\enlarge_images.ps1" -File $f -MaxHeightShare 0.68; powershell -NoProfile -ExecutionPolicy Bypass -File "$S\fit_multi_images.ps1" -File $f } }
if ($kind.codeBoxes -gt 0) { Step 'หัวกล่องโค้ดไม่ค้างท้ายหน้า' { powershell -NoProfile -ExecutionPolicy Bypass -File "$S\box_strand_fix.ps1" -File $f } }
Step 'หัวข้อ/สารบัญไม่มีตัวล่องหน' { node "$S\clean_heads.mjs" $f }
Step 'เลขหน้า + ฟอนต์เดียว (ขั้นสุดท้าย)' { powershell -NoProfile -ExecutionPolicy Bypass -File "$S\add_pagenum.ps1" -File $f }

# ---- 7) ตรวจ
if ($ST) {
  $ST.phases[0].status = 'done'; $ST.phases[0].sub = "$($ST.steps.Count) ขั้นตอน"; $ST.phases[1].status = 'active'; $ST.phases[1].sub = 'ส่งออก PDF'; $ST.round = 1
  $ST.agents.prep = [ordered]@{ on = $false; state = 'คุมงาน'; task = 'จัดเล่มครบแล้ว รอผลตรวจ'; p = 100 }
  $ST.agents.rev.on = $true; $ST.agents.rev.state = 'อ่านทุกหน้า'; $ST.agents.rev.task = 'ส่งออก PDF แล้วอ่านทุกหน้าแบบภาพ'
  Say 'prep' 'rev' 'จัดเล่มเสร็จแล้ว ฝากอ่านทุกหน้าหน่อยนะ' @{ act = 'walk' }
}
$chk = Join-Path $env:TEMP ("rf_check\" + [guid]::NewGuid().ToString('N').Substring(0, 8))
$exp = (powershell -NoProfile -ExecutionPolicy Bypass -File "$S\export_check.ps1" -File $f -Out $chk | Out-String)
if ($ST) { $ST.phases[1].sub = 'ตรวจในไฟล์ Word'; $ST.phases[1].frac = 0.3; foreach ($l in $ST.agents.rev.lanes) { $l.p = 30 }; Save-State }
$aud = (node "$S\audit_xml.mjs" $f | Out-String)
if ($ST) { $ST.phases[1].sub = 'อ่านภาพทุกหน้า'; $ST.phases[1].frac = 0.55; foreach ($l in $ST.agents.rev.lanes) { $l.p = 60 }; Save-State }
$pc = (powershell -NoProfile -ExecutionPolicy Bypass -File "$S\pdf_check.ps1" -Pdf ([IO.Path]::ChangeExtension($f, '.pdf')) -Out $chk | Out-String)
if ($ST) {
  # หน้าที่ต้องเช็ค + ภาพแคป (ภาพอยู่ใน <State>\shots — แอปลบเมื่อกด "จบงาน")
  $total = [int](([regex]::Match($pc, 'pages:\s*(\d+)')).Groups[1].Value)
  $ST.pages.total = $total; for ($i = 1; $i -le $total; $i++) { $ST.pages.map["$i"] = 'ok' }
  $lab = [ordered]@{ 'blank' = 'หน้าว่าง'; 'heading/table header at page end' = 'หัวข้อหรือหัวตารางค้างท้ายหน้า'; 'past right margin' = 'มีของล้นขอบขวา'; 'big white gap mid-page' = 'ช่องว่างใหญ่กลางหน้า'; 'starts lower than normal' = 'เนื้อหาเริ่มต่ำกว่าหน้าอื่น'; 'less than 25% used' = 'ใช้พื้นที่หน้าไม่ถึง 1 ใน 4' }
  $flag = [ordered]@{}
  foreach ($k in $lab.Keys) {
    if ($kind.kind -eq 'testresult' -and $k -match 'gap|25%') { continue }   # เล่มผลทดสอบ: รูปใหญ่ขึ้นหน้าใหม่ = ปกติ
    $m = [regex]::Match($pc, [regex]::Escape($k) + ':\s*([0-9,]*)'); if (-not $m.Success) { continue }
    foreach ($p in ($m.Groups[1].Value -split ',' | Where-Object { $_ })) { if (-not $flag.Contains($p)) { $flag[$p] = @() }; $flag[$p] += $lab[$k] }
  }
  $shotDir = Join-Path $State 'shots'; New-Item -ItemType Directory -Force $shotDir | Out-Null
  $W = Join-Path (Split-Path $S) 'windows\pdf-to-png.ps1'; $q = [math]::Max(1, [math]::Ceiling($total / 4)); $k = 0
  foreach ($p in ($flag.Keys | Sort-Object { [int]$_ })) {
    $k++; if ($k -gt 150) { Say 'rev' 'prep' "ยังมีอีก $($flag.Count - 150) หน้าที่ควรดู (ดูรายการเต็มในไฟล์สรุป)"; break }
    $n = [int]$p; $why = ($flag[$p] -join ' · '); $lane = [math]::Min(3, [math]::Floor(($n - 1) / $q))
    powershell -NoProfile -ExecutionPolicy Bypass -File $W -Pdf ([IO.Path]::ChangeExtension($f, '.pdf')) -OutDir $shotDir -Width 1100 -From $n -To $n | Out-Null
    $png = Join-Path $shotDir ('p{0:D3}.png' -f $n)
    $ST.pages.map[$p] = 'issue'; $ST.pages.notes[$p] = $why
    $ST.shots[$p] = [ordered]@{ before = $png; issue = $why }
    $ST.agents.rev.task = "หน้า ${n}: $why"; $ST.agents.rev.lanes[$lane].at = $n
    Say 'rev' 'prep' $why @{ tag = 'issue'; pages = @($n); lane = $lane }
  }
  foreach ($l in $ST.agents.rev.lanes) { $l.p = 100; $l.Remove('at') }
  $nf = $flag.Count
  $ST.agents.rev = [ordered]@{ on = $false; state = 'ส่งผลแล้ว'; task = $(if ($nf) { "อ่านครบ $total หน้า มี $nf หน้าที่ควรเปิดดู" } else { "อ่านครบ $total หน้า ไม่พบจุดผิดปกติ" }); lanes = $ST.agents.rev.lanes }
  $ST.phases[1].status = 'done'; $ST.phases[1].sub = $(if ($nf) { "ควรดู $nf หน้า" } else { 'ไม่พบปัญหา' }); $ST.phases[2].status = 'active'; $ST.phases[2].sub = 'กำลังเขียน'
  Say 'rev' 'sum' $(if ($nf) { "อ่านครบ $total หน้า มี $nf หน้าที่ควรเปิดดู ฝากสรุปด้วย" } else { "อ่านครบ $total หน้า เรียบร้อยทุกหน้า ฝากสรุปด้วย" }) @{ act = 'walk'; lane = 0 }
  $ST.agents.sum = [ordered]@{ on = $true; state = 'กำลังเขียน'; task = 'เทียบกับต้นฉบับ หาว่าเปลี่ยนอะไรไปบ้าง พร้อมเลขหน้า'; p = 30 }; Save-State
}

# ---- 8) สรุป
$sum = New-Object System.Text.StringBuilder
[void]$sum.AppendLine("สรุปการจัดเล่ม: $name"); [void]$sum.AppendLine("ไฟล์ผลลัพธ์: $f (+ .pdf)"); [void]$sum.AppendLine("ต้นฉบับสำรองไว้ที่: $bk"); [void]$sum.AppendLine("ชนิดเล่ม: $($kind.kind)"); [void]$sum.AppendLine('')
[void]$sum.AppendLine('สิ่งที่ปรับ:')
$skipped = @()
foreach ($e in $log) {
  $o = ($e.o -replace '\s+', ' ').Trim()
  $nums = [regex]::Matches($o, '(?<![\w.])(\d+)(?![\w.])') | ForEach-Object { [int]$_.Groups[1].Value }
  if ($o -match 'no intro|already present' -or ($nums.Count -gt 0 -and ($nums | Measure-Object -Sum).Sum -eq 0)) { $skipped += $e.l; continue }
  $pairs = [regex]::Matches($o, '"?([A-Za-z][A-Za-z ]+?)"?\s*[:=]\s*(\d+)') | Where-Object { [int]$_.Groups[2].Value -gt 0 } | ForEach-Object { "$($_.Groups[1].Value.Trim()) $($_.Groups[2].Value)" }
  $detail = $(if ($pairs) { ($pairs | Select-Object -First 4) -join ', ' } else { $o.Substring(0, [math]::Min(90, $o.Length)) })
  [void]$sum.AppendLine(" - $($e.l)  ($detail)")
}
if ($skipped) { [void]$sum.AppendLine(''); [void]$sum.AppendLine('ตรวจแล้วไม่ต้องแก้: ' + ($skipped -join ' / ')) }
[void]$sum.AppendLine(''); [void]$sum.AppendLine('ผลตรวจทุกหน้า:'); [void]$sum.AppendLine($(if ($kind.kind -eq 'testresult') { ($pc -split "\r?\n" | Where-Object { $_ -notmatch '^big white gap' }) -join "
" } else { $pc.Trim() })); [void]$sum.AppendLine(''); [void]$sum.AppendLine('ผลตรวจในไฟล์:'); [void]$sum.AppendLine($aud.Trim())
$sf = Join-Path $OutDir ($name + ' (จัดเล่ม) - สรุปการปรับ.txt'); [IO.File]::WriteAllText($sf, $sum.ToString(), (New-Object Text.UTF8Encoding $true))
Write-Output ''; Write-Output $sum.ToString()

# ---- 9) ผลสรุปสำหรับแอปฮารุ
if ($ST) {
  $cj = Join-Path $State 'changes.json'
  node "$S\change_report.mjs" $src $f $cj | Out-Null
  $ST.agents.sum.task = 'หาเลขหน้าของแต่ละจุดที่เปลี่ยน'; $ST.agents.sum.p = 70; Save-State
  powershell -NoProfile -ExecutionPolicy Bypass -File "$S\locate_pages.ps1" -File $f -Json $cj | Out-Null
  $cats = @(); try { $cats = @(Get-Content -LiteralPath $cj -Raw -Encoding UTF8 | ConvertFrom-Json) } catch {}
  $changes = @(); $auto = 0
  foreach ($c in ($cats | Sort-Object { - [int]$_.count })) {
    $auto += [int]$c.count
    $changes += [ordered]@{ label = $c.label; count = [int]$c.count; ex = @($c.examples | Select-Object -First 6 | ForEach-Object { [ordered]@{ p = $(if ($_.page) { [int]$_.page } else { $null }); b = [string]$_.before; a = [string]$_.after } }) }
  }
  $checks = @($ST.shots.Keys | Sort-Object { [int]$_ } | ForEach-Object { [ordered]@{ p = [int]$_; t = $ST.shots[$_].issue } })
  $nf = $checks.Count; $mins = [math]::Max(1, [math]::Round(((NowMs) - $ST.realStart) / 60000))
  $ST.summary = [ordered]@{
    headline = $(if ($nf) { "จัดเล่มเสร็จแล้ว มี $nf หน้าที่ควรเปิดดู" } else { 'จัดเล่มเสร็จแล้ว พร้อมส่ง' })
    sub = 'เนื้อหาเหมือนต้นฉบับ จัดใหม่ให้ระเบียบ ชิดซ้าย-ขวา ฟอนต์เดียว ตารางเท่ากัน หัวข้อไม่ค้างท้ายหน้า'
    verdict = $(if ($nf) { "ดู $nf หน้า" } else { 'ผ่าน' })
    stats = @([ordered]@{ v = "$($ST.pages.total)"; l = 'หน้า' }, [ordered]@{ v = "$nf"; l = 'หน้าที่ควรเปิดดู' }, [ordered]@{ v = "$($ST.steps.Count)"; l = 'ขั้นตอนที่ทำ' }, [ordered]@{ v = "$auto"; l = 'จุดที่ปรับอัตโนมัติ' }, [ordered]@{ v = "$mins นาที"; l = 'เวลาที่ใช้' })
    changes = $changes; checks = $checks
    files = @([ordered]@{ ext = 'docx'; name = (Split-Path $f -Leaf); path = $f }, [ordered]@{ ext = 'pdf'; name = (Split-Path ([IO.Path]::ChangeExtension($f, '.pdf')) -Leaf); path = [IO.Path]::ChangeExtension($f, '.pdf') }, [ordered]@{ ext = 'txt'; name = (Split-Path $sf -Leaf); path = $sf })
  }
  foreach ($ph in $ST.phases) { $ph.status = 'done' }; $ST.phases[2].sub = $(if ($nf) { "ดู $nf หน้า" } else { 'พร้อมส่ง' })
  $ST.agents.sum = [ordered]@{ on = $false; state = 'เสร็จ'; task = 'ส่งสรุปให้ฮารุแล้ว'; p = 100 }
  Say 'sum' 'prep' $(if ($nf) { "สรุปเสร็จแล้ว มี $nf หน้าที่ควรเปิดดู" } else { 'สรุปเสร็จแล้ว เล่มนี้เรียบร้อย' }) @{ act = 'walk' }
  $ST.status = 'done'; $ST.ended = (NowMs); Save-State
  if ($Loop -and $nf) { Write-Output (powershell -NoProfile -ExecutionPolicy Bypass -File "$S\review_loop.ps1" -State $State | Out-String) }
}
