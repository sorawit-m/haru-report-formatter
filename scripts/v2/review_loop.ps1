param([Parameter(Mandatory = $true)][string]$State, [int]$MaxRounds = 3)
# วงตรวจ-แก้หลังจัดเล่ม (Trained by Haru): ผู้ตรวจพบหน้าไหน -> ผู้แก้ (Claude Opus) แก้หน้านั้นทันที -> ส่งออก PDF ใหม่ -> ผู้ตรวจอ่านซ้ำ -> วนจนไม่เจอ (สูงสุด MaxRounds รอบ)
# อ่าน/เขียน <State>\state.json ให้แอปฮารุแสดงผลสด ; ผู้แก้เรียก `claude -p --model opus` ด้วยบัญชีที่ล็อกอินในเครื่อง
$ErrorActionPreference = 'Continue'
$S = Split-Path $MyInvocation.MyCommand.Path
$W = Join-Path (Split-Path $S) 'windows\pdf-to-png.ps1'
$OutputEncoding = New-Object Text.UTF8Encoding $false; [Console]::OutputEncoding = New-Object Text.UTF8Encoding $false
$sj = Join-Path $State 'state.json'
function NowMs { [int64]([DateTimeOffset]::Now.ToUnixTimeMilliseconds()) }
function Load { Get-Content -LiteralPath $sj -Raw -Encoding UTF8 | ConvertFrom-Json }
$ST = Load
function Save { $j = $ST | ConvertTo-Json -Depth 14 -Compress; $tmp = "$sj.tmp"; [IO.File]::WriteAllText($tmp, $j, (New-Object Text.UTF8Encoding $false)); Move-Item -LiteralPath $tmp $sj -Force; [IO.File]::WriteAllText((Join-Path $State 'state.js'), "window.HARU_STATE&&window.HARU_STATE($j);", (New-Object Text.UTF8Encoding $false)) }
function Say([string]$from, [string]$to, [string]$text, [hashtable]$extra = @{}) { $m = [ordered]@{ from = $from; to = $to; text = $text; t = (Get-Date -Format 'HH:mm'); ts = (NowMs) }; foreach ($k in $extra.Keys) { $m[$k] = $extra[$k] }; $ST.feed = @($ST.feed) + ([pscustomobject]$m); Save }
function SetP($obj, [string]$name, $val) { if ($obj.PSObject.Properties[$name]) { $obj.$name = $val } else { $obj | Add-Member -NotePropertyName $name -NotePropertyValue $val -Force } }
$docx = $ST.files.docx; $pdf = $ST.files.pdf; $total = [int]$ST.pages.total; $q = [math]::Max(1, [math]::Ceiling($total / 4))
$shotDir = Join-Path $State 'shots'; New-Item -ItemType Directory -Force $shotDir, (Join-Path $shotDir 'after') | Out-Null
$ST.phases = @($ST.phases) + ([pscustomobject][ordered]@{ name = 'ตรวจ-แก้วนจนผ่าน'; status = 'active'; sub = 'เริ่ม'; frac = 0 })
$ST.status = 'running'; SetP $ST 'ended' $null; Save
$lab = [ordered]@{ 'blank' = 'หน้าว่าง'; 'heading/table header at page end' = 'หัวข้อหรือหัวตารางค้างท้ายหน้า'; 'past right margin' = 'มีของล้นขอบขวา'; 'starts lower than normal' = 'เนื้อหาเริ่มต่ำกว่าหน้าอื่น' }
for ($round = 1; $round -le $MaxRounds; $round++) {
  $issues = @($ST.pages.map.PSObject.Properties | Where-Object { $_.Value -eq 'issue' } | ForEach-Object { [int]$_.Name } | Sort-Object)
  if (-not $issues.Count) { break }
  $ST.round = $round + 1; $ph = $ST.phases[-1]; $ph.sub = "รอบ $round · แก้ $($issues.Count) หน้า"
  $ST.agents.fix = [pscustomobject][ordered]@{ on = $true; state = "แก้รอบ $round"; task = 'เปิดไฟล์ Word'; p = 0 }; Save
  Say 'fix' 'rev' "รับ $($issues.Count) หน้า เริ่มแก้ทีละหน้าเลย"
  $i = 0
  foreach ($n in $issues) {
    $i++; $why = [string]$ST.pages.notes."$n"; $lane = [math]::Min(3, [math]::Floor(($n - 1) / $q))
    $ST.agents.fix.task = "หน้า ${n}: $why"; $ST.agents.fix.p = [int](100 * ($i - 1) / $issues.Count); $ph.frac = ($i - 1) / $issues.Count * .7; Save
    $prompt = @"
คุณคือผู้แก้ของทีมฮารุ (สกิล report-formatter ที่ ~/.claude/skills/report-formatter) แก้ไฟล์ "$docx"
ปัญหาที่ผู้ตรวจพบ: หน้า $n ของ PDF "$pdf" — $why
ทำ: ดูภาพหน้า $n (scripts/windows/pdf-to-png.ps1 -From $n -To $n) หาว่าเกิดจากอะไรในไฟล์ Word แล้วแก้ให้หาย โดยใช้วิธีตามบทเรียนของสกิล (keep_heads, fit_tables, cells_top ฯลฯ)
ห้าม: เปลี่ยนเนื้อหา/คำ, ใส่เครดิตในไฟล์, แตะต้นฉบับ, ส่งออก PDF (ระบบส่งออกเองหลังแก้ครบ)
ตอบบรรทัดเดียวภาษาไทย ว่าแก้อะไร (ถ้าแก้ไม่ได้ ให้บอกเหตุผลสั้นๆ)
"@
    $out = ($prompt | claude -p --model opus --output-format json --permission-mode acceptEdits --allowedTools "Read,Glob,Grep,Edit,Write,Skill,Bash,PowerShell" 2>&1 | Out-String)
    $res = ''; try { $res = ([string](($out | ConvertFrom-Json).result)).Trim() } catch { $res = ($out -replace '\s+', ' ').Trim() }
    if (-not $res) { $res = 'แก้แล้ว' }; if ($res.Length -gt 200) { $res = $res.Substring(0, 199) + '…' }
    $ST.pages.map."$n" = 'fixed'; SetP $ST.shots."$n" 'fix' $res
    Say 'fix' 'rev' $res @{ tag = 'fixed'; pages = @($n); lane = $lane }
  }
  # ส่งออกใหม่ + ผู้ตรวจอ่านซ้ำ
  $ST.agents.fix.task = 'ส่งออก PDF ใหม่ให้ผู้ตรวจ'; $ST.agents.fix.p = 100; $ph.sub = "รอบ $round · ส่งออก PDF"; $ph.frac = .75; Save
  $chk = Join-Path $env:TEMP ("rf_loop\" + [guid]::NewGuid().ToString('N').Substring(0, 8))
  powershell -NoProfile -ExecutionPolicy Bypass -File "$S\export_check.ps1" -File $docx -Out $chk | Out-Null
  $ST.agents.fix = [pscustomobject][ordered]@{ on = $false; state = 'รอผลตรวจซ้ำ'; task = "แก้รอบ $round ครบ $($issues.Count) หน้า"; p = 100 }
  $ST.agents.rev.on = $true; $ST.agents.rev.state = "ตรวจซ้ำรอบ $round"; $ST.agents.rev.task = 'อ่านซ้ำหน้าที่แก้ และทุกหน้าที่อาจเลื่อน'; $ph.sub = "รอบ $round · ตรวจซ้ำ"; Save
  Say 'fix' 'rev' "แก้ครบแล้ว PDF ใหม่มาแล้ว ช่วยตรวจซ้ำหน่อย" @{ act = 'walk' }
  $pc = (powershell -NoProfile -ExecutionPolicy Bypass -File "$S\pdf_check.ps1" -Pdf $pdf -Out $chk | Out-String)
  $newTotal = [int](([regex]::Match($pc, 'pages:\s*(\d+)')).Groups[1].Value); if ($newTotal) { $total = $newTotal; $ST.pages.total = $total }
  $flag = [ordered]@{}
  foreach ($k in $lab.Keys) { $m = [regex]::Match($pc, [regex]::Escape($k) + ':\s*([0-9,]*)'); if (-not $m.Success) { continue }; foreach ($p in ($m.Groups[1].Value -split ',' | Where-Object { $_ })) { if (-not $flag.Contains($p)) { $flag[$p] = @() }; $flag[$p] += $lab[$k] } }
  foreach ($n in $issues) {
    $lane = [math]::Min(3, [math]::Floor(($n - 1) / $q)); $png = Join-Path $shotDir ("after\p{0:D3}_r$round.png" -f $n)
    powershell -NoProfile -ExecutionPolicy Bypass -File $W -Pdf $pdf -OutDir (Join-Path $shotDir 'after') -Width 1100 -From $n -To $n | Out-Null
    $raw = Join-Path $shotDir ('after\p{0:D3}.png' -f $n); if (Test-Path $raw) { Move-Item $raw $png -Force }
    SetP $ST.shots."$n" 'after' $png
    if ($flag.Contains("$n")) { $ST.pages.map."$n" = 'issue'; $ST.pages.notes."$n" = ($flag["$n"] -join ' · '); Say 'rev' 'fix' ("ยังไม่หาย: " + $ST.pages.notes."$n") @{ tag = 'issue'; pages = @($n); lane = $lane } }
    else { $ST.pages.map."$n" = 'pass'; Say 'rev' 'fix' "หน้า $n ผ่านแล้ว" @{ tag = 'ok'; pages = @($n); lane = $lane } }
  }
  # หน้าใหม่ที่เพิ่งมีปัญหา (เช่นเนื้อหาเลื่อน)
  foreach ($p in $flag.Keys) { $n = [int]$p; if ($issues -contains $n) { continue }; if ($ST.pages.map."$p" -in 'issue', 'pass') { continue }
    $lane = [math]::Min(3, [math]::Floor(($n - 1) / $q)); powershell -NoProfile -ExecutionPolicy Bypass -File $W -Pdf $pdf -OutDir $shotDir -Width 1100 -From $n -To $n | Out-Null
    SetP $ST.pages.map $p 'issue'; SetP $ST.pages.notes $p ($flag[$p] -join ' · '); SetP $ST.shots $p ([pscustomobject][ordered]@{ before = (Join-Path $shotDir ('p{0:D3}.png' -f $n)); issue = ($flag[$p] -join ' · ') })
    Say 'rev' 'fix' ($flag[$p] -join ' · ') @{ tag = 'issue'; pages = @($n); lane = $lane } }
  $ST.agents.rev.on = $false; $ph.frac = [math]::Min(.95, $round / $MaxRounds); Save
}
# สรุปใหม่
$left = @($ST.pages.map.PSObject.Properties | Where-Object { $_.Value -eq 'issue' } | ForEach-Object { [int]$_.Name } | Sort-Object)
$fixedN = @($ST.pages.map.PSObject.Properties | Where-Object { $_.Value -eq 'pass' }).Count
$ph = $ST.phases[-1]; $ph.status = 'done'; $ph.sub = $(if ($left.Count) { "เหลือ $($left.Count) หน้า" } else { 'ผ่านทุกหน้า' })
if ($ST.summary) {
  $ST.summary.checks = @($left | ForEach-Object { [pscustomobject][ordered]@{ p = $_; t = [string]$ST.pages.notes."$_" } })
  $ST.summary.verdict = $(if ($left.Count) { "ดู $($left.Count) หน้า" } else { 'ผ่าน' })
  $ST.summary.headline = $(if ($left.Count) { "แก้ไป $fixedN หน้า เหลือ $($left.Count) หน้าที่ควรเปิดดู" } else { "แก้ครบ $fixedN หน้า ตรวจซ้ำผ่านทุกหน้า" })
}
$ST.agents.fix = [pscustomobject][ordered]@{ on = $false; state = 'เสร็จ'; task = "แก้ไป $fixedN หน้า"; p = 100 }
Say 'rev' 'prep' $(if ($left.Count) { "จบวงตรวจ-แก้ เหลือ $($left.Count) หน้าที่ต้องให้คนตัดสิน" } else { 'ตรวจซ้ำผ่านทุกหน้าแล้ว!' }) @{ tag = $(if ($left.Count) { 'issue' } else { 'ok' }) }
$ST.status = 'done'; SetP $ST 'ended' (NowMs); Save
"loop done: fixed $fixedN, left $($left.Count)"
