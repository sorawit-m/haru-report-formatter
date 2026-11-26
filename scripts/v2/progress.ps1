# หน้าความคืบหน้า + รายงานผล (HTML รีเฟรชเอง) — Trained by Haru
# dot-source แล้วเรียก: Init-Progress <html> <ชื่อเล่ม> <ชนิด> ; Set-Step <ชื่อ> running|done|skip <รายละเอียด> ; Set-Report <changes.json> <ผลตรวจ> ; Close-Progress
$script:PG = @{ path = ''; title = ''; kind = ''; steps = New-Object System.Collections.ArrayList; report = ''; done = $false; started = Get-Date }
function HtmlEnc([string]$s) { return [System.Net.WebUtility]::HtmlEncode($s) }
function Write-Progress-Html {
  $p = $script:PG; $el = (Get-Date) - $p.started
  $refresh = $(if ($p.done) { '' } else { '<meta http-equiv="refresh" content="2">' })
  $rows = ($p.steps | ForEach-Object {
    $icon = switch ($_.status) { 'done' { '<span class="ok">✓</span>' } 'running' { '<span class="spin">⏳</span>' } 'skip' { '<span class="skip">–</span>' } default { '<span class="wait">•</span>' } }
    "<li class='$($_.status)'>$icon <b>$(HtmlEnc $_.name)</b><div class='d'>$(HtmlEnc $_.detail)</div></li>" }) -join "`n"
  $ndone = @($p.steps | Where-Object { $_.status -in 'done', 'skip' }).Count; $pct = $(if ($p.steps.Count) { [int](100 * $ndone / $p.steps.Count) } else { 0 })
  $status = $(if ($p.done) { 'เสร็จแล้ว ✨' } else { 'กำลังทำงาน…' })
  $html = @"
<!doctype html><html lang="th"><head><meta charset="utf-8">$refresh<title>ฮารุ · จัดเล่ม — $(HtmlEnc $p.title)</title>
<style>
:root{--bg:#f6f8fb;--card:#fff;--ink:#14213d;--mute:#6b7a90;--acc:#0f4c81;--ok:#1a7f37}
body{margin:0;background:var(--bg);font-family:'TH SarabunPSK','Sarabun','Leelawadee UI',sans-serif;color:var(--ink);font-size:20px}
header{background:linear-gradient(120deg,#0f4c81,#1d6fb8);color:#fff;padding:22px 32px}
header h1{margin:0;font-size:34px} header .sub{opacity:.9;font-size:20px}
.wrap{max-width:1100px;margin:22px auto;padding:0 16px}
.card{background:var(--card);border-radius:14px;box-shadow:0 2px 10px #0001;padding:18px 24px;margin-bottom:18px}
.bar{height:12px;background:#e3e9f2;border-radius:8px;overflow:hidden}.bar i{display:block;height:100%;width:$pct%;background:var(--acc);transition:width .5s}
ul.steps{list-style:none;padding:0;margin:0} ul.steps li{padding:8px 4px;border-bottom:1px solid #eef1f6} ul.steps li .d{color:var(--mute);font-size:17px;margin-left:30px}
.ok{color:var(--ok);font-weight:bold}.wait{color:#bbb}.skip{color:#aaa}.spin{display:inline-block;animation:p 1s infinite}@keyframes p{50%{opacity:.3}}
li.running b{color:var(--acc)} table{border-collapse:collapse;width:100%;font-size:18px} td,th{border:1px solid #dfe5ee;padding:5px 8px;vertical-align:top;text-align:left} th{background:#0f4c81;color:#fff}
h2{margin:6px 0 12px;font-size:26px;color:var(--acc)} .pg{white-space:nowrap;font-weight:bold;color:var(--acc)}
footer{text-align:center;color:var(--mute);padding:18px;font-size:18px}
</style></head><body>
<header><h1>ฮารุ · ระบบจัดเล่มรายงาน</h1><div class="sub">$(HtmlEnc $p.title) · ชนิดเล่ม: $(HtmlEnc $p.kind) · $status · เวลา $([int]$el.TotalMinutes) นาที $($el.Seconds) วินาที</div></header>
<div class="wrap">
<div class="card"><div class="bar"><i></i></div><div style="margin-top:6px;color:var(--mute)">ความคืบหน้า $pct% ($ndone/$($p.steps.Count) ขั้นตอน)</div></div>
$($p.report)
<div class="card"><h2>ขั้นตอนการทำงาน</h2><ul class="steps">$rows</ul></div>
</div><footer>Report Formatter · <b>Trained by Haru</b></footer></body></html>
"@
  [IO.File]::WriteAllText($p.path, $html, (New-Object Text.UTF8Encoding $false))
}
function Init-Progress([string]$path, [string]$title, [string]$kind, [string[]]$plan) {
  $script:PG.path = $path; $script:PG.title = $title; $script:PG.kind = $kind; $script:PG.started = Get-Date
  foreach ($s in $plan) { [void]$script:PG.steps.Add(@{ name = $s; status = 'wait'; detail = '' }) }
  Write-Progress-Html; try { Start-Process $path } catch {}
}
function Set-Step([string]$name, [string]$status, [string]$detail = '') {
  $s = $script:PG.steps | Where-Object { $_.name -eq $name } | Select-Object -First 1
  if (-not $s) { $s = @{ name = $name; status = 'wait'; detail = '' }; [void]$script:PG.steps.Add($s) }
  $s.status = $status; if ($detail) { $s.detail = $detail }; Write-Progress-Html
}
function Set-Report([string]$changesJson, [string]$checks, [string]$outFiles) {
  $cats = @(); try { $cats = Get-Content -LiteralPath $changesJson -Raw -Encoding UTF8 | ConvertFrom-Json } catch {}
  $sb = New-Object System.Text.StringBuilder
  [void]$sb.Append("<div class='card'><h2>ไฟล์ผลลัพธ์</h2><div>$(HtmlEnc $outFiles)</div></div>")
  [void]$sb.Append("<div class='card'><h2>แก้อะไรไปบ้าง (ละเอียด)</h2><table><tr><th>หมวด</th><th>จำนวน</th><th>ตัวอย่าง (หน้า PDF / เลขหน้าที่พิมพ์ · ก่อน → หลัง)</th></tr>")
  foreach ($c in $cats) {
    $ex = ($c.examples | ForEach-Object { $pg = $(if ($_.page) { "<span class='pg'>หน้า $($_.page)" + $(if ($_.printed) { " (พิมพ์ $($_.printed))" } else { '' }) + "</span> " } else { '' }); $t = $(if ($_.text) { " — " + (HtmlEnc $_.text) } else { '' }); "<div>$pg$(HtmlEnc $_.before) → $(HtmlEnc $_.after)$t</div>" }) -join ''
    [void]$sb.Append("<tr><td>$(HtmlEnc $c.label)</td><td>$($c.count)</td><td>$ex</td></tr>") }
  [void]$sb.Append("</table></div>")
  [void]$sb.Append("<div class='card'><h2>จุดที่ควรเช็ค</h2><pre style='font-family:inherit;white-space:pre-wrap;margin:0'>$(HtmlEnc $checks)</pre></div>")
  $script:PG.report = $sb.ToString(); Write-Progress-Html
}
function Close-Progress { $script:PG.done = $true; Write-Progress-Html }
