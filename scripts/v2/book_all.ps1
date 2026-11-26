param([string]$Key, [string]$Orig, [string]$Widths = '')
# ขั้นตอนเต็มที่เล่ม (1)(2) ผ่านแล้ว — เริ่มจากต้นฉบับทุกครั้ง
$ErrorActionPreference = 'Stop'
$SP = Split-Path $MyInvocation.MyCommand.Path
$D = (Join-Path $env:USERPROFILE 'Downloads')
$f = (Get-ChildItem -LiteralPath $D -Filter "($Key) *.docx" | Select-Object -First 1).FullName
New-Item -ItemType Directory -Force "$SP\broken" | Out-Null
Copy-Item -LiteralPath $f "$SP\broken\" -Force
Copy-Item -LiteralPath $Orig $f -Force
"== 1) original, cover size, remove intro"
node "$SP\cover_size_fix.mjs" $f
& powershell -NoProfile -ExecutionPolicy Bypass -File "$SP\cut_intro.ps1" -Only $Key
node "$SP\renum.mjs" $f
"== 2) table cells styled as headings"
node "$SP\fix_heading_cells.mjs" $f
"== 3) text flow + break points"
node "$SP\text_flow.mjs" $f
node "$SP\more_breaks.mjs" $f
node "$SP\camel_breaks.mjs" $f
node "$SP\tag_thai.mjs" $f --no-compound
"== 4) one font, 16pt content, cover black"
node "$SP\one_font.mjs" $f
node "$SP\content_size.mjs" $f 16
node "$SP\cover_black.mjs" $f
"== 5) table widths"
node "$SP\grid_trim.mjs" $f
if ($Widths) { $wa = $Widths -split ';'; node "$SP\set_widths2.mjs" $f @wa }
node "$SP\fit_tables.mjs" $f
"== 6) English-heavy justify, spacing paragraphs"
node "$SP\latin_justify.mjs" $f
node "$SP\mixed_back.mjs" $f
node "$SP\drop_after_table.mjs" $f
node "$SP\thin_separators.mjs" $f
node "$SP\toc_title_nonum.mjs" $f
node "$SP\table_cleanup.mjs" $f
"== 7) Word: table font, TOC1 bold, cover spacing + picture, fields"
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
try {
  $d = $app.Documents.Open($f)
  $n16 = 0; $n14 = 0
  foreach ($t in $d.Tables) { $cols = 0; try { $cols = $t.Rows.Item(1).Cells.Count } catch {}; $want = $(if ($cols -ge 7) { 14 } else { 16 })
    if ($t.Range.Font.Size -ne $want) { $t.Range.Font.Size = $want; $t.Range.Font.SizeBi = $want }; if ($want -eq 14) { $n14++ } else { $n16++ } }
  "  tables 16pt: $n16, wide tables (7+ columns) 14pt: $n14"
  foreach ($i in -20..-28) { $ts = $d.Styles.Item($i); $ts.Font.Size = 16; $ts.Font.SizeBi = 16 }
  $s = $d.Styles.Item(-20); $s.Font.Bold = $true; $s.Font.BoldBi = $true
  $first = $true
  foreach ($p in $d.Sections.Item(1).Range.Paragraphs) { $t = $p.Range.Text.Trim(); if ($t.Length -le 1) { continue }
    if ($first) { $p.SpaceBefore = 12; $first = $false }
    if ($t.Contains('โครงการ')) { $p.LineSpacingRule = 5; $p.LineSpacing = 12 * 1.3 } }
  $ps = $d.Sections.Item(1).PageSetup
  foreach ($s in $d.Sections.Item(1).Range.ShapeRange) { if ($s.Width -lt $ps.PageWidth * 0.9) { continue }
    $s.LockAspectRatio = 0; $s.RelativeHorizontalPosition = 1; $s.RelativeVerticalPosition = 1; $s.Left = 0; $s.Top = 0; $s.Width = $ps.PageWidth; $s.Height = $ps.PageHeight; "  cover picture pinned to page" }
  $d.Content.Fields.Update() | Out-Null
  foreach ($t in $d.TablesOfContents) { $t.Update(); $t.Range.Font.Size = 16; $t.Range.Font.SizeBi = 16 }
  $d.Save(); $d.Close()
} finally { $app.Quit() }
& powershell -NoProfile -ExecutionPolicy Bypass -File "$SP\add_pagenum.ps1" -File $f
"== 8) export + check every page"
& powershell -NoProfile -ExecutionPolicy Bypass -File "$SP\export_check.ps1" -File $f -Out (Join-Path $env:TEMP "rf_check\all$Key")
"done: $f"
