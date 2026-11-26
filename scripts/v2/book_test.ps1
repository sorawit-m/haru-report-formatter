param([string]$Key, [string]$Orig, [string]$Widths, [int]$DenseSize = 13)
# เล่มรายการทดสอบระบบ (ตาราง 9 คอลัมน์ + รูป): เตรียมต้นฉบับ -> ขั้นตอนมาตรฐาน -> ขั้นตอนเฉพาะที่เล่ม (3) ผ่านแล้ว -> ตรวจ
$ErrorActionPreference = 'Stop'
$SP = Split-Path $MyInvocation.MyCommand.Path
$D = (Join-Path $env:USERPROFILE 'Downloads')
$pre = Join-Path $D "_pre$Key.docx"; Copy-Item -LiteralPath $Orig $pre -Force
"== 0) remove Test by / Approve by rows, space after รูปที่"
node "$SP\test_report_pre.mjs" $pre
& powershell -NoProfile -ExecutionPolicy Bypass -File "$SP\book_all.ps1" -Key $Key -Orig $pre -Widths $Widths
Remove-Item -LiteralPath $pre -ErrorAction SilentlyContinue
$f = (Get-ChildItem -LiteralPath $D -Filter "($Key) *.docx" | Select-Object -First 1).FullName
"== 9) test-table extras (as book 3)"
node "$SP\dense_fix.mjs" $f ($Widths.Split('=')[1])
node "$SP\steps_list.mjs" $f Step
node "$SP\no_repeat_wide.mjs" $f
node "$SP\compound_breaks.mjs" $f --dense-back
node "$SP\tag_thai.mjs" $f --no-compound
node "$SP\table_cleanup.mjs" $f
node "$SP\keep_heads.mjs" $f
node "$SP\pic_rows.mjs" $f
node "$SP\grid_trim.mjs" $f
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
try { $d = $app.Documents.Open($f); $n = 0
  foreach ($tb in $d.Tables) { $c0 = 0; try { $c0 = $tb.Columns.Count } catch { $c0 = 9 }; if ($c0 -ge 7) { $n++; foreach ($c in $tb.Range.Cells) { if ($c.RowIndex -gt 1) { $c.Range.Font.Size = $DenseSize; $c.Range.Font.SizeBi = $DenseSize } } } }
  "  test tables body at ${DenseSize}pt: $n"
  foreach ($t in $d.TablesOfContents) { $t.Update(); $t.Range.Font.Size = 16; $t.Range.Font.SizeBi = 16; $t.Range.Font.Italic = $false }
  foreach ($t in $d.TablesOfFigures) { $t.Update(); $t.Range.Font.Size = 16; $t.Range.Font.SizeBi = 16; $t.Range.Font.Italic = $false }
  $d.Save(); $d.Close() } finally { $app.Quit() }
& powershell -NoProfile -ExecutionPolicy Bypass -File "$SP\fit_images.ps1" -File $f
"done: $f"
