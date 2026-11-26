param([string]$File, [int]$TestFont = 0, [int]$TestHead = 0, [int]$PicText = 0)
# ขั้นตอนใน Word: สารบัญ 16pt (บรรทัดบทหนา ไม่เอียง), ปก (ระยะชื่อเรื่อง 12pt, บรรทัดโครงการ 1.3, รูปพื้นหลังเต็มหน้า), ฟิลด์, ตัวอักษรตารางทดสอบ
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
try {
  $d = $app.Documents.Open($File)
  foreach ($i in -20..-28) { $s = $d.Styles.Item($i); $s.Font.Size = 16; $s.Font.SizeBi = 16; $s.Font.Italic = $false; $s.Font.ItalicBi = $false; $s.Font.Bold = ($i -eq -20); $s.Font.BoldBi = ($i -eq -20) }
  $first = $true
  foreach ($p in $d.Sections.Item(1).Range.Paragraphs) { $t = $p.Range.Text.Trim(); if ($t.Length -le 1) { continue }
    if ($first) { $p.SpaceBefore = 12; $first = $false }
    if ($t.Contains('โครงการ')) { $p.LineSpacingRule = 5; $p.LineSpacing = 12 * 1.3 } }
  $ps = $d.Sections.Item(1).PageSetup
  foreach ($s in $d.Sections.Item(1).Range.ShapeRange) { if ($s.Width -lt $ps.PageWidth * 0.9) { continue }
    $s.LockAspectRatio = 0; $s.RelativeHorizontalPosition = 1; $s.RelativeVerticalPosition = 1; $s.Left = 0; $s.Top = 0; $s.Width = $ps.PageWidth; $s.Height = $ps.PageHeight }
  $nt = 0; $nd = 0
  foreach ($tb in $d.Tables) { $c0 = 0; try { $c0 = $tb.Columns.Count } catch { $c0 = 9 }
    # หัวตารางทดสอบ: เล่มผลทดสอบ (เนื้อหา 12) ใช้หัว 14 ตามเล่ม 6 ที่ผ่าน — หัวไม่ตก 2 บรรทัด ("Expected Result")
    if ($c0 -ge 7 -and $TestFont -gt 0) { $nd++; foreach ($c in $tb.Range.Cells) { if ($c.RowIndex -gt 1 -and $c.Range.InlineShapes.Count -eq 0) { $c.Range.Font.Size = $TestFont; $c.Range.Font.SizeBi = $TestFont } elseif ($c.RowIndex -gt 1 -and $PicText -gt 0) { $c.Range.Font.Size = $PicText; $c.Range.Font.SizeBi = $PicText } elseif ($c.RowIndex -eq 1 -and ($TestHead -gt 0 -or $TestFont -eq 12)) { $hs = $(if ($TestHead -gt 0) { $TestHead } else { 14 }); $c.Range.Font.Size = $hs; $c.Range.Font.SizeBi = $hs } } }
    elseif ($c0 -gt 1 -and $tb.Range.Font.Size -ne 16) { $tb.Range.Font.Size = 16; $tb.Range.Font.SizeBi = 16; $nt++ } }
  $d.Content.Fields.Update() | Out-Null
  foreach ($t in @($d.TablesOfContents) + @($d.TablesOfFigures)) { $t.Update(); $t.Range.Font.Size = 16; $t.Range.Font.SizeBi = 16; $t.Range.Font.Italic = $false }
  "tables set to 16pt: $nt; test tables at ${TestFont}pt: $nd"
  $d.Save(); $d.Close()
} finally { $app.Quit() }
