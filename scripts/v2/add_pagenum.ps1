param([string]$File)
& node (Join-Path (Split-Path $MyInvocation.MyCommand.Path) 'one_font.mjs') $File | Out-Null
# เลขหน้า: ทุกรายงานต้องมี — ถ้า section สุดท้าย (เนื้อหา) ยังไม่มีเลขหน้า ใส่ชิดขวา 16pt TH SarabunPSK เริ่ม 1 (ปก/สารบัญไม่มีเลข)
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
try {
  $d = $app.Documents.Open($File); $n = $d.Sections.Count
  $has = $false; for ($i = 1; $i -le $n; $i++) { foreach ($k in 1..3) { try { foreach ($fd in $d.Sections.Item($i).Footers.Item($k).Range.Fields) { if ($fd.Type -eq 33) { $has = $true } } } catch {} } }
  if ($has) { "page numbers already present" } else {
    $s = $d.Sections.Item($n); $ft = $s.Footers.Item(1); $ft.LinkToPrevious = $false
    $r = $ft.Range; $r.Text = ''; $r.Fields.Add($r, 33) | Out-Null
    $ft.Range.ParagraphFormat.Alignment = 2; $ft.Range.Font.Name = 'TH SarabunPSK'; $ft.Range.Font.NameBi = 'TH SarabunPSK'; $ft.Range.Font.Size = 16; $ft.Range.Font.SizeBi = 16
    if ($n -gt 1) { $ft.PageNumbers.RestartNumberingAtSection = $true; $ft.PageNumbers.StartingNumber = 1 }
    foreach ($t in $d.TablesOfContents) { $t.Update(); $t.Range.Font.Size = 16; $t.Range.Font.SizeBi = 16; $t.Range.Font.Italic = $false }
    foreach ($t in $d.TablesOfFigures) { $t.Update(); $t.Range.Font.Size = 16; $t.Range.Font.SizeBi = 16; $t.Range.Font.Italic = $false }
    "page numbers added to section $n (start 1)" }
  # ItalicFix: TOC lines never italic
  foreach ($i in -20..-28) { $st = $d.Styles.Item($i); $st.Font.Italic = $false; $st.Font.ItalicBi = $false }
  foreach ($t in @($d.TablesOfContents) + @($d.TablesOfFigures)) { foreach ($p in $t.Range.Paragraphs) { if ($p.Range.Font.Italic -ne 0) { $p.Range.Font.Italic = $false; $p.Range.Font.ItalicBi = $false } } }
  $d.Save(); $d.Close()
} finally { $app.Quit() }
