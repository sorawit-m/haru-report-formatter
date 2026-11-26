param([string]$File)
# รูปในตาราง: ถ้ากว้างกว่าช่อง ให้ย่อลงตามสัดส่วนให้พอดีช่อง (เว้นขอบช่อง)
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
try { $d = $app.Documents.Open($File); $n = 0
  foreach ($s in $d.InlineShapes) { if (-not $s.Range.Information(12)) { continue }
    $cell = $s.Range.Cells.Item(1); $max = $cell.Width - $cell.LeftPadding - $cell.RightPadding - 2
    if ($s.Width -gt $max) { $k = $max / $s.Width; $s.LockAspectRatio = -1; $s.Width = $max; $n++ } }
  "images scaled to fit their cell: $n"; $d.Save(); $d.Close() } finally { $app.Quit() }
