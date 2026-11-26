param([string]$File, [double]$MaxHeightShare = 0.62)
# รูปในตาราง: ขยายให้เต็มความกว้างช่อง (ชัดขึ้น) แต่สูงไม่เกิน MaxHeightShare ของพื้นที่ข้อความในหน้า (ให้รูปกับข้อความในแถวอยู่หน้าเดียวกันได้)
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
try {
  $d = $app.Documents.Open($File); $n = 0; $grow = 0
  foreach ($s in $d.InlineShapes) {
    if (-not $s.Range.Information(12)) { continue }
    $c = $s.Range.Cells.Item(1); $ps = $s.Range.Sections.Item(1).PageSetup
    $maxW = $c.Width - $c.LeftPadding - $c.RightPadding - 4
    $maxH = ($ps.PageHeight - $ps.TopMargin - $ps.BottomMargin) * $MaxHeightShare
    if ($s.Width -le 0 -or $s.Height -le 0) { continue }
    $k = [math]::Min($maxW / $s.Width, $maxH / $s.Height)
    if ([math]::Abs($k - 1) -lt 0.02) { continue }
    $old = $s.Width; $s.LockAspectRatio = -1; $s.Width = $s.Width * $k; $n++; if ($s.Width -gt $old) { $grow++ }
  }
  "images resized: $n (enlarged: $grow)"
  $d.Save(); $d.Close()
} finally { $app.Quit() }
