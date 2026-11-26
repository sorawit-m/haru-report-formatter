param([string]$File, [double]$StartShare = 0.58, [double]$MinWidthPt = 330)
# ขยายรูปในตารางทดสอบให้ใหญ่ที่สุดเท่าที่กลุ่ม (แถวทดสอบ + ผลที่พบ + รูป) ยังอยู่หน้าเดียวกันได้
# เริ่มที่ความสูง StartShare ของพื้นที่ข้อความ แล้วย่อเฉพาะรูปที่ทำให้กลุ่มแตกหน้า ทีละ 10%
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
try {
  $d = $app.Documents.Open($File)
  $pics = @(); foreach ($s in $d.InlineShapes) { if ($s.Range.Information(12)) { $pics += $s } }
  foreach ($s in $pics) {
    $c = $s.Range.Cells.Item(1); $ps = $s.Range.Sections.Item(1).PageSetup
    $maxW = $c.Width - $c.LeftPadding - $c.RightPadding - 4; $maxH = ($ps.PageHeight - $ps.TopMargin - $ps.BottomMargin) * $StartShare
    $k = [math]::Min($maxW / $s.Width, $maxH / $s.Height); $s.LockAspectRatio = -1; $s.Width = $s.Width * $k
  }
  for ($round = 1; $round -le 8; $round++) {
    $d.Repaginate(); $shrunk = 0
    foreach ($s in $pics) {
      $row = $s.Range.Rows.Item(1); $tb = $s.Range.Tables.Item(1); $ri = $row.Index
      if ($ri -lt 3) { continue }
      $first = $tb.Rows.Item($ri - 2).Range; $pgA = $first.Information(3); $pgB = $s.Range.Information(3)
      if ($pgA -ne $pgB -and $s.Width -gt $MinWidthPt) { $s.Width = [math]::Max($MinWidthPt, $s.Width * 0.9); $shrunk++ }
    }
    "round ${round}: pictures shrunk so their test case fits one page: $shrunk"
    if ($shrunk -eq 0) { break }
  }
  $w = $pics | ForEach-Object { $_.Width }; "final picture width avg {0:N0} min {1:N0} max {2:N0}" -f ($w | Measure -Average).Average, ($w | Measure -Min).Minimum, ($w | Measure -Max).Maximum
  $d.Save(); $d.Close()
} finally { $app.Quit() }
