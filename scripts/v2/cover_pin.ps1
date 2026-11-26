param([string]$File)
# รูปพื้นหลังปก: ยึดกับหน้ากระดาษ ชิดมุมซ้ายบน เต็มหน้า (ไม่ขยับตามย่อหน้า)
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
try { $d = $app.Documents.Open($File); $ps = $d.Sections.Item(1).PageSetup
  foreach ($s in $d.Sections.Item(1).Range.ShapeRange) {
    if ($s.Width -lt $ps.PageWidth * 0.9) { continue }
    $s.LockAspectRatio = 0
    $s.RelativeHorizontalPosition = 1; $s.RelativeVerticalPosition = 1
    $s.Left = 0; $s.Top = 0; $s.Width = $ps.PageWidth; $s.Height = $ps.PageHeight
    "cover picture pinned: {0} -> page, {1:N0}x{2:N0}" -f $s.Name, $s.Width, $s.Height }
  $d.Save(); $d.Close() } finally { $app.Quit() }
