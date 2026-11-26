param([string]$File, [double]$MaxHeightShare = 0.6)
# ย่อหน้าที่มีหลายรูปเรียงกัน: ย่อให้วางข้างกันได้ในบรรทัดเดียว (กว้างรวม <= ช่อง) และสูงไม่เกิน MaxHeightShare ของหน้า
# กันไม่ให้รูปตัดขึ้นบรรทัดใหม่จนแถวสูงเกินหน้า (ทำให้ "ผลที่พบ" แยกจากรูป)
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
try {
  $d = $app.Documents.Open($File); $n = 0
  foreach ($t in $d.Tables) { foreach ($p in $t.Range.Paragraphs) {
    $cnt = $p.Range.InlineShapes.Count; if ($cnt -lt 2) { continue }
    $c = $p.Range.Cells.Item(1); $ps = $p.Range.Sections.Item(1).PageSetup
    $avail = $c.Width - $c.LeftPadding - $c.RightPadding - 6 * $cnt; $maxH = ($ps.PageHeight - $ps.TopMargin - $ps.BottomMargin) * $MaxHeightShare
    # common height h: sum(w_i/h_i * h) <= avail
    $ratio = 0; foreach ($s in $p.Range.InlineShapes) { $ratio += $s.Width / $s.Height }
    $h = [math]::Min($maxH, $avail / $ratio)
    foreach ($s in $p.Range.InlineShapes) { $s.LockAspectRatio = -1; $s.Height = $h }
    $n++ } }
  "multi-picture paragraphs fitted on one line: $n"; $d.Save(); $d.Close()
} finally { $app.Quit() }
