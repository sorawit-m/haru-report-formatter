param([string]$File)
# กล่องโค้ด (ตาราง 1 คอลัมน์): ถ้าหัวกล่องอยู่คนละหน้ากับบรรทัดถัดไป (หัวค้างท้ายหน้า) -> ย่อหน้าก่อนกล่องขึ้นหน้าใหม่ (กล่องตามไปทั้งก้อน)
# ไม่ใช้ keepNext ในกล่อง เพราะทำให้ Word ดันทั้งกล่องไปหน้าถัดไปจนเหลือหน้าว่างครึ่งหน้า
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
try {
  $d = $app.Documents.Open($File)
  foreach ($t in $d.Tables) { if ($t.Columns.Count -ne 1) { continue }; foreach ($p in $t.Range.Paragraphs) { if ($p.KeepWithNext) { $p.KeepWithNext = $false } } }
  for ($round = 1; $round -le 4; $round++) {
    $d.Repaginate(); $fixed = 0
    foreach ($t in $d.Tables) {
      if ($t.Columns.Count -ne 1 -or $t.Range.Paragraphs.Count -lt 3) { continue }
      $p1 = $t.Range.Paragraphs.Item(1); $p2 = $t.Range.Paragraphs.Item(2)
      if ($p1.Range.Information(3) -eq $p2.Range.Information(3)) { continue }
      $before = $t.Range.Paragraphs.Item(1).Previous()
      if ($before -and -not $before.Range.Information(12) -and -not $before.PageBreakBefore) { $before.PageBreakBefore = $true; $fixed++ }
    }
    "round ${round}: stranded box headings moved: $fixed"
    if ($fixed -eq 0) { break }
  }
  $d.Save(); $d.Close()
} finally { $app.Quit() }
