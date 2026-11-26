param([string]$File)
# ตั้งความกว้างคอลัมน์ (cm) ให้ตารางที่หัวเหมือนกันทุกตารางเท่ากัน — ทำทีละแถว (แถวที่รวมช่อง = กว้างรวม)
$specs = @(
    @{ h = 'Field|Type|Required|คำอธิบาย'; w = @(5.5, 3.2, 2.0, 5.3) },
  @{ h = 'Field|Type|คำอธิบาย'; w = @(5.5, 3.2, 7.3) }
)
$cm = 28.3465
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
try {
  $d = $app.Documents.Open($File)
  $ps = $d.Sections.Item($d.Sections.Count).PageSetup; "text width {0:N2} cm" -f (($ps.PageWidth - $ps.LeftMargin - $ps.RightMargin) / $cm)
  $count = @{}; $skipped = 0
  foreach ($t in $d.Tables) {
    $c1 = $t.Rows.Item(1).Cells; $hs = @(); foreach ($c in $c1) { $hs += (($c.Range.Text -replace "[\r\a​]", '').Trim()) }
    $key = $hs -join '|'
    $spec = $specs | Where-Object { $_.h -eq $key } | Select-Object -First 1
    if (-not $spec) { continue }
    $W = $spec.w; $total = ($W | Measure-Object -Sum).Sum
    $t.AllowAutoFit = $false
    $t.Rows.LeftIndent = 0
    foreach ($r in $t.Rows) {
      $n = $r.Cells.Count
      if ($n -eq $W.Count) { for ($i = 1; $i -le $n; $i++) { $r.Cells.Item($i).Width = $W[$i - 1] * $cm } }
      elseif ($n -eq 1) { $r.Cells.Item(1).Width = $total * $cm }
      else { $skipped++ }
    }
    $count[$key] = 1 + [int]$count[$key]
  }
  $count.GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value) tables" }
  "rows skipped (unexpected cell count): $skipped"
  $d.Save(); $d.Close()
} finally { $app.Quit() }
