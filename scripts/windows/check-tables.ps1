param([string[]]$Files)
$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
foreach ($file in $Files) {
  $d = $w.Documents.Open($file, $false, $true)
  $d.Repaginate()
  $pages = $d.ComputeStatistics(2)
  $tc = $d.Tables.Count; $span = 0; $spanNoHdr = 0; $list = @()
  for ($t = 1; $t -le $tc; $t++) {
    $tb = $d.Tables.Item($t)
    $p1 = $tb.Range.Characters.First.Information(3); $p2 = $tb.Range.Characters.Last.Information(3)
    if ($p2 -gt $p1) {
      $span++
      $hf = 0; try { $hf = $tb.Range.Cells.Item(1).Range.Rows.HeadingFormat } catch {}
      if ($hf -eq 0) {
        $spanNoHdr++
        if ($list.Count -lt 8) { $list += "p$p1 '" + (($tb.Range.Cells.Item(1).Range.Text -replace '[\r\a]', ' ').Trim()) + "'" }
      }
    }
  }
  "== $([IO.Path]::GetFileName($file)) pages=$pages tables=$tc spanPages=$spanNoHdr/$span without repeated header"
  $list | % { "     $_" }
  $d.Close($false)
}
$w.Quit()
