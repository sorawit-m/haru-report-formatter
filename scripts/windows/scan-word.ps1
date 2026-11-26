param([string[]]$Files)
$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0
foreach ($file in $Files) {
  $d = $w.Documents.Open($file, $false, $true)
  "== $([IO.Path]::GetFileName($file))"
  # coloured (red-ish) text and highlights
  foreach ($col in 238, 255, 192) {
    $r = $d.Content; $f = $r.Find; $f.ClearFormatting(); $f.Text = ""; $f.Format = $true; $f.Font.Color = $col; $f.Wrap = 0
    $n = 0
    while ($f.Execute() -and $n -lt 15) { $n++; "  colour $col p$($r.Information(3)): " + ($r.Text -replace '[\r\a\v]', ' ').Substring(0, [Math]::Min(90, $r.Text.Length)); $r.Collapse(0) }
  }
  $r = $d.Content; $f = $r.Find; $f.ClearFormatting(); $f.Text = ""; $f.Format = $true; $f.Highlight = $true; $f.Wrap = 0
  $n = 0; while ($f.Execute() -and $n -lt 10) { $n++; "  highlight p$($r.Information(3)): " + ($r.Text -replace '[\r\a\v]', ' ').Substring(0, [Math]::Min(80, $r.Text.Length)); $r.Collapse(0) }
  # empty list items
  $e = 0
  foreach ($p in $d.ListParagraphs) {
    if (($p.Range.Text -replace '[\r\a\s]', '') -eq '' -and $p.Range.InlineShapes.Count -eq 0) { $e++; if ($e -le 10) { "  empty list item p$($p.Range.Information(3)) inTable=$($p.Range.Information(12))" } }
  }
  "  empty list items total=$e"
  # comments / revisions left in
  "  comments=$($d.Comments.Count) revisions=$($d.Revisions.Count)"
  $d.Close($false)
}
$w.Quit()
