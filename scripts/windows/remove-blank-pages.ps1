param([string]$File, [string]$Pdf = "")
function Log($m) { Write-Output ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $m) }
$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0; $w.ScreenUpdating = $false
$opt = $w.Options; $oldSpell = $opt.CheckSpellingAsYouType; $oldGram = $opt.CheckGrammarAsYouType
$opt.CheckSpellingAsYouType = $false; $opt.CheckGrammarAsYouType = $false
$d = $w.Documents.Open($File, $false, $false)
$total = 0
for ($pass = 0; $pass -lt 3; $pass++) {
  $d.Repaginate()
  $n = $d.ComputeStatistics(2); $fixed = 0
  for ($pg = $n; $pg -ge 2; $pg--) {
    $s = $d.GoTo(1, 1, $pg).Start
    $e = if ($pg -lt $n) { $d.GoTo(1, 1, $pg + 1).Start } else { $d.Content.End }
    if ($e -le $s) { continue }
    $r = $d.Range($s, $e)
    if ($r.Text -notmatch '^[\r\s\f]*$' -or $r.Tables.Count -gt 0 -or $r.InlineShapes.Count -gt 0 -or $r.ShapeRange.Count -gt 0) { continue }
    if ($r.Sections.Count -gt 1 -or $r.Text -match '\f') { continue }
    $prevIsTable = $false; $nextIsTable = $false
    try { $prevIsTable = $d.Range($s - 1, $s - 1).Information(12) } catch {}
    try { $nextIsTable = $d.Range($e, $e).Information(12) } catch {}
    if ($prevIsTable -and $nextIsTable) {
      # must keep one paragraph so the tables do not merge: make it 1pt high
      $r.Font.Size = 1; $r.Font.SizeBi = 1; $r.ParagraphFormat.SpaceBefore = 0; $r.ParagraphFormat.SpaceAfter = 0
      $r.ParagraphFormat.LineSpacingRule = 4; $r.ParagraphFormat.LineSpacing = 1
    } elseif ($e -lt $d.Content.End) {
      $r.Delete() | Out-Null
    }
    $fixed++
  }
  $total += $fixed
  Log "pass $pass pages=$n blank pages fixed=$fixed"
  if ($fixed -eq 0) { break }
}
$d.Repaginate()
foreach ($toc in $d.TablesOfContents) { $toc.Update() }; foreach ($tof in $d.TablesOfFigures) { $tof.Update() }
$d.Repaginate(); foreach ($toc in $d.TablesOfContents) { $toc.UpdatePageNumbers() }
foreach ($tof in $d.TablesOfFigures) { $tof.UpdatePageNumbers() }
Log ("final pages=" + $d.ComputeStatistics(2) + " blank fixed=$total")
$d.Save()
if ($Pdf) { $d.ExportAsFixedFormat($Pdf, 17) }
$d.Close($false)
$opt.CheckSpellingAsYouType = $oldSpell; $opt.CheckGrammarAsYouType = $oldGram; $w.ScreenUpdating = $true
$w.Quit()
