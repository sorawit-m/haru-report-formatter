param([string]$File)
# ตรวจตามรายการปัญหาเดิมทั้งหมด (ฝั่ง Word) — พิมพ์ OK / FAIL ทีละข้อ
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
try {
  $d = $app.Documents.Open($File, $false, $true)
  $res = [ordered]@{}
  # page numbers
  $last = $d.Sections.Item($d.Sections.Count); $pg = $false; foreach ($i in 1..3) { try { if ($last.Footers.Item($i).Range.Fields.Count -gt 0 -or $last.Footers.Item($i).PageNumbers.Count -gt 0) { $pg = $true } } catch {} }
  $res['page numbers in body footer'] = $pg
  # TOC
  if ($d.TablesOfContents.Count) { $toc = $d.TablesOfContents.Item(1).Range; $bs = 0; $bi = 0; foreach ($tp in $toc.Paragraphs) { if ($tp.Range.Font.Size -ne 16) { $bs++ }; if ($tp.Range.Font.Italic -ne 0) { $bi++ } }; $res['TOC 16pt'] = ($bs -eq 0); $res['TOC not italic'] = ($bi -eq 0)
    $p = $toc.Paragraphs.Item(1).Previous(); $title = $null; for ($k = 0; $k -lt 3 -and $p; $k++) { if ($p.Range.Text.Trim() -eq 'สารบัญ') { $title = $p; break }; $p = $p.Previous() }
    $res['TOC title has no number'] = ($title -and $title.Range.ListFormat.ListType -eq 0) } else { $res['TOC present'] = $false }
  # tables
  $nb = 0; $ind = 0; $over = 0; $ps = $last.PageSetup; $tw = $ps.PageWidth - $ps.LeftMargin - $ps.RightMargin
  foreach ($t in $d.Tables) { $cols = 0; try { $cols = $t.Columns.Count } catch { $cols = 2 }
    if ($cols -gt 1) { foreach ($c in $t.Range.Cells) { if ($c.RowIndex -ne 1) { break }; if ($c.Shading.BackgroundPatternColor -ne -16777216 -and $c.Shading.BackgroundPatternColor -ne 16777215 -and $c.Range.Text.Trim([char]13,[char]7,' ').Length -gt 0 -and $c.Range.Font.Bold -ne -1) { $nb++ } }
      foreach ($p in $t.Range.Paragraphs) { if ($p.Range.InlineShapes.Count -eq 0 -and $p.Range.Text.Trim([char]13,[char]7,' ').Length -gt 0 -and $p.FirstLineIndent -gt 0.5) { $ind++ } } }
    try { if ($t.Rows.LeftIndent -lt -2) { $over++ } } catch {} }
  $res['header rows bold'] = ($nb -eq 0); $res["  (non-bold header cells: $nb)"] = $true
  $res['no first-line indent in tables'] = ($ind -eq 0); $res["  (indented table paragraphs: $ind)"] = $true
  $res['no table into left margin'] = ($over -eq 0)
  # fonts
  $odd = @{}; foreach ($s in $d.StoryRanges) { if ($s.StoryType -ne 1) { continue } }
  $res['body font TH SarabunPSK'] = ($d.Styles.Item(-1).Font.Name -eq 'TH SarabunPSK')
  $res['body 16pt'] = ($d.Styles.Item(-1).Font.Size -eq 16)
  $res.GetEnumerator() | ForEach-Object { "{0,-34} {1}" -f $_.Key, $(if ($_.Key.StartsWith('  ')) { '' } elseif ($_.Value) { 'OK' } else { 'FAIL' }) }
  $d.Close($false)
} finally { $app.Quit() }
