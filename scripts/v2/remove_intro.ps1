param([string]$File)
# ลบบท "บทนำ" (หัวข้อระดับ 1) ถึงก่อนบทถัดไป คงการเริ่มเลขหน้าใหม่ของ section
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0; $app.Options.Pagination = $false
function Find-Lvl1($doc, $startPos, [string]$must) { $p = $doc.Range($startPos, $startPos).Paragraphs.Item(1); while ($p -ne $null) { if ($p.OutlineLevel -eq 1) { $t = $p.Range.Text.Trim(); if ($t -and (-not $must -or $t.Contains($must))) { return $p } }; $p = $p.Next() }; return $null }
try {
  $doc = $app.Documents.Open($File); $from = 0; if ($doc.TablesOfContents.Count -gt 0) { $from = $doc.TablesOfContents.Item(1).Range.End }
  $intro = Find-Lvl1 $doc $from 'บทนำ'
  if (-not $intro) { "no intro chapter" } else {
    $next = Find-Lvl1 $doc $intro.Range.End ''
    $pn = $intro.Range.Sections.Item(1).Headers.Item(1).PageNumbers; $restart = $pn.RestartNumberingAtSection; $startNo = $pn.StartingNumber
    $s = $intro.Range.Start; $e = $(if ($next) { $next.Range.Start } else { $doc.Content.End - 1 })
    $doc.Range($s, $e).Delete() | Out-Null
    if ($restart) { $p2 = $doc.Range($s, $s).Sections.Item(1).Headers.Item(1).PageNumbers; $p2.RestartNumberingAtSection = $true; $p2.StartingNumber = $startNo }
    "intro chapter removed ($($e - $s) characters)"
  }
  $doc.Save(); $doc.Close()
} finally { $app.Quit() }
