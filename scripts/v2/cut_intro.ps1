param([string]$Book3 = '', [string]$Only = '?')
# Pass 1: ลบบทนำ (หัวข้อระดับ 1 ที่มีคำว่า บทนำ ถึงก่อนบทถัดไป) ทุกเล่ม เล็ก->ใหญ่; (5) ต่อท้ายด้วยบทสถาปัตยกรรมข้อมูลจาก (3)
$ErrorActionPreference = 'Stop'
$D = (Join-Path $env:USERPROFILE 'Downloads')
$files = Get-ChildItem -LiteralPath $D -Filter "($Only) *.docx" | Sort-Object Length
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
$app.Options.Pagination = $false
function Find-Lvl1($doc, $startPos, [string]$must) {
  $p = $doc.Range($startPos, $startPos).Paragraphs.Item(1)
  while ($p -ne $null) {
    if ($p.OutlineLevel -eq 1) { $t = $p.Range.Text.Trim(); if ($t -and (-not $must -or $t.Contains($must))) { return $p } }
    $p = $p.Next()
  }
  return $null
}
try {
  foreach ($f in $files) {
    $doc = $app.Documents.Open($f.FullName)
    # skip TOC: start after the last TOC range
    $from = 0; if ($doc.TablesOfContents.Count -gt 0) { $from = $doc.TablesOfContents.Item(1).Range.End }
    $intro = Find-Lvl1 $doc $from 'บทนำ'
    if (-not $intro) { "$($f.Name.Substring(0,3)) no intro found"; $doc.Close($false); continue }
    $next = Find-Lvl1 $doc $intro.Range.End ''
    $sec = $intro.Range.Sections.Item(1); $pn = $sec.Headers.Item(1).PageNumbers
    $restart = $pn.RestartNumberingAtSection; $startNo = $pn.StartingNumber
    $s = $intro.Range.Start; $e = $next.Range.Start
    $doc.Range($s, $e).Delete() | Out-Null
    $sec2 = $doc.Range($s, $s).Sections.Item(1); $pn2 = $sec2.Headers.Item(1).PageNumbers
    if ($restart) { $pn2.RestartNumberingAtSection = $true; $pn2.StartingNumber = $startNo }
    "$($f.Name.Substring(0,3)) intro removed ($($e - $s) chars), restart=$restart start=$startNo"
    if ($f.Name.StartsWith('(5)') -and $Book3) {
      $src = $app.Documents.Open($Book3, $false, $true)
      $h = Find-Lvl1 $src $src.TablesOfContents.Item(1).Range.End 'DATA ARCHITECTURE'
      if (-not $h) { $h = Find-Lvl1 $src $src.TablesOfContents.Item(1).Range.End 'สถาปัตยกรรมข้อมูล' }
      $sr = $src.Range($h.Range.Start, $src.Content.End)
      $end = $doc.Content.End - 1
      $doc.Range($end, $end).InsertParagraphAfter()
      $dr = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
      $dr.FormattedText = $sr.FormattedText
      $hp = $doc.Range($dr.Start, $dr.Start).Paragraphs.Item(1)
      if (-not $hp.Format.PageBreakBefore) { $hp.Format.PageBreakBefore = $true }
      "(5) appended chapter from (3): $($sr.End - $sr.Start) chars, heading style=$($hp.Style.NameLocal)"
      $src.Close($false)
    }
    $doc.Save(); $doc.Close()
  }
} finally { $app.Quit() }
