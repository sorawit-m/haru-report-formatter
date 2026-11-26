# format-word.ps1 — จัดรูปแบบเล่มรายงานด้วย Microsoft Word (Windows)  *** บันทึกไฟล์นี้เป็น UTF-8 with BOM เสมอ ***
# ใช้: powershell -ExecutionPolicy Bypass -File format-word.ps1 -In ต้นฉบับ.docx -Out ผล.docx -Config config.json [-Replacements คำแก้.tsv] [-Pdf ผล.pdf]
param([string]$In, [string]$Out, [string]$Config, [string]$Replacements = "", [string]$Pdf = "")
$ErrorActionPreference = 'Stop'
trap { Write-Output ("ERROR: " + $_.Exception.Message + " @ " + $_.InvocationInfo.ScriptLineNumber); try { $d.Close($false) } catch {}; try { $w.Quit() } catch {}; break }
function Log($m) { Write-Output ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $m) }
$cfg = Get-Content -LiteralPath $Config -Raw -Encoding UTF8 | ConvertFrom-Json
$HFONT = if ($cfg.heading_font) { $cfg.heading_font } else { 'TH SarabunPSK' }
$CODEF = if ($cfg.code_fonts) { $cfg.code_fonts } else { 'Consol|Courier|Mono' }
$KEEPROWS = if ($cfg.keep_whole_table_max_rows) { [int]$cfg.keep_whole_table_max_rows } else { 8 }
$KEEPCHARS = if ($cfg.keep_whole_table_max_chars) { [int]$cfg.keep_whole_table_max_chars } else { 2000 }
$PADPT = if ($cfg.table_cell_padding_pt) { [double]$cfg.table_cell_padding_pt } else { 5.4 }
$LABEL = if ($cfg.table_caption_label) { $cfg.table_caption_label } else { 'ตารางที่' }

# ---------- 1. XML-level: fonts, faint table borders, leading spaces in cells ----------
Copy-Item -LiteralPath $In -Destination $Out -Force
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::Open($Out, 'Update')
$utf8 = New-Object System.Text.UTF8Encoding($false)
$bsz = [string][Math]::Round(([double]$(if ($cfg.table_border_pt) { $cfg.table_border_pt } else { 0.5 })) * 8)
foreach ($e in @($zip.Entries)) {
  if ($e.FullName -match '^word/.*\.xml$') {
    $sr = New-Object IO.StreamReader($e.Open(), $utf8); $x = $sr.ReadToEnd(); $sr.Close()
    $y = $x
    if ($cfg.font_replace) { foreach ($pr in $cfg.font_replace.PSObject.Properties) { $y = $y.Replace('"' + $pr.Name + '"', '"' + $pr.Value + '"') } }
    if ($cfg.thin_border_to_black) {
      # faint borders (thin 0.125/0.25 pt, or a light colour) look "missing" -> black, except colours listed in keep_border_colors (e.g. code boxes)
      $keepC = @(); if ($cfg.keep_border_colors) { $keepC = @($cfg.keep_border_colors | % { $_.ToUpper() }) }
      $y = [regex]::Replace($y, '<w:(top|left|bottom|right|insideH|insideV|start|end) w:val="single" w:sz="(\d+)" w:space="(\d+)" w:color="([0-9A-Fa-f]{6})"', {
        param($m)
        $c = $m.Groups[4].Value.ToUpper(); $sz = [int]$m.Groups[2].Value
        $r = [Convert]::ToInt32($c.Substring(0,2),16); $g = [Convert]::ToInt32($c.Substring(2,2),16); $bl = [Convert]::ToInt32($c.Substring(4,2),16)
        $light = (0.299*$r + 0.587*$g + 0.114*$bl) -gt 140
        if ($keepC -contains $c -or (-not $light -and $sz -ge 4)) { return $m.Value }
        return ('<w:' + $m.Groups[1].Value + ' w:val="single" w:sz="' + [Math]::Max($sz, [int]$bsz) + '" w:space="' + $m.Groups[3].Value + '" w:color="000000"')
      })
    }
    if ($cfg.no_italic) { $y = [regex]::Replace($y, '<w:i(?:Cs)?(?: w:val="(?:1|true|on)")?/>', '') }
    if ($y -ne $x) {
      $name = $e.FullName; $e.Delete()
      $ne = $zip.CreateEntry($name); $sw = New-Object IO.StreamWriter($ne.Open(), $utf8); $sw.Write($y); $sw.Close()
    }
  }
}
$zip.Dispose()
Log "fonts unified in XML"

$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0; $w.ScreenUpdating = $false
# faster: no live spell/grammar check and no background pagination while editing (restored at the end — they are the user's Word settings)
$opt = $w.Options; $oldSpell = $opt.CheckSpellingAsYouType; $oldGram = $opt.CheckGrammarAsYouType; $oldPag = $opt.Pagination
$opt.CheckSpellingAsYouType = $false; $opt.CheckGrammarAsYouType = $false; $opt.Pagination = $false
$d = $w.Documents.Open($Out, $false, $false)
$d.TrackRevisions = $false
try { $d.Compatibility(14) = $true } catch {}   # wdDoNotExpandShiftReturn

# heading styles -> TH SarabunPSK + keep with next
foreach ($s in $d.Styles) {
  try {
    if ($s.Type -eq 1 -and $s.ParagraphFormat.OutlineLevel -lt 10) {
      # change a built-in heading style ONLY when needed: touching it makes Word re-attach a list template that is
      # linked to it in numbering.xml ("บทความ III.", "ส่วน 2.03" appearing in front of headings)
      if ($s.ParagraphFormat.KeepWithNext -ne -1) { $s.ParagraphFormat.KeepWithNext = -1 }
      if ($s.ParagraphFormat.KeepTogether -ne -1) { $s.ParagraphFormat.KeepTogether = -1 }
    }
  } catch {}
}
foreach ($id in -2, -3, -4, -5) { try { $st = $d.Styles.Item([int]$id); if ($st.Font.Name -ne $HFONT) { $st.Font.Name = $HFONT }; if ($st.Font.NameBi -ne $HFONT) { $st.Font.NameBi = $HFONT } } catch {} }

# ---------- 2. typo replacements ----------
if ($Replacements -and (Test-Path $Replacements)) {
  $n = 0
  foreach ($line in [IO.File]::ReadAllLines($Replacements, [Text.Encoding]::UTF8)) {
    if (-not $line.Trim() -or $line.StartsWith('#')) { continue }
    $p = $line.Split("`t"); if ($p.Count -lt 2) { continue }
    $r = $d.Content; $f = $r.Find; $f.ClearFormatting(); $f.Replacement.ClearFormatting()
    $ok = $f.Execute($p[0], $true, $false, $false, $false, $false, $true, 0, $false, $p[1], 2)
    if ($ok) { $n++ } else { Log "  not found: $($p[0])" }
  }
  Log "typo replacements applied: $n"
}

# ---------- 2b. remove paragraphs pasted twice ----------
if ($cfg.remove_duplicate_paragraphs) {
  $dupDelP = 0
  # find long lines (>80 chars, not table cells) that repeat within 8 lines of each other
  $lines = $d.Content.Text -split "`r"; $prefs = @()
  for ($i = 0; $i -lt $lines.Count; $i++) {
    $t = $lines[$i]; if ($t.Length -lt 80 -or $t -match "[\a]") { continue }
    for ($j = [Math]::Max(0, $i - 8); $j -lt $i; $j++) { if ($lines[$j] -eq $t) { $prefs += $t.Substring(0, [Math]::Min(120, $t.Length)); break } }
  }
  foreach ($pref in ($prefs | Select-Object -Unique)) {
    $r = $d.Content; $f = $r.Find; $f.ClearFormatting(); $f.Text = $pref; $f.Wrap = 0; $f.MatchWildcards = $false
    if (-not $f.Execute()) { continue }
    $p1 = $r.Paragraphs.Item(1); $t1 = $p1.Range.Text
    $r2 = $d.Range($p1.Range.End, $d.Content.End); $f2 = $r2.Find; $f2.Text = $pref; $f2.Wrap = 0
    if ($f2.Execute()) {
      $p2 = $r2.Paragraphs.Item(1)
      if ($p2.Range.Text -eq $t1 -and ($p2.Range.Start - $p1.Range.End) -lt 3000) { $p2.Range.Delete() | Out-Null; $dupDelP++ }
    }
  }
  Log "duplicated paragraphs removed: $dupDelP"
}

# ---------- 3. stray manual line breaks inside justified Thai text ----------
$rng = $d.Content; $f = $rng.Find; $f.ClearFormatting(); $f.Text = "^l"; $f.Forward = $true; $f.Wrap = 0
$removed = 0; $guard = 0
while ($f.Execute() -and $guard -lt 3000) {
  $guard++
  $al = $rng.ParagraphFormat.Alignment
  $next = $d.Range($rng.End, [Math]::Min($rng.End + 1, $d.Content.End)).Text
  $font = $rng.Font.Name
  if ($cfg.remove_stray_line_breaks -and ($al -eq 9 -or $al -eq 3) -and $next -notmatch '^\s' -and $font -notmatch $CODEF) {
    $rng.Text = ""; $removed++
  }
  $rng.Collapse(0)
}
Log "manual line breaks removed: $removed"

# ---------- 4. collapse runs of empty paragraphs (not on cover section) ----------
if ($d.Sections.Count -ge 2) {
  for ($pass = 0; $pass -lt 6; $pass++) {
    $r = $d.Range($d.Sections.Item(2).Range.Start, $d.Content.End)
    $f = $r.Find; $f.ClearFormatting(); $f.Replacement.ClearFormatting()
    $ok = $f.Execute("^p^p^p", $false, $false, $false, $false, $false, $true, 0, $false, "^p^p", 2)
    if (-not $ok) { break }
  }
}
Log "empty paragraph runs collapsed"
# cover: picture paragraphs formatted as a bullet list show a stray dot -> plain centred paragraph
$cv = 0; foreach ($p in $d.Sections.Item(1).Range.Paragraphs) { try { if ($p.Range.ListFormat.ListType -ne 0 -and $p.Range.InlineShapes.Count -gt 0) { $p.Range.ListFormat.RemoveNumbers(); $p.Style = $d.Styles.Item(-1); $p.Alignment = 1; $p.LeftIndent = 0; $p.FirstLineIndent = 0; $cv++ } } catch {} }
Log "cover bullet dots removed: $cv"
$el = 0
foreach ($p in @($d.ListParagraphs)) { try { if (-not $p.Range.Information(12) -and ($p.Range.Text -replace '[\r\s]', '') -eq '' -and $p.Range.InlineShapes.Count -eq 0) { $p.Range.Delete() | Out-Null; $el++ } } catch {} }
Log "empty bullet items removed: $el"

# ---------- 5. table captions -> "ตารางที่ {SEQ Table}" ----------
$capCount = 0
# decide: real numbering fields (SEQ) only when the book has a typed "list of tables" that must be rebuilt
#         and the captions are numbered 1..N in order; otherwise just relabel and keep the numbers
$hasTabList = $false; $r0 = $d.Content; $f0 = $r0.Find; $f0.Text = 'สารบัญตาราง'; $f0.Wrap = 0; if ($f0.Execute()) { $hasTabList = ($d.TablesOfFigures.Count -eq 0) }
$nums = @(); foreach ($m in [regex]::Matches(($d.Content.Text), "(?m)(?<=\r)ตาราง\s*(\d+)\s")) { $nums += [int]$m.Groups[1].Value }
$seqOk = $hasTabList -and $nums.Count -gt 0 -and (($nums -join ',') -eq ((1..$nums.Count) -join ','))
$Kind = if ($seqOk) { 'seq' } else { 'relabel' }
if ($true) {
  $startPos = if ($d.Sections.Count -ge 3) { $d.Sections.Item(3).Range.Start } else { 0 }
  $r = $d.Range($startPos, $d.Content.End); $f = $r.Find; $f.ClearFormatting()
  $f.Text = "ตาราง[ ]@[0-9]@[ ]"; $f.MatchWildcards = $true; $f.Forward = $true; $f.Wrap = 0
  $g = 0
  while ($f.Execute() -and $g -lt 500) {
    $g++
    $para = $r.Paragraphs.Item(1)
    if ($para.Range.Start -eq $r.Start -and -not $r.Information(12)) {
      $num = ($r.Text -replace '[^0-9]', '')
      $r.MoveEnd(1, -1) | Out-Null   # keep the original space after the number
      if ($Kind -eq 'relabel') {
        $r.Text = "$LABEL $num"; $r.Collapse(0)
      } else {
        $r.Text = "$LABEL "
        $r.Collapse(0)
        $fld = $d.Fields.Add($r, -1, "SEQ Table \* ARABIC", $false)
        $r = $d.Range($fld.Result.End, $fld.Result.End)
      }
      $capCount++
    } else { $r.Collapse(0) }
    $r.SetRange($r.End, $d.Content.End)
    $f = $r.Find; $f.Text = "ตาราง[ ]@[0-9]@[ ]"; $f.MatchWildcards = $true; $f.Forward = $true; $f.Wrap = 0
  }
}
Log "table captions converted ($Kind): $capCount"

# ---------- 5b. (disabled) headings inside table cells were a text-dump illusion; left here off on purpose ----------
$pat = ''
$moved = 0
if ($pat) {
  $refStyle = $null
  $g = 0; $pos = 0
  while ($g -lt 3000) {
    $g++
    $r = $d.Range($pos, $d.Content.End); $f = $r.Find; $f.ClearFormatting(); $f.Text = $pat; $f.MatchWildcards = $true; $f.Wrap = 0
    if (-not $f.Execute()) { break }
    $pos = $r.End
    if (-not $r.Information(12)) { if (-not $refStyle) { $p0 = $r.Paragraphs.Item(1); if ($p0.Range.Start -eq $r.Start -and $p0.OutlineLevel -lt 10) { $refStyle = $p0.Style } }; continue }
    $cell = $r.Cells.Item(1); $hp = $r.Paragraphs.Item(1)
    if ($hp.Range.Start -ne $cell.Range.Start -or $r.Start -ne $hp.Range.Start) { continue }
    $htxt = ($hp.Range.Text -replace '[\r\a]', '').Trim()
    if ($Kind -eq 'arch' -and $htxt -notmatch '^4\.3(\.\d+)+ \S') { continue }
    if ($hp.Range.End -ge $cell.Range.End) { continue }   # heading is the whole cell -> leave
    try {
      $tb = $r.Tables.Item(1); $ri = $cell.RowIndex
      $new = $tb.Split($ri)
      $sp = $d.Range($new.Range.Start - 1, $new.Range.Start - 1).Paragraphs.Item(1)
      $hp2 = $new.Range.Cells.Item(1).Range.Paragraphs.Item(1)
      $sp.Range.FormattedText = $hp2.Range.FormattedText
      $hp2.Range.Delete() | Out-Null
      $sp2 = $d.Range($new.Range.Start - 1, $new.Range.Start - 1).Paragraphs.Item(1)
      if ($refStyle -and $Kind -eq 'api') { $sp2.Style = $refStyle }
      $sp2.KeepWithNext = -1
      $pos = $new.Range.Start
      $moved++
    } catch { Log "  heading move failed at '$htxt': $($_.Exception.Message)" }
  }
}
Log "headings moved out of tables: $moved"

# ---------- 6. tables: header repeat, no row split, keep small tables together ----------
$shadeAuto = -16777216
function IsShaded($cell) { try { $c = $cell.Shading.BackgroundPatternColor; return ($c -ne $shadeAuto -and $c -ne 16777215) } catch { return $false } }
$tc = $d.Tables.Count; $hdrSet = 0; $dupDel = 0; $kept = 0; $smallList = @()
for ($t = 1; $t -le $tc; $t++) {
  $tb = $d.Tables.Item($t)
  $nrows = $tb.Rows.Count; $ncols = $tb.Columns.Count
  $cells = $tb.Range.Cells
  # header detection
  $hdr = 0
  try {
    $first = $cells.Item(1)
    $txt1 = ($first.Range.Text -replace '[\r\a]', '').Trim()
    $r1cells = @(); $r2cells = @()
    for ($i = 1; $i -le [Math]::Min($cells.Count, 3 * $ncols + 6); $i++) { $c = $cells.Item($i); if ($c.RowIndex -eq 1) { $r1cells += $c } elseif ($c.RowIndex -eq 2) { $r2cells += $c } }
    $lastR1 = $r1cells[-1]
    if ($nrows -ge 2 -and $r1cells.Count -ge 2) {
      $r1sh = IsShaded $lastR1
      $r2sh = if ($r2cells.Count) { IsShaded $r2cells[-1] } else { $false }
      $r1bold = ($lastR1.Range.Font.Bold -eq -1) -and ($first.Range.Font.Bold -eq -1)
      if (($r1sh -and -not $r2sh) -or ($r1bold -and $r2cells.Count -and $r2cells[-1].Range.Font.Bold -ne -1) -or ($r1sh -and $r1bold)) { $hdr = 1 }
    }
    if ($txt1 -match '^(Table Name|ชื่อตาราง)') {
      # multi-row header: extend through the column-title row
      $hdr = 1
      for ($i = 1; $i -le [Math]::Min($cells.Count, 6 * $ncols); $i++) {
        $c = $cells.Item($i); if ($c.ColumnIndex -ne 1 -or $c.RowIndex -gt 4) { continue }
        $tx = ($c.Range.Text -replace '[\r\a]', '').Trim()
        if ($tx -match '^(NO|No\.?|ลำดับ|#)$') { $hdr = $c.RowIndex }
        elseif ($tx -match '^Table ') { $hdr = [Math]::Max($hdr, $c.RowIndex) }
      }
    }
  } catch { $hdr = 0 }

  if ($hdr -gt 0) {
    try {
      $lastHdrCell = $null
      for ($i = 1; $i -le $cells.Count; $i++) { $c = $cells.Item($i); if ($c.RowIndex -le $hdr) { $lastHdrCell = $c } else { break } }
      $hr = $d.Range($tb.Range.Start, $lastHdrCell.Range.End)
      $hr.Rows.HeadingFormat = -1; $hdrSet++
      # delete manually re-typed header rows further down (left over from old pagination)
      $sigCells = @(); for ($i = 1; $i -le $cells.Count; $i++) { $c = $cells.Item($i); if ($c.RowIndex -eq $hdr) { $sigCells += $c } elseif ($c.RowIndex -gt $hdr) { break } }
      if ($sigCells.Count -ge 2) {
        $sig = (($sigCells | % { ($_.Range.Text -replace '[\r\a\s]', '') }) -join '|')
        for ($ri = $nrows; $ri -gt $hdr; $ri--) {
          try {
            $row = $tb.Rows.Item($ri)
            $rs = (($row.Cells | % { ($_.Range.Text -replace '[\r\a\s]', '') }) -join '|')
            if ($rs -eq $sig) { $row.Delete(); $dupDel++ }
          } catch { break }
        }
      }
    } catch { Log "  table $t header set failed: $($_.Exception.Message)" }
  }
  # rows should not split across pages (unless huge)
  $len = $tb.Range.End - $tb.Range.Start
  try {
    if ($len -lt 1800 * [Math]::Max(1, $nrows)) { $tb.Range.Rows.AllowBreakAcrossPages = 0 }
  } catch {}
  # keep caption/title paragraph with table
  if ($tb.Range.Start -gt 1) {
    $prev = $d.Range($tb.Range.Start - 1, $tb.Range.Start - 1).Paragraphs.Item(1)
    if (-not $prev.Range.Information(12)) {
      $pt = ($prev.Range.Text -replace '[\r\s]', '')
      if ($pt.Length -gt 0 -and $pt.Length -lt 250) { $prev.KeepWithNext = -1 }
    }
  }
  # small tables: keep whole table on one page; big tables: keep header + first 2 rows together
  # only tiny tables are kept whole; bigger ones keep header + 2 rows (so editing in Word does not make big blocks jump pages)
  $small = ($nrows -le $KEEPROWS -and $len -lt $KEEPCHARS)
  # rows that hold a picture (e.g. a screenshot row) are never chained to the rows above: the chain would not fit on one page
  $picRows = @{}; foreach ($s in $tb.Range.InlineShapes) { try { $picRows[$s.Range.Cells.Item(1).RowIndex] = 1 } catch {} }
  if ($picRows.Count -gt 0) { $small = $false }
  try {
    if ($small) {
      $tb.Range.ParagraphFormat.KeepWithNext = -1
      for ($i = $cells.Count; $i -ge 1; $i--) { $c = $cells.Item($i); if ($c.RowIndex -eq $nrows) { $c.Range.ParagraphFormat.KeepWithNext = 0 } else { break } }
      $kept++; $smallList += $t
    } else {
      $lim = [Math]::Max($hdr, 1) + 1
      for ($i = 1; $i -le $cells.Count; $i++) { $c = $cells.Item($i); if ($c.RowIndex -le $lim) { if (-not $picRows.ContainsKey($c.RowIndex + 1)) { $c.Range.ParagraphFormat.KeepWithNext = -1 } } else { break } }
    }
  } catch {}
  if ($t % 200 -eq 0) { Log "  tables processed $t/$tc" }
}
Log "tables=$tc headerRepeat=$hdrSet dupHeaderRowsDeleted=$dupDel keptTogether=$kept"
# 6b. a "small" table that is still taller than one page must be allowed to break (header repeats instead)
$d.Repaginate(); $unkept = 0
foreach ($ti in $smallList) {
  try {
    $tb = $d.Tables.Item($ti)
    if ($tb.Range.Characters.Last.Information(3) -gt $tb.Range.Characters.First.Information(3)) {
      $tb.Range.ParagraphFormat.KeepWithNext = 0
      $cells = $tb.Range.Cells; $hr = 1
      try { for ($i = 1; $i -le $cells.Count; $i++) { $c = $cells.Item($i); if ($c.Range.Rows.HeadingFormat -eq -1) { $hr = [Math]::Max($hr, $c.RowIndex) } else { break } } } catch {}
      for ($i = 1; $i -le $cells.Count; $i++) { $c = $cells.Item($i); if ($c.RowIndex -le $hr + 1) { $c.Range.ParagraphFormat.KeepWithNext = -1 } else { break } }
      if ($hr -eq 1) { try { $d.Range($tb.Range.Start, $cells.Item(1).Range.End).Rows.HeadingFormat = -1 } catch {} }
      $unkept++
    }
  } catch {}
}
Log "tables too tall to keep together -> allowed to break: $unkept"

# 6c. arch: tables without borders get the same grid as the rest of the book
if ($d.Tables.Count -gt 400) { Log "borderless-table pass skipped (large book: faint/missing borders were already fixed in XML)" }
else {
  $bd = 0
  foreach ($tb in $d.Tables) {
    try {
      if ($tb.Columns.Count -ge 2 -and ($tb.Borders.Item(-2).LineStyle -ne 1 -or $tb.Borders.Item(-7).LineStyle -ne 1)) {
        foreach ($bi in -1, -2, -3, -4, -5, -6) { $tb.Borders.Item($bi).LineStyle = 1; $tb.Borders.Item($bi).LineWidth = 4; $tb.Borders.Item($bi).Color = 0 }
        # cell-level borders override the table ones (e.g. missing top line on the header row)
        $cb = $tb.Range.Cells.Borders; foreach ($bi in -1, -2, -3, -4) { $cb.Item($bi).LineStyle = 1; $cb.Item($bi).LineWidth = 4; $cb.Item($bi).Color = 0 }
        $bd++
      }
    } catch {}
  }
  Log "borderless tables given grid borders: $bd"
}

# 6d. every table: cell padding, no leading spaces in cells, all rows the same total width
$padFix = 0; $lead = 0; $wFix = 0
foreach ($tb in $d.Tables) {
  try {
    $ncols = $tb.Columns.Count
    if ($tb.LeftPadding -lt 4) { $tb.LeftPadding = $PADPT; $tb.RightPadding = $PADPT; $padFix++ }
    if (-not $cfg.equalize_row_width -or $d.Tables.Count -gt 400) { continue }   # very large books: per-cell pass is too slow
    $rowSum = @{}; $cl = @()
    foreach ($c in $tb.Range.Cells) {
      $cl += $c
      $ri = $c.RowIndex; $rowSum[$ri] = [double]$rowSum[$ri] + $c.Width
      if ($false) {
        $first = $c.Range.Characters.First
        $g = 0
        while ($g -lt 12 -and ($first.Text -eq ' ' -or $first.Text -eq [string][char]160) -and $c.Range.Font.Name -notmatch 'Consol|Courier|Mono') { $first.Delete() | Out-Null; $first = $c.Range.Characters.First; $g++ }
        if ($g -gt 0) { $lead++ }
      }
    }
    $vmerge = $false; try { $null = $tb.Rows.Item(1).Cells.Count } catch { $vmerge = $true }   # vertically merged cells -> row sums are misleading
    if ($vmerge) {
      # only pull the whole table back inside the page if it sticks out
      $ps = $tb.Range.Sections.Item(1).PageSetup; $textW = $ps.PageWidth - $ps.LeftMargin - $ps.RightMargin - $ps.Gutter
      $max = ($rowSum.Values | Measure-Object -Maximum).Maximum
      if ($max -gt $textW + 2) { foreach ($c in $cl) { $c.Width = $c.Width * $textW / $max }; $wFix++ }
    }
    elseif ($rowSum.Count -ge 2) {
      $max = ($rowSum.Values | Measure-Object -Maximum).Maximum; $min = ($rowSum.Values | Measure-Object -Minimum).Minimum
      $ps = $tb.Range.Sections.Item(1).PageSetup
      $textW = $ps.PageWidth - $ps.LeftMargin - $ps.RightMargin - $ps.Gutter
      $target = [Math]::Min($max, $textW)
      if ($max - $min -gt 3 -or $max -gt $textW + 2) {
        foreach ($c in $cl) { $s = $rowSum[$c.RowIndex]; if ($s -gt 0 -and [Math]::Abs($target - $s) -gt 1) { $c.Width = $c.Width * $target / $s } }
        $wFix++
      }
    }
  } catch {}
}
Log "tables: padding fixed=$padFix cells with leading spaces=$lead rows width-equalised tables=$wFix"

# ---------- 7. figures stay with their caption ----------
foreach ($s in $d.InlineShapes) { try { if (-not $s.Range.Information(12)) { $s.Range.Paragraphs.Item(1).KeepWithNext = -1 } } catch {} }

# ---------- 7b. figure caption sharing a paragraph with its picture -> own paragraph ----------
$capStyle = $null; $split = 0
$seqF = @(); foreach ($fl in $d.Fields) { if ($fl.Type -eq 12 -and $fl.Code.Text -match 'Figure') { $seqF += $fl } }
$refCap = $null
foreach ($fl in $seqF) { $pp = $fl.Code.Paragraphs.Item(1); if ($pp.Range.InlineShapes.Count -eq 0) { $capStyle = $pp.Style; $refCap = $pp; break } }
foreach ($fl in $seqF) {
  $pp = $fl.Code.Paragraphs.Item(1)
  if ($pp.Range.InlineShapes.Count -gt 0) {
    $r = $d.Range($pp.Range.Start, $fl.Code.Start); $f = $r.Find; $f.ClearFormatting(); $f.Text = 'รูปที่'; $f.Wrap = 0; $f.MatchWildcards = $false
    if ($f.Execute()) {
      $r.InsertParagraphBefore()
      $cp = $fl.Code.Paragraphs.Item(1)
      if ($capStyle) { $cp.Style = $capStyle }
      if ($refCap) {
        $cp.Format = $refCap.Format
        $rf = $refCap.Range.Font
        $cp.Range.Font.Italic = $rf.Italic; $cp.Range.Font.ItalicBi = $rf.ItalicBi
        $cp.Range.Font.Bold = $rf.Bold; $cp.Range.Font.BoldBi = $rf.BoldBi
        $cp.Range.Font.Size = $rf.Size; $cp.Range.Font.SizeBi = $rf.SizeBi; $cp.Range.Font.Color = $rf.Color
      }
      $cp.Range.ParagraphFormat.Alignment = 1
      $d.Range($cp.Range.Start - 1, $cp.Range.Start - 1).Paragraphs.Item(1).KeepWithNext = -1
      $split++
    }
  }
}
Log "figure captions split from pictures: $split"

# ---------- 7c. heading + short description stay on the same page as the figure that follows ----------
$chained = 0
foreach ($s in $d.InlineShapes) {
  try {
    if ($s.Range.Information(12)) { continue }
    $pic = $s.Range.Paragraphs.Item(1)
    $q = $pic; $acc = 0
    for ($k = 0; $k -lt 10; $k++) {
      if ($q.Range.Start -lt 2) { break }
      $q = $d.Range($q.Range.Start - 1, $q.Range.Start - 1).Paragraphs.Item(1)
      if ($q.Range.Information(12) -or $q.Range.InlineShapes.Count -gt 0) { break }
      if ($q.Range.Fields.Count -gt 0 -and $q.Range.Text -match '^(รูปที่|ตารางที่)') { break }   # previous caption
      $acc += $q.Range.Text.Length
      if ($acc -gt 900) { break }        # long text may flow; only chain short intros
      $q.KeepWithNext = -1; $q.KeepTogether = -1
      if ($q.OutlineLevel -lt 10) { $chained++; break }    # reached the section heading
    }
  } catch {}
}
Log "figures chained to their heading/intro: $chained"

# ---------- 8. chapter headings start on a new page; front-matter headings too ----------
$heads = @()
$d.Bookmarks.ShowHidden = $true
$seenStart = @{}
foreach ($bm in $d.Bookmarks) {
  if ($bm.Name -notlike '_Toc*') { continue }
  $p = $bm.Range.Paragraphs.Item(1)
  if ($p.Range.Information(12)) { continue }   # a table cell styled as a heading is not a chapter
  if ($p.OutlineLevel -eq 1 -and -not $seenStart.ContainsKey($p.Range.Start)) { $seenStart[$p.Range.Start] = 1; $heads += $p }
}
$mainStart = if ($d.Sections.Count -ge 3) { $d.Sections.Item(3).Range.Start } else { 0 }
foreach ($word in 'สารบัญ', 'สารบัญรูปภาพ', 'สารบัญตาราง') {
  $r = $d.Range(0, $mainStart); $f = $r.Find; $f.ClearFormatting(); $f.Text = $word; $f.Wrap = 0
  $guard2 = 0
  while ($f.Execute() -and $guard2 -lt 50) {
    $guard2++
    if ($r.Start -ge $mainStart) { break }
    $p = $r.Paragraphs.Item(1)
    if ((($p.Range.Text -replace '[\r\v\f]', '').Trim()) -eq $word -and -not $seenStart.ContainsKey($p.Range.Start)) { $seenStart[$p.Range.Start] = 1; $heads += $p }
    if ($r.End -ge $mainStart) { break }; $r.SetRange($r.End, $mainStart); $f = $r.Find; $f.Text = $word; $f.Wrap = 0
  }
}
$seenToc = @{}
foreach ($p in $heads) {
  $tx = ($p.Range.Text -replace '[\r\v\f]', '').Trim()
  # remove empty / page-break-only paragraphs right before it
  for ($k = 0; $k -lt 6; $k++) {
    if ($p.Range.Start -lt 2) { break }
    $pp = $d.Range($p.Range.Start - 1, $p.Range.Start - 1).Paragraphs.Item(1)
    $pt = $pp.Range.Text
    if ($pt -match '^[\s\f]*\r$' -and $pt -notmatch '\x0C\r$' -or $pt -match '^\f\r$') {
      if ($pp.Range.Sections.Item(1).Index -ne $p.Range.Sections.Item(1).Index) { break }
      if ($pp.Range.InlineShapes.Count -gt 0) { break }
      try { $pp.Range.Delete() | Out-Null } catch { break }
    } else { break }
  }
  try { $fc = $d.Range($p.Range.Start, $p.Range.Start + 1); if ($fc.Text -eq [string][char]12) { $fc.Delete() | Out-Null } } catch {}
  try { $p.PageBreakBefore = -1 } catch {}
}
Log "chapter/front headings set to new page: $($heads.Count)"

# ---------- 9. report: rebuild static figure/table lists as real fields ----------
if ($d.TablesOfFigures.Count -eq 0 -and $Kind -eq 'seq') {
  function FindPara($text, $from) {
    $r = $d.Range($from, $d.Content.End); $f = $r.Find; $f.ClearFormatting(); $f.MatchWildcards = $false
    if ($f.Execute($text)) { return $r.Paragraphs.Item(1) } else { return $null }
  }
  $figH = FindPara "สารบัญรูปภาพ" 0
  $tabH = FindPara "สารบัญตาราง" 0
  if ($figH -and $tabH) {
    $figHdr = $figH.Next(1)           # "เรื่อง ... หน้า" line
    $tabHdr = $tabH.Next(1)
    # end of table list = paragraph before first chapter heading (or section end)
    $sec = $tabH.Range.Sections.Item(1)
    $endPos = $sec.Range.End - 1
    $chap = FindPara "บทนำ" $tabHdr.Range.End
    if ($chap -and $chap.Range.Start -lt $endPos) { $endPos = $chap.Range.Start }
    # table list
    $r = $d.Range($tabHdr.Range.End, $endPos); $r.Delete() | Out-Null
    $r = $d.Range($tabHdr.Range.End, $tabHdr.Range.End); $r.InsertParagraphBefore()
    $r = $d.Range($tabHdr.Range.End, $tabHdr.Range.End)
    $tof2 = $d.TablesOfFigures.Add($r, "Table"); $tof2.IncludeLabel = $true
    # figure list
    $r = $d.Range($figHdr.Range.End, $tabH.Range.Start); $r.Delete() | Out-Null
    $r = $d.Range($figHdr.Range.End, $figHdr.Range.End); $r.InsertParagraphBefore()
    $r = $d.Range($figHdr.Range.End, $figHdr.Range.End)
    $tof1 = $d.TablesOfFigures.Add($r, "Figure"); $tof1.IncludeLabel = $true
    $tabH.PageBreakBefore = -1; $figH.PageBreakBefore = -1
    # the empty paragraph left after the figure list can spill onto its own page -> neutralise it
    $pp = $d.Range($tabH.Range.Start - 1, $tabH.Range.Start - 1).Paragraphs.Item(1)
    if ($pp.Range.Text -match "^[\s\f\x15]*\r$") {
      try { $d.Range($pp.Range.End - 1, $pp.Range.End).Delete() | Out-Null }
      catch {
        try {
          $pp.Range.Font.Size = 1; $pp.Range.Font.SizeBi = 1
          $pp.SpaceBefore = 0; $pp.SpaceAfter = 0; $pp.LineSpacingRule = 4; $pp.LineSpacing = 1
          $pp.PageBreakBefore = -1; $tabH.PageBreakBefore = 0
        } catch {}
      }
    }
    Log "report figure/table lists rebuilt as fields"
  } else { Log "report list headings not found" }
}

# ---------- 10. arch: page numbers in body footer ----------
if ($cfg.add_page_numbers_if_missing -and $d.Sections.Count -ge 2) {
  $last = $d.Sections.Item($d.Sections.Count)
  $ft = $last.Footers.Item(1)
  if ($ft.Range.Fields.Count -eq 0) {
    $ft.LinkToPrevious = $false
    $fr = $ft.Range; $fr.Text = ""
    $d.Fields.Add($fr, 33) | Out-Null
    $ft.Range.ParagraphFormat.Alignment = $(switch ($cfg.page_number_align) { 'center' { 1 } 'left' { 0 } default { 2 } })
    $sz = if ($cfg.page_number_font_size) { $cfg.page_number_font_size } else { 16 }
    $ft.Range.Font.Name = $HFONT; $ft.Range.Font.NameBi = $HFONT; $ft.Range.Font.Size = $sz; $ft.Range.Font.SizeBi = $sz
    $ft.PageNumbers.RestartNumberingAtSection = $true; $ft.PageNumbers.StartingNumber = 1
    Log "page numbers added to the last section"
  }
}

# ---------- 10b. no italics anywhere ----------
if ($cfg.no_italic) {
  foreach ($s in $d.Styles) { try { if ($s.Font.Italic -ne 0) { $s.Font.Italic = 0; $s.Font.ItalicBi = 0 } } catch {} }
  $d.Content.Font.Italic = 0; $d.Content.Font.ItalicBi = 0
  foreach ($sec in $d.Sections) { foreach ($k in 1, 2, 3) { try { $sec.Headers.Item($k).Range.Font.Italic = 0; $sec.Footers.Item($k).Range.Font.Italic = 0 } catch {} } }
  Log "italics removed"
}

# ---------- 11. update fields, TOCs ----------
$d.Fields.Update() | Out-Null
# TOCs: when run.ps1 drives the pipeline, remove-blank-pages.ps1 rebuilds them at the very end -> skip here (saves minutes on big books)
if ($Pdf) {
  $d.Repaginate()
  foreach ($toc in $d.TablesOfContents) { $toc.Update() }
  foreach ($tof in $d.TablesOfFigures) { $tof.Update() }
  $d.Repaginate()
  foreach ($toc in $d.TablesOfContents) { $toc.UpdatePageNumbers() }
  foreach ($tof in $d.TablesOfFigures) { $tof.UpdatePageNumbers() }
}
Log ("pages after fix: " + $d.ComputeStatistics(2))
$d.Save()
if ($Pdf) { $d.ExportAsFixedFormat($Pdf, 17) }
$d.Close($false)
$opt.CheckSpellingAsYouType = $oldSpell; $opt.CheckGrammarAsYouType = $oldGram; $opt.Pagination = $oldPag; $w.ScreenUpdating = $true
$w.Quit()
Log "saved $Out"
