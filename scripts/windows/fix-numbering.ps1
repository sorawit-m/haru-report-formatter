# fix-numbering.ps1 — เทียบเลขหัวข้ออัตโนมัติ (เช่น "1.1", "บทที่ 1") ระหว่างต้นฉบับกับไฟล์ที่จัดแล้ว
# หัวข้อที่ต้นฉบับ "ไม่มีเลข" แต่ไฟล์ใหม่มีเลขงอกขึ้นมา (เช่น "ส่วน 2.03") -> เอาเลขออกเฉพาะย่อหน้านั้น
# หัวข้อที่ต้นฉบับ "มีเลข" ห้ามแตะ   *** บันทึกไฟล์นี้เป็น UTF-8 with BOM ***
# เร็ว: ไล่เฉพาะย่อหน้าที่มีเลขรายการ (ListParagraphs) ไม่ไล่ทุกย่อหน้าทั้งเล่ม
# (ย่อหน้าเนื้อหายาวที่หลุดเข้าสารบัญ จัดการใน restore-styles.mjs / fix-tables.mjs แล้ว)
param([string]$Orig, [string]$File)
function Log($m) { Write-Output ("[{0:HH:mm:ss}] {1}" -f (Get-Date), $m) }
$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0; $w.ScreenUpdating = $false
$opt = $w.Options; $oldSpell = $opt.CheckSpellingAsYouType; $oldGram = $opt.CheckGrammarAsYouType; $oldPag = $opt.Pagination
$opt.CheckSpellingAsYouType = $false; $opt.CheckGrammarAsYouType = $false; $opt.Pagination = $false
try {
  # key = text without any spaces / line breaks (a heading may gain or lose a Shift+Enter in formatting)
  function K($s) { ($s -replace '[\r\a\v\f\s​]', '') }
  function NumMap($d) {
    $m = @{}
    foreach ($p in $d.ListParagraphs) {
      if ($p.OutlineLevel -gt 9) { continue }
      $t = K $p.Range.Text; if (-not $t -or $t.Length -gt 150) { continue }
      if ($p.Range.Information(12)) { continue }
      if (-not $m.ContainsKey($t)) { $m[$t] = $p.Range.ListFormat.ListString }
    }
    return $m
  }
  $o = $w.Documents.Open($Orig, $false, $true); $om = NumMap $o; $o.Close($false)
  $d = $w.Documents.Open($File, $false, $false); $n = 0
  $cands = @(); foreach ($p in $d.ListParagraphs) { if ($p.OutlineLevel -le 9) { $cands += $p } }
  foreach ($p in $cands) {
    $t = ($p.Range.Text -replace '[\r\a\v\f]', '').Trim(); if (-not $t -or $t.Length -gt 150) { continue }
    $ls = $p.Range.ListFormat.ListString; if (-not $ls) { continue }
    if ($p.Range.Information(12)) { continue }
    # the heading already starts with its own typed number ("บทที่ 1 …", "2.3 …") -> the automatic number is a duplicate
    $typed = $t -match '^(บทที่\s*\d+|\d+(\.\d+)+\.?)\s'
    # in the original this exact heading was a numbered-list paragraph whose number was empty
    # (headings that were not list paragraphs at all are handled per paragraph in restore-styles.mjs — never guess from
    #  the document text: the table of contents repeats every heading and made real "บทที่ 3" numbers disappear)
    $k = K $t
    $hadNone = $om.ContainsKey($k) -and -not $om[$k]
    if ($typed -or $hadNone) { $p.Range.ListFormat.RemoveNumbers(); $n++ }
  }
  Log "stray heading numbers removed: $n (checked $($cands.Count) numbered heading paragraphs)"
  $d.Save(); $d.Close($false)
} finally {
  $opt.CheckSpellingAsYouType = $oldSpell; $opt.CheckGrammarAsYouType = $oldGram; $opt.Pagination = $oldPag; $w.ScreenUpdating = $true
  $w.Quit()
}
