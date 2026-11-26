# check-numbers.ps1 — ตรวจเลขอัตโนมัติ (1.1, 1.2.1, บทที่ …) ระหว่างต้นฉบับกับไฟล์ที่จัดแล้ว ทีละย่อหน้า
# รายงาน "เลขหาย" (ต้นฉบับมี ผลลัพธ์ไม่มี) และ "เลขงอก" (ต้นฉบับไม่มี ผลลัพธ์มี เช่น "บทความ I.")   *** UTF-8 with BOM ***
# ใช้: check-numbers.ps1 -Orig ต้นฉบับ.docx -File ผล.docx  -> พิมพ์ผล; ถ้ามีปัญหา บรรทัดแรกขึ้นต้นด้วย "NUMBERS-PROBLEM"
param([string]$Orig, [string]$File)
$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0; $w.ScreenUpdating = $false
$opt = $w.Options; $oldSpell = $opt.CheckSpellingAsYouType; $oldGram = $opt.CheckGrammarAsYouType
$opt.CheckSpellingAsYouType = $false; $opt.CheckGrammarAsYouType = $false
function Key($p) { $t = ($p.Range.Text -replace '[\r\a\v\f\s]', ''); if ($t.Length -gt 40) { $t.Substring(0, 40) } else { $t } }
function NumMap($d) { $m = @{}; foreach ($p in $d.ListParagraphs) { $k = Key $p; if ($k -and -not $m.ContainsKey($k)) { $m[$k] = $p.Range.ListFormat.ListString } }; return $m }
try {
  $o = $w.Documents.Open($Orig, $false, $true); $om = NumMap $o; $otext = ($o.Content.Text -replace '[\r\a\v\f\s]', ''); $o.Close($false)
  $d = $w.Documents.Open($File, $false, $true); $nm = NumMap $d; $ntext = ($d.Content.Text -replace '[\r\a\v\f\s]', ''); $d.Close($false)
  $lost = @(); $gained = @()
  # a stray number the original already had ("บทความ I.", "ส่วน 2.03", "(a)") being gone is a FIX, not a loss
  $junk = '^(บทความ|ส่วน|Article|Section)\s|^\(?[a-z]\)$|^\([ivx]+\)$'
  foreach ($k in $om.Keys) { if ($om[$k] -and -not $nm[$k] -and $ntext.Contains($k) -and $om[$k] -notmatch $junk) { $lost += "  หาย: '$($om[$k])' หน้า '$k'" } }
  foreach ($k in $nm.Keys) { if ($nm[$k] -and -not $om[$k] -and $otext.Contains($k)) { $gained += "  งอก: '$($nm[$k])' หน้า '$k'" } }
  if ($lost.Count -or $gained.Count) {
    "NUMBERS-PROBLEM เลขหาย=$($lost.Count) เลขงอก=$($gained.Count)"; $lost | Select-Object -First 15; $gained | Select-Object -First 15
  } else { "เลขอัตโนมัติตรงกับต้นฉบับทุกย่อหน้า (ตรวจ $($om.Count) ย่อหน้าที่มีเลข)" }
} finally {
  $opt.CheckSpellingAsYouType = $oldSpell; $opt.CheckGrammarAsYouType = $oldGram; $w.ScreenUpdating = $true; $w.Quit()
}
