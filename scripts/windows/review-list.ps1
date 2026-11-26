# review-list.ps1 — สร้างรายการ "จุดที่ควรเข้าไปเช็ค" พร้อมเลขหน้า (เลขหน้าที่พิมพ์ในเล่ม)   *** UTF-8 with BOM ***
# รับรายการข้อความ (ไฟล์ TSV: ประเภท<TAB>ข้อความ<TAB>เหตุผล) แล้วหาเลขหน้าใน Word
# ใช้: review-list.ps1 -File ผล.docx -Items รายการ.tsv -Out ผล_ควรเช็ค.txt
param([string]$File, [string]$Items, [string]$Out)
$rows = @(); foreach ($l in [IO.File]::ReadAllLines($Items, [Text.Encoding]::UTF8)) { if ($l.Trim()) { $rows += , ($l -split "`t") } }
$w = New-Object -ComObject Word.Application; $w.Visible = $false; $w.DisplayAlerts = 0; $w.ScreenUpdating = $false
$opt = $w.Options; $oldSpell = $opt.CheckSpellingAsYouType; $oldGram = $opt.CheckGrammarAsYouType
$opt.CheckSpellingAsYouType = $false; $opt.CheckGrammarAsYouType = $false
$lines = @()
try {
  $d = $w.Documents.Open($File, $false, $true)
  foreach ($r in $rows) {
    $kind = $r[0]; $text = $r[1]; $why = if ($r.Count -gt 2) { $r[2] } else { '' }
    $key = ($text -replace '&gt;', '>' -replace '&lt;', '<' -replace '&amp;', '&' -replace '&quot;', '"').Trim()
    $z = $key.IndexOf([char]0x200B); if ($z -gt 8) { $key = $key.Substring(0, $z) } elseif ($z -ge 0) { $key = $key.Replace([string][char]0x200B, '') }
    $key = $key.Substring(0, [Math]::Min(40, $key.Length))
    $rg = $d.Content; $f = $rg.Find; $f.ClearFormatting(); $f.Text = $key; $f.MatchWildcards = $false; $f.Wrap = 0
    $page = '?'; $printed = '?'
    if ($f.Execute()) { $page = $rg.Information(3); $printed = $rg.Information(1) }
    $lines += [pscustomobject]@{ page = $(if ($page -is [int]) { $page } else { 9999 }); text = ("หน้า {0} (PDF {1}) [{2}] {3}{4}" -f $printed, $page, $kind, $key, $(if ($why) { " — $why" } else { '' })) }
  }
  $d.Close($false)
} finally {
  $opt.CheckSpellingAsYouType = $oldSpell; $opt.CheckGrammarAsYouType = $oldGram; $w.ScreenUpdating = $true; $w.Quit()
}
$body = ($lines | Sort-Object page | ForEach-Object { '- ' + $_.text }) -join "`r`n"
$head = "จุดที่ควรเข้าไปเช็ค (เรียงตามหน้า) — แก้อัตโนมัติแล้ว แต่ควรดูด้วยตาว่าโอเค`r`n"
[IO.File]::WriteAllText($Out, $head + $(if ($body) { $body } else { '- ไม่มี' }), (New-Object Text.UTF8Encoding $true))
"review list: $($lines.Count) item(s) -> $Out"
