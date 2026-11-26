param([string]$File, [string]$Json)
# เติมเลขหน้าให้ตัวอย่างในรายงานการเปลี่ยนแปลง: หาตำแหน่งข้อความในเล่มที่จัดแล้ว -> เลขหน้า PDF และเลขหน้าที่พิมพ์
$cats = Get-Content -LiteralPath $Json -Raw -Encoding UTF8 | ConvertFrom-Json
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
try {
  $d = $app.Documents.Open($File, $false, $true); $all = $d.Content.Text
  # เริ่มหาหลังสารบัญ — ไม่งั้นหัวข้อไปเจอในสารบัญก่อน แล้วได้เลขหน้า 2 ทุกจุด
  $from = 0; try { if ($d.TablesOfContents.Count -gt 0) { $from = [math]::Min($all.Length - 1, $d.TablesOfContents.Item($d.TablesOfContents.Count).Range.End) } } catch {}
  foreach ($c in $cats) { foreach ($e in $c.examples) {
    $key = [string]$e.find; if (-not $key) { continue }
    $key = $key.Substring(0, [math]::Min(24, $key.Length)); $pos = $all.IndexOf($key, $from); if ($pos -lt 0) { $pos = $all.IndexOf($key) }
    if ($pos -lt 0) { continue }
    $r = $d.Range($pos, $pos)
    $e | Add-Member -NotePropertyName page -NotePropertyValue ([int]$r.Information(3)) -Force
    $e | Add-Member -NotePropertyName printed -NotePropertyValue ([string]$r.Information(1)) -Force
  } }
  $d.Close($false)
} finally { $app.Quit() }
[IO.File]::WriteAllText($Json, ($cats | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding $false))
"pages located"
