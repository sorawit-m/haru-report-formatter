param([string]$File, [string]$Orig)
# ภาพหน้าจอ + รูปซ้อน: ย่อทั้งคู่ด้วยอัตราส่วนเดียวกัน ให้ตำแหน่งสัมพันธ์เหมือนต้นฉบับ และไม่เกินขอบขวา
$tmp = (Join-Path $env:TEMP '_orig_tmp.docx'); Copy-Item -LiteralPath $Orig $tmp -Force
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
function Get-Pairs($doc) {
  $out = @()
  foreach ($s in $doc.Shapes) {
    if ($s.Type -ne 13 -or $s.Anchor.Information(3) -eq 1) { continue }
    $pg = $s.Anchor.Information(3); $ps = $s.Anchor.Sections.Item(1).PageSetup
    $r = $doc.GoTo(1, 1, $pg); $r2 = $doc.GoTo(1, 1, $pg + 1); $img = $doc.Range($r.Start, $r2.Start).InlineShapes.Item(1)
    $sx = $s.Left; if ($s.RelativeHorizontalPosition -eq 0) { $sx += $ps.LeftMargin }
    $out += [pscustomobject]@{ s = $s; img = $img; imgX = $img.Range.Information(5); imgW = $img.Width; sx = $sx; sw = $s.Width; top = $s.Top; right = $ps.PageWidth - $ps.RightMargin }
  }
  return ,$out
}
try {
  $o = $app.Documents.Open($tmp, $false, $true); $oa = Get-Pairs $o
  $n = $app.Documents.Open($File); $na = Get-Pairs $n
  for ($i = 0; $i -lt $na.Count; $i++) {
    $a = $oa[$i]; $b = $na[$i]
    $off = $a.sx - $a.imgX
    $k = [math]::Min(($b.right - 2 - $b.imgX) / $a.imgW, ($b.right - 2 - $b.imgX) / ($off + $a.sw))
    $b.img.LockAspectRatio = -1; $b.img.Width = $a.imgW * $k
    $s = $b.s; try { $s.LayoutInCell = 0 } catch {}; $s.LockAspectRatio = -1; $s.RelativeHorizontalPosition = 1
    $s.Width = $a.sw * $k; $s.Left = $b.imgX + $off * $k; $s.Top = $a.top * $k
    "pair {0}: scale {1:N3} | screenshot {2:N0}pt wide | overlay left {3:N0} right {4:N0} (margin {5:N0})" -f ($i + 1), $k, $b.img.Width, $s.Left, ($s.Left + $s.Width), $b.right
  }
  $o.Close($false); $n.Save(); $n.Close()
} finally { $app.Quit(); Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue }
