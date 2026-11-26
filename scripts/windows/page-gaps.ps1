# page-gaps.ps1 — หาหน้าที่มีที่ว่างท้ายหน้ามากผิดปกติ (จากภาพ PNG ของแต่ละหน้า)   *** UTF-8 with BOM ***
# ใช้: page-gaps.ps1 -Dir โฟลเดอร์ภาพหน้า [-Min 0.30]  -> "pNNN ว่างท้ายหน้า XX%"
param([string]$Dir, [double]$Min = 0.30)
Add-Type -AssemblyName System.Drawing
foreach ($f in (Get-ChildItem "$Dir\p*.png" | Sort-Object Name)) {
  $bmp = [System.Drawing.Bitmap]::FromFile($f.FullName)
  try {
    $h = $bmp.Height; $w = $bmp.Width
    $bodyBottom = [int]($h * 0.90)          # ignore footer / page number zone
    $last = 0; $top = [int]($h * 0.16)   # below the page header (logo + project line)
    for ($y = $bodyBottom; $y -gt $top; $y -= 2) {
      $ink = $false
      for ($x = [int]($w * 0.08); $x -lt [int]($w * 0.92); $x += 3) { $c = $bmp.GetPixel($x, $y); if (($c.R + $c.G + $c.B) -lt 600) { $ink = $true; break } }
      if ($ink) { $last = $y; break }
    }
    if ($last -eq 0) { "{0} หน้าว่าง (มีแต่หัว/ท้ายกระดาษ) — ต้องแก้" -f $f.BaseName; continue }
    $gap = ($bodyBottom - $last) / ($bodyBottom - $top)
    if ($gap -ge $Min) { "{0} ว่างท้ายหน้า {1:P0}" -f $f.BaseName, $gap }
  } finally { $bmp.Dispose() }
}
