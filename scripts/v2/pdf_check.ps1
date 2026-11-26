param([string]$Pdf, [string]$Out)
# ตรวจ PDF ที่มีอยู่แบบอ่านอย่างเดียว: หน้าว่าง / หัวข้อหรือหัวตารางค้างท้ายหน้า / ล้นขอบขวา / ช่องว่างกลางหน้า / หน้าที่เริ่มต่ำผิดปกติ / หน้าที่ใช้น้อยกว่า 25%
Remove-Item $Out -Recurse -ea 0; New-Item -ItemType Directory $Out | Out-Null
powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path (Split-Path (Split-Path $MyInvocation.MyCommand.Path)) 'windows\pdf-to-png.ps1') -Pdf $Pdf -OutDir $Out -Width 400 | Out-Null
Add-Type -AssemblyName System.Drawing
$empty = @(); $endHdr = @(); $overR = @(); $gaps = @(); $low = @(); $short = @(); $starts = @{}
foreach ($fi in (Get-ChildItem $Out -Filter "p*.png" | Sort-Object Name)) {
  $b = [Drawing.Bitmap]::FromFile($fi.FullName); $H = $b.Height; $W = $b.Width; $pg = [int]$fi.BaseName.Substring(1)
  $landscape = $W -gt $H
  $rows = @(); for ($y = [int]($H * 0.11); $y -lt [int]($H * 0.90); $y++) { $ink = $false; for ($x = 30; $x -lt $W - 30; $x += 2) { $c = $b.GetPixel($x, $y); if ($c.R -lt 251 -or $c.G -lt 251 -or $c.B -lt 251) { $ink = $true; break } }; $rows += $ink }
  $first = [array]::IndexOf($rows, $true); $last = [array]::LastIndexOf($rows, $true)
  if ($pg -gt 2) {
    if ($last -lt 0 -or $last -lt $rows.Count * 0.05) { $empty += $pg }
    elseif ($last -lt $rows.Count * 0.25) { $short += $pg }
    $run = 0; $best = 0; for ($i = [math]::Max(0,$first); $i -le $last; $i++) { if (-not $rows[$i]) { $run++; if ($run -gt $best) { $best = $run } } else { $run = 0 } }
    if ($best -gt 30) { $gaps += $pg }
    $starts[$pg] = $first
    $yl = [int]($H * 0.11) + $last; $dark = 0; $blue = 0; $tint = 0; $tot = 0
    for ($y = [math]::Max(0, $yl - 12); $y -le $yl; $y++) { for ($x = 30; $x -lt $W - 30; $x += 2) { $c = $b.GetPixel($x, $y); $tot++; if ($c.R -lt 150) { $dark++; if ($c.B -gt $c.R + 40) { $blue++ } }; if ($c.B -gt $c.R + 15 -and $c.R -gt 150) { $tint++ } } }
    if (($dark -gt 10 -and $blue / $dark -gt 0.6) -or ($tint / [math]::Max(1, $tot) -gt 0.4)) { $endHdr += $pg }
    $mr = $(if ($landscape) { 1.6 / 29.7 } else { 1.6 / 21 }); $xr = [int]($W * (1 - $mr)) + 3; $hit = 0
    for ($y = [int]($H * 0.12); $y -lt [int]($H * 0.92); $y += 2) { for ($x = $xr; $x -lt $W - 2; $x++) { if ($b.GetPixel($x, $y).R -lt 150) { $hit++ } } }
    if ($hit -gt 5) { $overR += $pg }
  }
  $b.Dispose()
}
$mode = ($starts.Values | Group-Object | Sort-Object Count -Descending | Select-Object -First 1).Name
$low = $starts.GetEnumerator() | Where-Object { $_.Value -gt [int]$mode + 20 } | Sort-Object Name | ForEach-Object { $_.Name }
"pages: " + (Get-ChildItem $Out -Filter p*.png).Count
"blank: " + ($empty -join ',')
"heading/table header at page end: " + ($endHdr -join ',')
"past right margin: " + ($overR -join ',')
"big white gap mid-page: " + ($gaps -join ',')
"starts lower than normal: " + ($low -join ',')
"less than 25% used: " + ($short -join ',')
