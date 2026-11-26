param([string]$File, [string]$Out)
# อัปเดตสารบัญ ส่งออก PDF แล้วตรวจทุกหน้า: หน้าว่าง / หัวข้อหรือหัวตารางค้างท้ายหน้า / ล้นขอบขวา
$pdf = [IO.Path]::ChangeExtension($File, '.pdf')
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
try { $d = $app.Documents.Open($File); foreach ($t in $d.TablesOfContents) { $t.Update() }; $d.Save(); "pages " + $d.ComputeStatistics(2); $d.ExportAsFixedFormat($pdf, 17); $d.Close() } finally { $app.Quit() }
Remove-Item $Out -Recurse -ea 0; New-Item -ItemType Directory $Out | Out-Null
powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path (Split-Path (Split-Path $MyInvocation.MyCommand.Path)) 'windows\pdf-to-png.ps1') -Pdf $pdf -OutDir $Out -Width 400 | Out-Null
Add-Type -AssemblyName System.Drawing
$empty = @(); $blueEnd = @(); $overR = @()
foreach ($fi in (Get-ChildItem $Out -Filter "p*.png" | Sort-Object Name)) {
  $b = [Drawing.Bitmap]::FromFile($fi.FullName); $H = $b.Height; $W = $b.Width; $pg = [int]$fi.BaseName.Substring(1)
  $last = 0; for ($y = [int]($H * 0.12); $y -lt [int]($H * 0.92); $y++) { for ($x = 30; $x -lt $W - 30; $x += 3) { $c = $b.GetPixel($x, $y); if ($c.R -lt 200) { $last = $y; break } } }
  if ($pg -gt 1) {
    if ($last -lt $H * 0.15) { $empty += $pg }
    $blue = 0; $dark = 0; for ($y = [math]::Max(0, $last - 14); $y -le $last; $y++) { for ($x = 30; $x -lt $W - 30; $x += 2) { $c = $b.GetPixel($x, $y); if ($c.R -lt 150) { $dark++; if ($c.B -gt $c.R + 40) { $blue++ } } } }
    if ($dark -gt 10 -and $blue / $dark -gt 0.6) { $blueEnd += $pg }
    $tint = 0; $tot = 0; for ($y = [math]::Max(0, $last - 10); $y -le $last; $y++) { for ($x = 30; $x -lt $W - 30; $x += 3) { $c2 = $b.GetPixel($x, $y); $tot++; if ($c2.B -gt $c2.R + 15 -and $c2.R -gt 150) { $tint++ } } }; $tintEnd = ($tint / [math]::Max(1,$tot)) -gt 0.4; if ($tintEnd -and $blueEnd -notcontains $pg) { $blueEnd += $pg }
    $xr = [int]($W * ((21 - 1.6) / 21)) + 3; $hit = 0; for ($y = [int]($H * 0.12); $y -lt [int]($H * 0.92); $y += 2) { for ($x = $xr; $x -lt $W - 2; $x++) { $c = $b.GetPixel($x, $y); if ($c.R -lt 150) { $hit++ } } }
    if ($hit -gt 5) { $overR += $pg }
  }
  $b.Dispose()
}
"blank: " + ($empty -join ','); "ends with heading/table header: " + ($blueEnd -join ','); "past right margin: " + ($overR -join ',')
