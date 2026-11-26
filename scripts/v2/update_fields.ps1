param([string[]]$Keys)
# Pass 2: อัปเดตเลข SEQ (รูป/ตาราง) แล้วสร้างสารบัญ / สารบัญรูป / สารบัญตารางใหม่
$ErrorActionPreference = 'Stop'
$D = (Join-Path $env:USERPROFILE 'Downloads')
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
try {
  foreach ($k in $Keys) {
    $f = Get-ChildItem -LiteralPath $D -Filter "($k) *.docx" | Select-Object -First 1
    $t0 = Get-Date; $doc = $app.Documents.Open($f.FullName)
    $doc.Content.Fields.Update() | Out-Null
    foreach ($t in $doc.TablesOfContents) { $t.Update() }
    foreach ($t in $doc.TablesOfFigures) { $t.Update() }
    $doc.Save(); $pages = $doc.ComputeStatistics(2); $doc.Close()
    "($k) updated: $pages pages, {0:N0}s" -f ((Get-Date) - $t0).TotalSeconds
  }
} finally { $app.Quit() }
