param([string]$Orig)
# เล่ม (6) รายงานผลการทดสอบ: ตารางยาวต่อโมดูล (แถวทดสอบ + ผลที่พบ + รูป) — หัวตารางซ้ำทุกหน้า
$ErrorActionPreference = 'Stop'
$SP = Split-Path $MyInvocation.MyCommand.Path; $D = (Join-Path $env:USERPROFILE 'Downloads'); $Key = 6
$W = "Test ID|Case ID|Feature|Description|Category|Step|Input Data|Expected Result|Test Result=2.4,1.3,2.6,3.0,2.0,5.0,3.6,3.2,1.3"
$pre = Join-Path $D "_pre6.docx"; Copy-Item -LiteralPath $Orig $pre -Force
"== 0) InPut Data -> Input Data"
node -e "import('file:///$((Join-Path (Split-Path $SP) 'lib\docx.mjs') -replace '\\','/')').then(m=>{const d=m.loadDocx(process.argv[1]);let x=d.get('word/document.xml');let n=0;x=x.replace(/>(\s*)InPut(\s*)</g,(a,b,c)=>{n++;return '>'+b+'Input'+c+'<'});x=x.replace(/>InPut Data</g,()=>{n++;return '>Input Data<'});d.set('word/document.xml',x);m.saveDocx(d,process.argv[1]);console.log('InPut fixed:',n)})" $pre
& powershell -NoProfile -ExecutionPolicy Bypass -File "$SP\book_all.ps1" -Key $Key -Orig $pre -Widths $W
Remove-Item -LiteralPath $pre -ErrorAction SilentlyContinue
$f = (Get-ChildItem -LiteralPath $D -Filter "($Key) *.docx" | Select-Object -First 1).FullName
"== 9) test-table extras (header row keeps repeating)"
node "$SP\steps_list.mjs" $f Step
node "$SP\compound_breaks.mjs" $f --dense-back
node "$SP\tag_thai.mjs" $f --no-compound
node "$SP\table_cleanup.mjs" $f
node "$SP\keep_heads.mjs" $f
node "$SP\pic_rows.mjs" $f
$app = New-Object -ComObject Word.Application; $app.Visible = $false; $app.DisplayAlerts = 0
try { $d = $app.Documents.Open($f); $n = 0
  foreach ($tb in $d.Tables) { $c0 = 0; try { $c0 = $tb.Columns.Count } catch { $c0 = 9 }; if ($c0 -ge 7) { $n++; foreach ($c in $tb.Range.Cells) { if ($c.RowIndex -gt 1) { $c.Range.Font.Size = 13; $c.Range.Font.SizeBi = 13 } } } }
  "  wide tables body at 13pt: $n"; $d.Save(); $d.Close() } finally { $app.Quit() }
& powershell -NoProfile -ExecutionPolicy Bypass -File "$SP\enlarge_images.ps1" -File $f
& powershell -NoProfile -ExecutionPolicy Bypass -File "$SP\add_pagenum.ps1" -File $f
"done: $f"
