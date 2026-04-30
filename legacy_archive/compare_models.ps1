# compare_models.ps1 - Compare model names between sheets
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$wb = $xl.Workbooks.Item(1)

$out = "c:\Users\i0215099\Desktop\MPS_UPDATE\model_compare.txt"
"=== MPS Sheet (Sheet4) Col4 data from row 6 ===" | Out-File $out -Encoding UTF8
$ws4 = $wb.Sheets.Item(4)
$mpsRange = $ws4.Range("A1:Z50").Value2
for ($r = 6; $r -le 30; $r++) {
    $m = if ($mpsRange[$r,4]) { ("" + $mpsRange[$r,4]).Trim() } else { "" }
    $c = if ($mpsRange[$r,3]) { ("" + $mpsRange[$r,3]).Trim() } else { "" }
    $p = if ($mpsRange[$r,5]) { ("" + $mpsRange[$r,5]).Trim() } else { "" }
    if ($m -ne "") { "R$r Model=[$m] Code=[$c] Product=[$p]" | Out-File $out -Append -Encoding UTF8 }
}

"" | Out-File $out -Append -Encoding UTF8
"=== Production Sheet (Sheet2) Col3 first 20 model names ===" | Out-File $out -Append -Encoding UTF8
$ws2 = $wb.Sheets.Item(2)
$range = $ws2.Range("A1:D500").Value2
for ($r = 6; $r -le 200; $r++) {
    $m = if ($range[$r,3]) { ("" + $range[$r,3]).Trim() } else { "" }
    if ($m -ne "" -and $m -notmatch "계") { "[$m]" | Out-File $out -Append -Encoding UTF8 }
    if ((Get-Content $out).Count -gt 60) { break }
}
"Done" | Out-File $out -Append -Encoding UTF8
