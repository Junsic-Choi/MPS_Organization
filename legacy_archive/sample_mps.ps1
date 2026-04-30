# sample_mps.ps1
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$wb = $xl.Workbooks.Item(1)
$ws = $wb.Sheets.Item(4)
$data = $ws.Range("A1:E500").Value2
"--- MPS Sheet Column 4 & 5 Sample ---" | Out-File "c:\Users\i0215099\Desktop\MPS_UPDATE\mps_sample.txt" -Encoding UTF8
for ($r = 1; $r -le 100; $r++) {
    $c4 = if ($data[$r, 4]) { ("" + $data[$r, 4]).Trim() } else { "-" }
    $c5 = if ($data[$r, 5]) { ("" + $data[$r, 5]).Trim() } else { "-" }
    if ($c4 -ne "-" -or $c5 -ne "-") {
        "Row $r : C4=[$c4] C5=[$c5]" | Out-File "c:\Users\i0215099\Desktop\MPS_UPDATE\mps_sample.txt" -Append -Encoding UTF8
    }
}
Write-Host "Done -> mps_sample.txt"
