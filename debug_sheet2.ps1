# debug_sheet2.ps1
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$wb = $xl.Workbooks.Item(1)
$ws = $wb.Sheets.Item(2)
$data = $ws.Range("A1:M20").Value2
"--- Sheet 2 Top 20 Rows ---" | Out-File "c:\Users\i0215099\Desktop\MPS_UPDATE\sheet2_debug.txt" -Encoding UTF8
for ($r = 1; $r -le 20; $r++) {
    $line = "Row $r :"
    for ($c = 1; $c -le 13; $c++) {
        $v = if ($data[$r, $c]) { ("" + $data[$r, $c]).Trim() } else { "-" }
        $line += " [$c]=$v"
    }
    $line | Out-File "c:\Users\i0215099\Desktop\MPS_UPDATE\sheet2_debug.txt" -Append -Encoding UTF8
}
Write-Host "Done -> sheet2_debug.txt"
