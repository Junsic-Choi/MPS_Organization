# export_mps_small.ps1
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$ws = $xl.Workbooks.Item(1).Sheets.Item(4)
$data = $ws.Range("A1:G100").Value2
"--- MPS Sheet 4 Top 100 Rows ---" | Out-File "mps_headers_debug.txt" -Encoding UTF8
for ($r = 1; $r -le 100; $r++) {
    $line = "R$r :"
    for ($c = 1; $c -le 7; $c++) {
        $v = if($data[$r, $c]){ (""+$data[$r, $c]).Trim() } else { "-" }
        $line += " [$c]=$v"
    }
    $line | Out-File "mps_headers_debug.txt" -Append -Encoding UTF8
}
