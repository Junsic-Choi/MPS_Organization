# export_mps_more.ps1
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$ws = $xl.Workbooks.Item(1).Sheets.Item(4)
$data = $ws.Range("A101:G300").Value2
"--- MPS Sheet 4 Rows 101-300 ---" | Out-File "mps_headers_debug_v2.txt" -Encoding UTF8
for ($r = 1; $r -le 200; $r++) {
    $line = "R$( $r + 100 ) :"
    for ($c = 1; $c -le 7; $c++) {
        $v = if($data[$r, $c]){ (""+$data[$r, $c]).Trim() } else { "-" }
        $line += " [$c]=$v"
    }
    $line | Out-File "mps_headers_debug_v2.txt" -Append -Encoding UTF8
}
