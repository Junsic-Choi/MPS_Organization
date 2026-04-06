# find_nhm.ps1
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$ws = $xl.Workbooks.Item(1).Sheets.Item(4)
$data = $ws.Range("A1:G1000").Value2
"--- NHM Search Results in Sheet 4 ---" | Out-File "nhm_debug.txt" -Encoding UTF8
for ($r = 1; $r -le 1000; $r++) {
    for ($c = 4; $c -le 5; $c++) { # Model or Product
        $v = if($data[$r, $c]){ (""+$data[$r, $c]).Trim() } else { "" }
        if ($v.Contains("NHM")) {
            "R$r C$c: [$v] (C4=$($data[$r,4]), C5=$($data[$r,5]))" | Out-File "nhm_debug.txt" -Append -Encoding UTF8
        }
    }
}
