# find_hc.ps1
$xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
$ws = $xl.Workbooks.Item(1).Sheets.Item(4)
$data = $ws.Range("A1:G1200").Value2
"--- HC/NHC Search ---" | Out-File "hc_debug.txt" -Encoding UTF8
for ($r = 1; $r -le 1200; $r++) {
    $v4 = if($data[$r,4]){ (""+$data[$r,4]).Trim() } else { "" }
    $v5 = if($data[$r,5]){ (""+$data[$r,5]).Trim() } else { "" }
    if ($v4 -match "HC|NHC" -or $v5 -match "HC|NHC") {
        "R$r: C4=[$v4] C5=[$v5]" | Out-File "hc_debug.txt" -Append -Encoding UTF8
    }
}
