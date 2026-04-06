# debug_find_model.ps1
try {
    $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    $wb = $xl.Workbooks.Item(1)
    $ws = $wb.Sheets.Item(4)
    $data = $ws.Range("A1:AD1000").Value2
    "--- Searching for NHM5000 in Sheet 4 ---" | Out-File "debug_search.txt" -Encoding UTF8
    for ($r = 1; $r -le 1000; $r++) {
        for ($c = 1; $c -le 30; $c++) {
            $v = if($data[$r,$c]){ (""+$data[$r,$c]).Trim() } else { "" }
            if ($v.Contains("NHM5000") -or $v.Contains("NHM 5000")) {
                "Found at R$r C$c: [$v]" | Out-File "debug_search.txt" -Append -Encoding UTF8
            }
        }
    }
    Write-Host "Done."
} catch {
    Write-Host "Error: $($_.Exception.Message)"
}
