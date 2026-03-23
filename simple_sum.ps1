try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $dir = Get-Location
    $sourcePath = "$dir\일반비_MPS2603-1(생산배포용).xlsx"
    $workPath = "$dir\data_working.xlsx"
    
    if (Test-Path $sourcePath) {
        Copy-Item -Path $sourcePath -Destination $workPath -Force
    }
    
    $workbook = $excel.Workbooks.Open($workPath, 0, $true)
    $ws = $workbook.Sheets.Item(2)
    
    $targetCols = @(5, 8, 9, 10, 11, 13) # Only Production!
    
    $sum = 0
    $currModel = ""
    for ($r = 7; $r -le 2000; $r++) {
        $model = $ws.Cells.Item($r, 3).Value2
        if ($null -ne $model -and "$model".Trim() -ne "") {
            $currModel = "$model".Trim()
        }
        
        foreach ($c in $targetCols) {
            $val = $ws.Cells.Item($r, $c).Value2
            # IF currModel is valid, then we consider this row's quantity valid!
            if ($null -ne $val -and [double]$val -gt 0 -and $currModel -ne "") {
                $sum += [double]$val
            }
        }
    }
    
    $workbook.Close($false)
    $excel.Quit()
    
    "Production (Feb-Jul) with Model Propagation = $sum" | Out-File "$dir\diag_sum_out.txt" -Encoding UTF8
}
catch {
    "ERROR: $_" | Out-File "$dir\diag_sum_out.txt" -Encoding UTF8
}
