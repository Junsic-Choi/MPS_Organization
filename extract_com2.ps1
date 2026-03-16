$ErrorActionPreference = "Stop"
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $dir = "C:\Users\i0215099\Desktop\MPS_UPDATE"
    $files = Get-ChildItem -Path $dir -Filter "*MPS2603-1*.xlsx"
    $path = $files[0].FullName
    
    $workbook = $excel.Workbooks.Open($path)
    
    $targetSheet = $null
    foreach ($s in $workbook.Sheets) {
        if ($s.Name -match "배포용") {
            $targetSheet = $s
        }
    }
    
    if (-not $targetSheet) { throw "Sheet not found" }
    
    $range = $targetSheet.UsedRange
    $maxRows = $range.Rows.Count
    $maxCols = $range.Columns.Count
    if ($maxCols -gt 50) { $maxCols = 50 }
    if ($maxRows -gt 15) { $maxRows = 15 } # Get more rows to find header
    
    $result = "Sheet: $($targetSheet.Name)`nRows: $($range.Rows.Count), Cols: $($range.Columns.Count)`n"
    
    for ($r = 1; $r -le $maxRows; $r++) {
        $rowText = ""
        for ($c = 1; $c -le $maxCols; $c++) {
            $val = $range.Item($r, $c).Text
            if ($val -eq $null) { $val = "" }
            # Remove linebreaks from value for easier reading
            $val = $val -replace "`n", " "
            $rowText += "$val|"
        }
        $result += "$rowText`n"
    }
    
    Out-File -FilePath "$dir\com_result2.txt" -InputObject $result -Encoding UTF8
    
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
catch {
    Out-File -FilePath "C:\Users\i0215099\Desktop\MPS_UPDATE\com_result2.txt" -InputObject "Error: $_" -Encoding UTF8
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
