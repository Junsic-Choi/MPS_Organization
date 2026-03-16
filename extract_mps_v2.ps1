$ErrorActionPreference = "Stop"
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $dir = "C:\Users\i0215099\Desktop\MPS_UPDATE"
    $files = Get-ChildItem -Path $dir -Filter "*.xlsx" | Where-Object { $_.Name -like "*MPS*" -and $_.Name -like "*(생산배포용)*" }
    $path = $files[0].FullName
    
    $tempPath = "$dir\temp_mps_debug.xls"
    Copy-Item $path $tempPath -Force
    
    $workbook = $excel.Workbooks.Open($tempPath, 0, $true)
    
    $targetSheet = $null
    foreach ($s in $workbook.Sheets) {
        if ($s.Name -eq "MPS") {
            $targetSheet = $s
            break
        }
    }
    
    if (-not $targetSheet) { $targetSheet = $workbook.Sheets.Item(4) }
    
    $result = "Sheet: $($targetSheet.Name)`n"
    for ($r = 1; $r -le 50; $r++) {
        $rowText = ""
        for ($c = 1; $c -le 20; $c++) {
            $val = $targetSheet.Cells.Item($r, $c).Text
            $rowText += "$val|"
        }
        $result += "$rowText`n"
    }
    
    $result | Out-File -FilePath "$dir\mps_debug_output.txt" -Encoding UTF8
    
    $workbook.Close($false)
    $excel.Quit()
}
catch {
    Write-Host "Error: $_"
}
finally {
    if ($excel) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
}
