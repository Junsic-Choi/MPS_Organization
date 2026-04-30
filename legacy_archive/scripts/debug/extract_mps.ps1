$ErrorActionPreference = "Continue"
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $dir = "c:\Users\i0215099\Desktop\MPS_UPDATE"
    $files = Get-ChildItem -Path $dir -Filter "*MPS*.xlsx" | Where-Object { $_.Name -like "*(생산배포용)*" }
    if ($files.Count -eq 0) {
        $files = Get-ChildItem -Path $dir -Filter "*MPS*.xlsx"
    }
    $path = $files[0].FullName
    
    # Use simple local path for output to avoid permissions/path issues
    $outputPath = Join-Path $dir "mps_tab_debug.txt"
    Write-Host "Processing: $path"
    
    # Open as ReadOnly
    $workbook = $excel.Workbooks.Open($path, 0, $true)
    
    $targetSheet = $null
    foreach ($s in $workbook.Sheets) {
        if ($s.Name -eq "MPS") {
            $targetSheet = $s
            break
        }
    }
    
    if ($null -eq $targetSheet) {
        $targetSheet = $workbook.Sheets.Item(4)
    }
    
    Write-Host "Target Sheet: $($targetSheet.Name)"
    
    $result = "Sheet: $($targetSheet.Name)`n"
    
    for ($r = 1; $r -le 20; $r++) {
        $rowText = ""
        for ($c = 1; $c -le 15; $c++) {
            $val = $targetSheet.Cells.Item($r, $c).Text
            if ($null -ne $val) {
                $val = $val.Replace("`n", " ").Replace("`r", "")
            }
            $rowText += "$val|"
        }
        $result += "$rowText`n"
    }
    
    $result | Out-File -FilePath $outputPath -Encoding UTF8
    Write-Host "Saved to $outputPath"
    
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
catch {
    Write-Error $_.Exception.Message
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
