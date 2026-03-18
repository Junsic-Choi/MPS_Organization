$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $dir = $PSScriptRoot
    $file = Get-ChildItem -Path $dir -Filter "*MPS*(생산배포용)*.xlsx" | Select-Object -First 1
    if (-not $file) { 
        Write-Output "ERROR: MPS file not found in $dir"
        exit 
    }
    Write-Output "Opening: $($file.FullName)"
    $wb = $excel.Workbooks.Open($file.FullName, 0, $true)
    Write-Output "Sheets found: $(($wb.Sheets | ForEach-Object { $_.Name }) -join ', ')"
    
    $sh = $null
    foreach ($s in $wb.Sheets) { if ($s.Name -eq '생산배포용') { $sh = $s; break } }
    
    if ($sh) {
        Write-Output "Found '생산배포용' sheet. Checking Row 7..."
        for ($c = 1; $c -le 25; $c++) {
            $val = $sh.Cells.Item(7, $c).Text
            Write-Output "Column $c: [$val]"
        }
        
        Write-Output "Checking Row 8 (Sample Data)..."
        for ($c = 1; $c -le 25; $c++) {
            $val = $sh.Cells.Item(8, $c).Text
            Write-Output "Row 8 Column $c: [$val]"
        }
    }
    else {
        Write-Output "ERROR: '생산배포용' sheet not found!"
    }
    $wb.Close($false)
}
catch {
    Write-Output "EXCEPTION: $($_.Exception.Message)"
}
finally {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
