$ErrorActionPreference = "Stop"
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $dir = "C:\Users\i0215099\Desktop\MPS_UPDATE"
    $files = Get-ChildItem -Path $dir -Filter "*MPS2603-1*.xlsx"
    if ($files.Count -eq 0) {
        throw "File not found"
    }
    $path = $files[0].FullName
    
    Out-File -FilePath "$dir\com_result.txt" -InputObject "Opening $path..." -Encoding UTF8
    
    $workbook = $excel.Workbooks.Open($path)
    Out-File -FilePath "$dir\com_result.txt" -InputObject "Workbook opened." -Append -Encoding UTF8
    
    $sheetsText = "Sheets: "
    $targetSheet = $null
    foreach ($sheet in $workbook.Sheets) {
        $sheetsText += $sheet.Name + ", "
        $targetSheet = $sheet # We want the first sheet if nothing matches
    }
    Out-File -FilePath "$dir\com_result.txt" -InputObject $sheetsText -Append -Encoding UTF8
    
    if ($workbook.Sheets.Count -ge 2) {
        $targetSheet = $workbook.Sheets.Item(1)
        foreach ($s in $workbook.Sheets) {
            if ($s.Name -like "*생산*") {
                $targetSheet = $s
            }
        }
    }
    
    Out-File -FilePath "$dir\com_result.txt" -InputObject "Target sheet: $($targetSheet.Name)" -Append -Encoding UTF8
    
    $range = $targetSheet.UsedRange
    $maxRows = $range.Rows.Count
    $maxCols = $range.Columns.Count
    if ($maxCols -gt 50) { $maxCols = 50 }
    if ($maxRows -gt 5) { $maxRows = 5 }
    
    $result = "Rows: $($range.Rows.Count), Cols: $($range.Columns.Count)`n"
    
    for ($r = 1; $r -le $maxRows; $r++) {
        $rowText = ""
        for ($c = 1; $c -le $maxCols; $c++) {
            $val = $range.Item($r, $c).Text
            $rowText += "$val`t"
        }
        $result += "$rowText`n"
    }
    
    Out-File -FilePath "$dir\com_result.txt" -InputObject $result -Append -Encoding UTF8
    
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
catch {
    Out-File -FilePath "C:\Users\i0215099\Desktop\MPS_UPDATE\com_result.txt" -InputObject "Error: $_" -Append -Encoding UTF8
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
