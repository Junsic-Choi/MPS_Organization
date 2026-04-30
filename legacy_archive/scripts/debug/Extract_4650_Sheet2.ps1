$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = Get-Location
$logPath = "$dir\ps_extract_log.txt"
Start-Transcript -Path $logPath -Force

try {
    $path = "$dir\data_working.xlsx"
    Write-Output "Opening workbook: $path"
    $workbook = $excel.Workbooks.Open($path)
    $ws = $workbook.Sheets.Item(2) # 생산배포용
    
    Write-Output "Scanning Row 4 for '생산' columns..."
    $targetCols = @()
    for ($c = 5; $c -le 100; $c++) {
        $v4 = "$($ws.Cells.Item(4, $c).Value2)"
        $v3 = "$($ws.Cells.Item(3, $c).Value2)"
        if ($v4 -like "*생산*") {
            $targetCols += @{ idx = $c; month = $v3 }
            Write-Output "  Found Col $c : Month=$v3, Label=$v4"
        }
    }
    
    $results = @()
    $lastRow = $ws.UsedRange.Rows.Count
    Write-Output "Processing $lastRow rows..."
    
    for ($r = 7; $r -le $lastRow; $r++) {
        $model = "$($ws.Cells.Item($r, 3).Value2)".Trim()
        if ($model -eq "") { continue }
        
        $site = "$($ws.Cells.Item($r, 1).Value2)"
        $group = "$($ws.Cells.Item($r, 2).Value2)"
        $rpm = "$($ws.Cells.Item($r, 4).Value2)"
        
        foreach ($col in $targetCols) {
            $val = $ws.Cells.Item($r, $col.idx).Value2
            if ($null -ne $val -and $val -gt 0) {
                $qty = [int]$val
                for ($q = 1; $q -le $qty; $q++) {
                    $results += [PSCustomObject]@{
                        Site    = $site
                        Group   = $group
                        Model   = $model
                        RPM     = $rpm
                        Month   = $col.month
                        Code    = ""
                        Product = ""
                    }
                }
            }
        }
    }
    
    Write-Output "Total Rows Extracted: $($results.Count)"
    $results | Export-Csv -Path "$dir\_FinalList_4650.csv" -NoTypeInformation -Encoding UTF8
}
catch {
    Write-Output "ERROR: $_"
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
    Stop-Transcript
}
