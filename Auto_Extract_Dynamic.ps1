$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = Get-Location
$log = "$dir\dynamic_extract_log.txt"
"Starting Dynamic Extraction..." | Out-File $log -Encoding UTF8

try {
    $path = "$dir\data_working.xlsx"
    $workbook = $excel.Workbooks.Open($path, 0, $true)
    $ws = $workbook.Sheets.Item(2)
    
    # 1. Detect target columns dynamically from Row 4
    $targetCols = @()
    for ($c = 1; $c -le 100; $c++) {
        $v4 = "$($ws.Cells.Item(4, $c).Text)"
        $v3 = "$($ws.Cells.Item(3, $c).Text)"
        if ($v4 -like "*생산*") {
            # Skip August if found
            if ($v3 -like "*8월*") { continue }
            
            $targetCols += @{ idx = $c; month = $v3 }
            "Found '생산' Col $c : Month=$v3" | Out-File $log -Append -Encoding UTF8
        }
    }
    
    $results = @()
    $lastRow = $ws.UsedRange.Rows.Count
    "Processing $lastRow rows..." | Out-File $log -Append -Encoding UTF8
    
    for ($r = 7; $r -le $lastRow; $r++) {
        $cellVal = $ws.Cells.Item($r, 3).Value2 # Model name
        if ($null -eq $cellVal) { continue }
        
        $model = "$cellVal".Trim()
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
    
    $results | Export-Csv -Path "$dir\_FinalList_Dynamic.csv" -NoTypeInformation -Encoding UTF8
    "SUCCESS. Total Rows: $($results.Count)" | Out-File $log -Append -Encoding UTF8
    
    # Also overwrite the target file for dashboard
    Copy-Item "$dir\_FinalList_Dynamic.csv" "$dir\_FinalList.csv" -Force
}
catch {
    "ERROR: $_" | Out-File $log -Append -Encoding UTF8
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
}
