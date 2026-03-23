$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = Get-Location
$log = "$dir\final_v2_log.txt"
"Starting Final Extraction V2..." | Out-File $log -Encoding UTF8

try {
    $path = "$dir\data_working.xlsx"
    if (-not (Test-Path $path)) { throw "data_working.xlsx not found" }
    
    $workbook = $excel.Workbooks.Open($path, 0, $true)
    $ws = $workbook.Sheets.Item(2) # 생산배포용
    
    # Mapping based on verified headers:
    # Col 5 (E): 2월
    # Col 8 (H): 3월
    # Col 9 (I): 4월
    # Col 10 (J): 5월
    # Col 11 (K): 6월
    # Col 13 (M): 7월
    $tCols = @(
        @{ idx = 5; m = "2월" },
        @{ idx = 8; m = "3월" },
        @{ idx = 9; m = "4월" },
        @{ idx = 10; m = "5월" },
        @{ idx = 11; m = "6월" },
        @{ idx = 13; m = "7월" }
    )
    
    $results = @()
    # Increased row limit to find all 4650 rows
    for ($r = 7; $r -le 10000; $r++) {
        $cellVal = $ws.Cells.Item($r, 3).Value2
        if ($null -eq $cellVal) {
            # Skip empty models but don't stop immediately unless many are empty
            $emptyCount++
            if ($emptyCount -gt 500) { break }
            continue
        }
        $emptyCount = 0
        
        $model = "$cellVal".Trim()
        $site = "$($ws.Cells.Item($r, 1).Value2)"
        $group = "$($ws.Cells.Item($r, 2).Value2)"
        $rpm = "$($ws.Cells.Item($r, 4).Value2)"
        
        foreach ($col in $tCols) {
            $val = $ws.Cells.Item($r, $col.idx).Value2
            if ($null -ne $val -and $val -gt 0) {
                $qty = [int]$val
                for ($q = 1; $q -le $qty; $q++) {
                    $results += [PSCustomObject]@{
                        Site    = $site
                        Group   = $group
                        Model   = $model
                        RPM     = $rpm
                        Month   = $col.m
                        Code    = ""
                        Product = ""
                    }
                }
            }
        }
    }
    
    $results | Export-Csv -Path "$dir\_FinalList.csv" -NoTypeInformation -Encoding UTF8
    "SUCCESS. Total Rows: $($results.Count)" | Out-File $log -Append -Encoding UTF8
}
catch {
    "ERROR: $_" | Out-File $log -Append -Encoding UTF8
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
}
