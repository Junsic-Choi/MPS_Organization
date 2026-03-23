$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = Get-Location
$log = "$dir\dynamic_extract_v2_log.txt"
if (Test-Path $log) { Remove-Item $log }
Start-Transcript -Path $log -Force

try {
    $path = "$dir\data_working.xlsx"
    $workbook = $excel.Workbooks.Open($path, 0, $true)
    $ws = $workbook.Sheets.Item(2)
    
    $tCols = @()
    # Search Row 4 for "Production" label (Korean: 생산)
    # Using Unicode value for "생산" to avoid encoding issues
    # "생" = 0xC0DD, "산" = 0xC0B0
    $label = [char]0xC0DD + [char]0xC0B0

    for ($c = 1; $c -le 100; $c++) {
        $v4 = "$($ws.Cells.Item(4, $c).Value2)"
        $v3 = "$($ws.Cells.Item(3, $c).Value2)"
        if ($v4 -match $label) {
            # Skip August (8)
            # if ($v3 -match "8") { continue }
            
            $tCols += @{ idx = $c; month = $v3 }
            Write-Output "Found Col $c : Month=$v3"
        }
    }
    
    $results = @()
    $lastRow = $ws.UsedRange.Rows.Count
    Write-Output "Processing $lastRow rows..."
    
    $suffix = [char]0xC6D4 # "월"
    
    for ($r = 7; $r -le $lastRow; $r++) {
        $cellVal = $ws.Cells.Item($r, 3).Value2
        if ($null -eq $cellVal -or "$cellVal" -eq "") { continue }
        
        $model = "$cellVal".Trim()
        $site = "$($ws.Cells.Item($r, 1).Value2)"
        $group = "$($ws.Cells.Item($r, 2).Value2)"
        $rpm = "$($ws.Cells.Item($r, 4).Value2)"
        
        foreach ($col in $tCols) {
            $val = $ws.Cells.Item($r, $col.idx).Value2
            if ($null -ne $val -and [double]$val -gt 0) {
                $qty = [math]::Floor([double]$val)
                for ($q = 1; $q -le $qty; $q++) {
                    $results += [PSCustomObject]@{
                        Site    = $site
                        Group   = $group
                        Model   = $model
                        RPM     = $rpm
                        Month   = "$($col.month)$suffix"
                        Code    = ""
                        Product = ""
                    }
                }
            }
        }
    }
    
    $results | Export-Csv -Path "$dir\_FinalList.csv" -NoTypeInformation -Encoding UTF8
    Write-Output "SUCCESS. Total Rows: $($results.Count)"
}
catch {
    Write-Output "ERROR: $_"
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
    Stop-Transcript
}
