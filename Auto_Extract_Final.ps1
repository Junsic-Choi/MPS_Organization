$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = Get-Location
$log = "$dir\final_delivery_log.txt"
if (Test-Path $log) { Remove-Item $log }
Start-Transcript -Path $log -Force

try {
    $sourcePath = "$dir\일반비_MPS2603-1(생산배포용).xlsx"
    $workPath = "$dir\data_working.xlsx"
    if (Test-Path $sourcePath) {
        Copy-Item -Path $sourcePath -Destination $workPath -Force
        Write-Output "Copied original file to data_working.xlsx for safe extraction."
    }
    
    $workbook = $excel.Workbooks.Open($workPath, 0, $true)
    $ws = $workbook.Sheets.Item(2)
    
    $tCols = @()
    $labelProd = [char]0xC0DD + [char]0xC0B0 # "생산"
    $labelSales = [char]0xD310 + [char]0xB9E4 # "판매"

    for ($c = 1; $c -le 100; $c++) {
        $v4 = "$($ws.Cells.Item(4, $c).Value2)"
        $v3 = "$($ws.Cells.Item(3, $c).Value2)"
        if ($v4 -match $labelProd -or $v4 -match $labelSales) {
            # Skip August (8)
            if ($v3 -match "8") { continue }
            $tCols += @{ idx = $c; month = $v3; cat = $v4 }
            Write-Output "Found Col $c : Month=$v3, Category=$v4"
        }
    }
    
    $results = @()
    $lastRow = $ws.UsedRange.Rows.Count
    Write-Output "Processing $lastRow rows..."
    
    $suffix = [char]0xC6D4 # "월"
    $currSite = ""; $currGroup = ""; $currModel = ""; $currRPM = ""
    $consecutiveEmpty = 0

    for ($r = 7; $r -le $lastRow; $r++) {
        $vSite = $ws.Cells.Item($r, 1).Value2
        $vGroup = $ws.Cells.Item($r, 2).Value2
        $vModel = $ws.Cells.Item($r, 3).Value2
        $vRPM = $ws.Cells.Item($r, 4).Value2

        if ($null -eq $vSite -and $null -eq $vGroup -and $null -eq $vModel -and $null -eq $vRPM) {
            # Special check: also check if ALL target columns are empty in this row
            $rowHasData = $false
            foreach ($col in $tCols) {
                if ($null -ne $ws.Cells.Item($r, $col.idx).Value2) { $rowHasData = $true; break }
            }
            if (-not $rowHasData) {
                $consecutiveEmpty++
                if ($consecutiveEmpty -gt 20) { break }
                continue
            }
        }
        $consecutiveEmpty = 0

        # Propagate Model Info
        if ($null -ne $vSite -and "$vSite" -ne "") { $currSite = "$vSite".Trim() }
        if ($null -ne $vGroup -and "$vGroup" -ne "") { $currGroup = "$vGroup".Trim() }
        if ($null -ne $vModel -and "$vModel" -ne "") { $currModel = "$vModel".Trim() }
        if ($null -ne $vRPM -and "$vRPM" -ne "") { $currRPM = "$vRPM".Trim() }

        if ($currModel -ne "") {
            foreach ($col in $tCols) {
                # DO NOT PROPAGATE QUANTITIES - use only the actual cell value
                $val = $ws.Cells.Item($r, $col.idx).Value2
                if ($null -ne $val -and [double]$val -gt 0) {
                    $qty = [math]::Floor([double]$val)
                    for ($q = 1; $q -le $qty; $q++) {
                        $results += [PSCustomObject]@{
                            Site    = $currSite
                            Group   = $currGroup
                            Model   = $currModel
                            RPM     = $currRPM
                            Month   = "$($col.month)$suffix"
                            Code    = ""
                            Product = ""
                        }
                    }
                }
            }
        }
    }
    
    $results | Export-Csv -Path "$dir\_FinalList_4650_Latest.csv" -NoTypeInformation -Encoding UTF8
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
