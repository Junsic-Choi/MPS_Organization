$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = Get-Location
$log = "$dir\dynamic_extract_v3_log.txt"
if (Test-Path $log) { Remove-Item $log }
Start-Transcript -Path $log -Force

try {
    $path = "$dir\data_working.xlsx"
    $workbook = $excel.Workbooks.Open($path, 0, $true)
    $ws = $workbook.Sheets.Item(2)
    
    $tCols = @()
    $label = [char]0xC0DD + [char]0xC0B0 # "생산"

    $suffix = [char]0xC6D4 # "월"
    $excludeLabel = "8" + $suffix

    $lastRow = $ws.UsedRange.Rows.Count
    Write-Output "Loading Excel Data into Memory for Speed..."
    $dataMatrix = $ws.Range("A1", $ws.Cells.Item($lastRow, 100)).Value2

    for ($c = 1; $c -le 100; $c++) {
        $v4 = "$($dataMatrix[4, $c])"
        $v3 = "$($dataMatrix[3, $c])"
        if ($v4 -match $label -and $v3 -notmatch $excludeLabel) {
            $monthStr = "$v3".Trim()
            if ($monthStr -notmatch $suffix -and $monthStr -ne "") { $monthStr += $suffix }
            $tCols += @{ idx = $c; month = $monthStr }
            Write-Output "Found Col $c : Month=$monthStr"
        }
    }
    
    $results = [System.Collections.Generic.List[PSCustomObject]]::new()
    Write-Output "Processing $lastRow rows with value propagation..."
    
    # State variables for propagation
    $currSite = ""
    $currGroup = ""
    $currModel = ""
    $currRPM = ""

    for ($r = 7; $r -le $lastRow; $r++) {
        $vSite = $dataMatrix[$r, 1]
        $vGroup = $dataMatrix[$r, 2]
        $vModel = $dataMatrix[$r, 3]
        $vRPM = $dataMatrix[$r, 4]

        # Propagate if not null
        if ($null -ne $vSite -and "$vSite" -ne "") { $currSite = "$vSite".Trim() }
        if ($null -ne $vGroup -and "$vGroup" -ne "") { $currGroup = "$vGroup".Trim() }
        if ($null -ne $vModel -and "$vModel" -ne "") { $currModel = "$vModel".Trim() }
        if ($null -ne $vRPM -and "$vRPM" -ne "") { $currRPM = "$vRPM".Trim() }

        # If we have at least a model, process the quantities
        if ($currModel -ne "") {
            foreach ($col in $tCols) {
                $val = $dataMatrix[$r, $col.idx]
                $num = $val -as [double]
                if ($null -ne $num -and $num -gt 0) {
                    $qty = [math]::Floor($num)
                    for ($q = 1; $q -le $qty; $q++) {
                        $results.Add([PSCustomObject]@{
                                Site    = $currSite
                                Group   = $currGroup
                                Model   = $currModel
                                RPM     = $currRPM
                                Month   = "$($col.month)"
                                Code    = ""
                                Product = ""
                            })
                    }
                }
            }
        }
    }
    
    $results | Export-Csv -Path "$dir\_FinalList.csv" -NoTypeInformation -Encoding UTF8
    Write-Output "SUCCESS. Total Rows: $($results.Count)"
    if ($results.Count -eq 4650) {
        Write-Output "Verified: Expected 4650 rows reached."
    }
    else {
        Write-Output "Warning: Target row count was 4650, but found $($results.Count) rows."
    }
}
catch {
    Write-Output "ERROR: $_"
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
    Stop-Transcript
}
