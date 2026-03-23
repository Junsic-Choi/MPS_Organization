$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = Get-Location
$log = "$dir\final_mps_4650_log.txt"
if (Test-Path $log) { Remove-Item $log }
Start-Transcript -Path $log -Force

try {
    $sourcePath = "$dir\일반비_MPS2603-1(생산배포용).xlsx"
    $path = "$dir\data_working.xlsx"
    if (Test-Path $sourcePath) {
        Copy-Item -Path $sourcePath -Destination $path -Force
    }
    $workbook = $excel.Workbooks.Open($path, 0, $true)
    $ws = $workbook.Sheets.Item(4) # MPS Sheet
    
    # User's referenced columns: I=9, M=13, R=18, W=23, AC=29, AI=35
    $targetCols = @(
        @{ idx = 9; month = "2" },
        @{ idx = 13; month = "3" },
        @{ idx = 18; month = "4" },
        @{ idx = 23; month = "5" },
        @{ idx = 29; month = "6" },
        @{ idx = 35; month = "7" }
    )
    
    $results = @()
    $lastRow = $ws.UsedRange.Rows.Count
    Write-Output "Processing $lastRow rows on Sheet 4 (MPS)..."
    
    $suffix = [char]0xC6D4 # "월"
    
    $currSite = ""
    $currGroup = ""
    $currModel = ""
    $currRPM = ""
    $consecutiveEmpty = 0

    # Data starts from Row 7 (standard layout)
    for ($r = 7; $r -le $lastRow; $r++) {
        $vSite = $ws.Cells.Item($r, 1).Value2
        $vGroup = $ws.Cells.Item($r, 2).Value2
        $vModel = $ws.Cells.Item($r, 3).Value2
        $vRPM = $ws.Cells.Item($r, 4).Value2

        if ($null -eq $vSite -and $null -eq $vGroup -and $null -eq $vModel -and $null -eq $vRPM) {
            $consecutiveEmpty++
            if ($consecutiveEmpty -gt 50) { break }
            continue
        }
        $consecutiveEmpty = 0

        if ($null -ne $vSite -and "$vSite" -ne "") { $currSite = "$vSite".Trim() }
        if ($null -ne $vGroup -and "$vGroup" -ne "") { $currGroup = "$vGroup".Trim() }
        if ($null -ne $vModel -and "$vModel" -ne "") { $currModel = "$vModel".Trim() }
        if ($null -ne $vRPM -and "$vRPM" -ne "") { $currRPM = "$vRPM".Trim() }

        if ($currModel -ne "") {
            foreach ($col in $targetCols) {
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
    
    $results | Export-Csv -Path "$dir\_FinalList_MPS.csv" -NoTypeInformation -Encoding UTF8
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
