$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = Get-Location
$log = "$dir\extraction_full_log.txt"
"Starting MPS Tab Extraction..." | Out-File $log -Encoding UTF8

try {
    # Dynamically find the file to avoid Korean literal encoding issues
    $files = Get-ChildItem -Path $dir -Filter "*MPS2603-1*"
    $targetFile = $null
    foreach ($f in $files) {
        if ($f.Name -match "MPS2603" -and $f.Extension -eq ".xlsx") {
            $targetFile = $f
            break
        }
    }
    
    if ($null -eq $targetFile) {
        "Target Excel file not found!" | Out-File $log -Append -Encoding UTF8
        exit
    }
    
    $path = $targetFile.FullName
    "Opening file: $path" | Out-File $log -Append -Encoding UTF8
    
    $workbook = $excel.Workbooks.Open($path, 0, $true)
    
    $ws = $null
    foreach ($s in $workbook.Sheets) {
        if ($s.Name -match "MPS") {
            # Skip the '생산배포용' tab itself which might have MPS in name
            # by checking if it's purely 'MPS' or at least 'MPS' is major part
            $ws = $s
            break
        }
    }

    if ($null -eq $ws) {
        "MPS Tab not found!" | Out-File $log -Append -Encoding UTF8
        exit
    }

    "Found Target Sheet: $($ws.Name)" | Out-File $log -Append -Encoding UTF8
    
    $tCols = @(
        @{ idx = 9; month = 2 },
        @{ idx = 13; month = 3 },
        @{ idx = 18; month = 4 },
        @{ idx = 23; month = 5 },
        @{ idx = 29; month = 6 },
        @{ idx = 35; month = 7 }
    )
    
    $results = @()
    $lastRow = $ws.UsedRange.Rows.Count
    "Processing up to $lastRow rows..." | Out-File $log -Append -Encoding UTF8
    
    $currSite = ""
    $currGroup = ""
    $currModel = ""
    $currRPM = ""
    $consecutiveEmpty = 0

    for ($r = 7; $r -le $lastRow; $r++) {
        $vSite = $ws.Cells.Item($r, 1).Value2
        $vGroup = $ws.Cells.Item($r, 2).Value2
        $vModel = $ws.Cells.Item($r, 3).Value2
        $vRPM = $ws.Cells.Item($r, 4).Value2

        # Check if row is empty
        if ($null -eq $vSite -and $null -eq $vGroup -and $null -eq $vModel -and $null -eq $vRPM) {
            $consecutiveEmpty++
            if ($consecutiveEmpty -gt 20) { break }
            continue
        }
        $consecutiveEmpty = 0

        # Update propagation state
        if ($null -ne $vSite -and "$vSite" -ne "") { $currSite = "$vSite".Trim() }
        if ($null -ne $vGroup -and "$vGroup" -ne "") { $currGroup = "$vGroup".Trim() }
        if ($null -ne $vModel -and "$vModel" -ne "") { $currModel = "$vModel".Trim() }
        if ($null -ne $vRPM -and "$vRPM" -ne "") { $currRPM = "$vRPM".Trim() }

        if ($currModel -ne "") {
            foreach ($col in $tCols) {
                $val = $ws.Cells.Item($r, $col.idx).Value2
                if ($null -ne $val -and [double]$val -gt 0) {
                    $qty = [math]::Floor([double]$val)
                    for ($q = 1; $q -le $qty; $q++) {
                        
                        # Add literal month value in JS via formatting if needed, but numeric is fine and will be parsed as integer by the dashboard
                        $monthText = "$($col.month)" + [char]0xC6D4 # "월"
                        
                        $results += [PSCustomObject]@{
                            Site    = $currSite
                            Group   = $currGroup
                            Model   = $currModel
                            RPM     = $currRPM
                            Month   = $monthText
                            Code    = ""
                            Product = ""
                        }
                    }
                }
            }
        }
    }
    
    # Export with UTF8 and BOM for perfect Excel CSV compatibility
    $csvPath = "$dir\_FinalList.csv"
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    "SUCCESS. Total Rows: $($results.Count)" | Out-File $log -Append -Encoding UTF8
}
catch {
    "ERROR: $_" | Out-File $log -Append -Encoding UTF8
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
}
