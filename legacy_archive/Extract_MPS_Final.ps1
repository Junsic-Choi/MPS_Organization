$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = Get-Location
$log = "$dir\extraction_full_log.txt"
"Starting MPS Tab Extraction..." | Out-File $log -Encoding UTF8

try {
    $files = Get-ChildItem -Path $dir -Filter "*MPS2603-1*"
    $targetFile = $null
    foreach ($f in $files) {
        if ($f.Name -match "MPS2603" -and $f.Extension -eq ".xlsx") {
            $targetFile = $f
            break
        }
    }
    
    if ($null -eq $targetFile) {
        Add-Content $log "Target Excel file not found!"
        exit
    }
    
    $path = $targetFile.FullName
    Add-Content $log "Opening file: $path"
    
    $workbook = $excel.Workbooks.Open($path, 0, $true)
    
    $ws = $null
    foreach ($s in $workbook.Sheets) {
        if ($s.Name -match "MPS") {
            $ws = $s
            break
        }
    }

    if ($null -eq $ws) {
        Add-Content $log "MPS Tab not found!"
        exit
    }

    $wsName = $ws.Name
    Add-Content $log "Found Target Sheet: $wsName"
    
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
    Add-Content $log "Processing up to $lastRow rows..."
    
    $consecutiveEmpty = 0
    $currSite = ""
    $currGroup = ""
    $currModel = ""
    
    Add-Content $log "Starting Loop..."
    
    for ($r = 7; $r -le $lastRow; $r++) {
        if ($r % 100 -eq 0) { Add-Content $log "Processing row $r..." }
        try {
            $vSiteRaw = $ws.Cells.Item($r, 7).Value2
            $vGroupRaw = $ws.Cells.Item($r, 2).Value2
            $vModelRaw = $ws.Cells.Item($r, 3).Value2
            $vCodeRaw = $ws.Cells.Item($r, 4).Value2
            $vProductRaw = $ws.Cells.Item($r, 5).Value2

            $vCode = ""
            if ($null -ne $vCodeRaw) { $vCode = "$vCodeRaw".Trim() }
            $vProduct = ""
            if ($null -ne $vProductRaw) { $vProduct = "$vProductRaw".Trim() }
            $vModel = ""
            if ($null -ne $vModelRaw) { $vModel = "$vModelRaw".Trim() }

            if ($vCode -eq "" -and $vProduct -eq "" -and $vModel -eq "") {
                $consecutiveEmpty++
                if ($consecutiveEmpty -gt 20) { break }
                continue
            }
            $consecutiveEmpty = 0

            if ($null -ne $vSiteRaw -and "$vSiteRaw" -ne "") {
                $siteStr = "$vSiteRaw".Trim()
                if ($siteStr -eq "1840") { $currSite = "01. 남산" }
                elseif ($siteStr -eq "1842") { $currSite = "03. 창원" }
                elseif ($siteStr -eq "07" -or $siteStr -match "삼광") { $currSite = "07. 삼광" }
                else { $currSite = $siteStr }
            }
            
            if ($null -ne $vGroupRaw -and "$vGroupRaw" -ne "") { $currGroup = "$vGroupRaw".Trim() }
            if ($vModel -ne "") { $currModel = $vModel }

            if ($currModel -ne "" -or $vCode -ne "") {
                $dispModel = $currModel
                if ($vModel -eq "" -and $vCode -ne "") { $dispModel = $vCode }
                
                foreach ($col in $tCols) {
                    $val = $ws.Cells.Item($r, $col.idx).Value2
                    if ($null -ne $val -and ($val -as [double]) -gt 0) {
                        $qty = [math]::Floor([double]$val)
                        $m = $col.month
                        $monthText = "$m" + "월"
                        for ($q = 1; $q -le $qty; $q++) {
                            $obj = New-Object PSObject
                            $obj | Add-Member NoteProperty Site $currSite
                            $obj | Add-Member NoteProperty Group $currGroup
                            $obj | Add-Member NoteProperty Model $dispModel
                            $obj | Add-Member NoteProperty RPM $vCode
                            $obj | Add-Member NoteProperty Month $monthText
                            $obj | Add-Member NoteProperty Code $vCode
                            $obj | Add-Member NoteProperty Product $vProduct
                            $results += $obj
                        }
                    }
                }
            }
        } 
        catch {
            $errLoop = $_.Exception.Message
            Add-Content $log "Error at row $r : $errLoop"
        }
    }
    
    $resCount = $results.Count
    Add-Content $log "Finished loop. Total results: $resCount"
    
    $csvPath = "$dir\_FinalList.csv"
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Add-Content $log "SUCCESS. Total Rows: $resCount"
}
catch {
    $errMain = $_.ToString()
    Add-Content $log "ERROR: $errMain"
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
}
