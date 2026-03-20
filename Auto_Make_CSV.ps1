$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = Get-Location
$logPath = "$dir\debug_output.txt"
Start-Transcript -Path $logPath -Force

try {
    Write-Output "Searching for Excel files..."
    $files = Get-ChildItem -Path "$dir" -Filter "*MPS*.xlsx"
    if ($files.Count -eq 0) { throw "Excel file not found" }
    
    $path = $files[0].FullName
    Write-Output "Opening workbook: $path"
    $workbook = $excel.Workbooks.Open($path)
    
    # 1. Build Metadata Map from Sheet 2 (생산배포용)
    Write-Output "Building metadata map from Sheet 2..."
    $wsMeta = $workbook.Sheets.Item(2)
    $metaMap = @{}
    $lastRowMeta = $wsMeta.UsedRange.Rows.Count
    $lastSite = ""
    $lastGroup = ""
    
    for ($r = 7; $r -le $lastRowMeta; $r++) {
        $site = "$($wsMeta.Cells.Item($r, 1).Value2)"
        $group = "$($wsMeta.Cells.Item($r, 2).Value2)"
        $model = "$($wsMeta.Cells.Item($r, 3).Value2)".Trim()
        $rpm = "$($wsMeta.Cells.Item($r, 4).Value2)"
        
        if ($site -eq "") { $site = $lastSite } else { $lastSite = $site }
        if ($group -eq "") { $group = $lastGroup } else { $lastGroup = $group }
        
        if ($model -ne "") {
            $key = $model.ToUpper() -replace "LYNX ", ""
            if (-not $metaMap.ContainsKey($key)) {
                $metaMap[$key] = @{ Site = $site; Group = $group; Model = $model; RPM = $rpm }
            }
        }
    }
    Write-Output "Map built with $($metaMap.Count) unique models."

    # 2. Extract Data from Sheet 4 (MPS)
    Write-Output "Extracting data from Sheet 4..."
    $wsMps = $workbook.Sheets.Item(4)
    $targetCols = @(9, 13, 18, 23, 29, 35) # I, M, R, W, AC, AI
    $months = @()
    foreach ($col in $targetCols) {
        $h = "$($wsMps.Cells.Item(3, $col).Value2)"
        if ($h -match "(\d+)") { $months += "$($Matches[1])월" } else { $months += $h }
    }
    Write-Output "Target Months: $($months -join ', ')"

    $results = @()
    $lastCode = ""
    $lastProduct = ""
    $lastMeta = $null
    
    # Process up to 10000 rows, but break if we see consecutive empties
    $consecutiveEmpties = 0
    for ($r = 7; $r -le 10000; $r++) {
        $code = "$($wsMps.Cells.Item($r, 4).Value2)".Trim()
        $prod = "$($wsMps.Cells.Item($r, 5).Value2)".Trim()
        
        if ($code -eq "" -and $prod -eq "") {
            $consecutiveEmpties++
            if ($consecutiveEmpties -gt 10 -and $r -gt 1000) { break }
            continue
        }
        $consecutiveEmpties = 0
        
        if ($code -ne "") { $lastCode = $code }
        if ($prod -ne "") { 
            $lastProduct = $prod 
            $mpsKey = $prod.ToUpper().Split("-")[0]
            
            # Match
            $found = $null
            if ($metaMap.ContainsKey($mpsKey)) {
                $found = $metaMap[$mpsKey]
            }
            else {
                foreach ($mK in $metaMap.Keys) {
                    if ($mK -match [regex]::Escape($mpsKey) -or $mpsKey -match [regex]::Escape($mK)) {
                        $found = $metaMap[$mK]
                        break
                    }
                }
            }
            $lastMeta = $found
        }
        
        # Check quantities
        foreach ($i in 0..($months.Count - 1)) {
            $val = $wsMps.Cells.Item($r, $targetCols[$i]).Value2
            if ($null -eq $val) { continue }
            $qty = [int]$val
            if ($qty -gt 0) {
                for ($q = 1; $q -le $qty; $q++) {
                    $results += [PSCustomObject]@{
                        Site    = if ($lastMeta) { $lastMeta.Site } else { "" }
                        Group   = if ($lastMeta) { $lastMeta.Group } else { "" }
                        Model   = if ($lastMeta) { $lastMeta.Model } else { "" }
                        RPM     = if ($lastMeta) { $lastMeta.RPM } else { "" }
                        Month   = $months[$i]
                        Code    = $lastCode
                        Product = $lastProduct
                    }
                }
            }
        }
    }
    
    Write-Output "Extraction complete. Total rows: $($results.Count)"
    $results | Export-Csv -Path "$dir\_FinalList.csv" -NoTypeInformation -Encoding UTF8
}
catch {
    Write-Host "ERROR: $_"
}
finally {
    Stop-Transcript
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
}
