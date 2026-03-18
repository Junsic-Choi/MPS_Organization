$ErrorActionPreference = "Stop"
$logFile = "run_debug_v2.txt"
"--- Start ---" | Out-File $logFile -Encoding utf8

# Unicode constants
$strProd = [char]0xc0dd + [char]0xc0b0 # 생산
$strMonth = [char]0xc6d4 # 월
$strTotal = [char]0xd569 + [char]0xacc4 # 합계
$strDist = [char]0xbc30 + [char]0xd3ec # 배포

function Write-Log($msg, $color = "White") {
    if ($null -eq $msg) { $msg = "NULL MSG" }
    $ts = Get-Date -Format "HH:mm:ss"
    $line = "[$ts] $msg"
    Write-Host $line -ForegroundColor $color
    $line | Out-File $logFile -Append -Encoding utf8 -ErrorAction SilentlyContinue
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $dir = $PSScriptRoot
    if ($null -eq $dir) { $dir = Get-Location }
    Write-Log ("Current Dir: " + $dir)
    
    $files = Get-ChildItem -Path $dir -Filter "*MPS*.xls*" | Sort-Object LastWriteTime -Descending
    if ($files.Count -eq 0) { throw "MPS File NOT found" }
    $path = $files[0].FullName
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($path)
    Write-Log ("File: " + $files[0].Name) "Green"
    
    $tempPath = $dir + "\temp_processing_file.xls"
    Copy-Item $path $tempPath -Force
    
    Write-Log "Opening..." "Gray"
    $workbook = $excel.Workbooks.Open($tempPath, 0, $true)
    
    # ============================================================
    # 2. Find target sheet
    # ============================================================
    $targetSheet = $null
    foreach ($sh in $workbook.Sheets) {
        if ($sh.Name -match $strDist) { $targetSheet = $sh; break }
    }
    if ($null -eq $targetSheet) {
        if ($workbook.Sheets.Count -ge 2) { $targetSheet = $workbook.Sheets.Item(2) }
        else { $targetSheet = $workbook.Sheets.Item(1) }
    }
    Write-Log ("Selected Sheet: " + $targetSheet.Name) "Green"
    
    # ============================================================
    # 3. Header Detection (Row 3 = month labels, Row 4 = categories)
    # ============================================================
    $monthCols = @()
    $r3 = $targetSheet.Range("A3:AZ3").Value2
    $r4 = $targetSheet.Range("A4:AZ4").Value2
    
    for ($col = 5; $col -le 50; $col++) {
        $cat = $r4[1, $col]
        if ($null -eq $cat) { continue }
        if ($cat.ToString().Trim() -match $strProd) {
            for ($lc = $col; $lc -ge 1; $lc--) {
                $v = $r3[1, $lc]
                if ($null -ne $v) {
                    $vStr = $v.ToString().Trim()
                    if ($vStr -match "(\d+)\.?(\d+)?" + $strMonth -or $vStr -match "^(\d+)$") {
                        $mVal = if ($Matches[2]) { $Matches[2] } else { $Matches[1] }
                        $monthCols += [PSCustomObject]@{ Col = $col; Name = $mVal }
                        break
                    }
                }
            }
        }
    }

    if ($monthCols.Count -eq 0) { throw "Could NOT detect data columns" }
    Write-Log ("Month columns found: " + $monthCols.Count) "Gray"

    $csvPath = $dir + "\" + $baseName + "_FinalList.csv"
    
    # ============================================================
    # 4. Load Master Data from MPS tab
    #    Col4=ModelCode, Col5=Product
    #    Build a map: 기종명(from '생산배포용' Col3) -> List of (ModelCode, Product)
    #
    #    Strategy:
    #    A) From MPS tab: Build a flat list of (ModelCode, Product)
    #    B) The 기종명 in '생산배포용' is essentially a prefix of the Product name
    #       e.g. 기종="HM1000" matches Product containing "HM1000"
    #    C) For 기종s with multiple Models, store all of them
    # ============================================================
    
    # List of {ModelCode, Product} from MPS tab
    $mpsList = @()
    $mpsTab = $null
    foreach ($sh in $workbook.Sheets) {
        if ($sh.Name -eq "MPS") { $mpsTab = $sh; break }
    }
    
    if ($null -ne $mpsTab) {
        $mpsMax = $mpsTab.UsedRange.Rows.Count + $mpsTab.UsedRange.Row
        # Row 6+ : Col1=NR, Col2=PL, Col3=CH, Col4=Model, Col5=Product, Col7=Site
        $mpsRange = $mpsTab.Range("D6:H" + $mpsMax)
        $md = $mpsRange.Value2
        if ($md -is [System.Array]) {
            for ($i = 1; $i -le $md.GetUpperBound(0); $i++) {
                if ($null -ne $md[$i, 1]) {
                    $modelCode = $md[$i, 1].ToString().Trim()
                    $productName = if ($null -ne $md[$i, 2]) { $md[$i, 2].ToString().Trim() } else { "" }
                    if ($modelCode -ne "") {
                        $mpsList += [PSCustomObject]@{ C = $modelCode; P = $productName }
                    }
                }
            }
        }
    }
    
    # Build 기종 -> model/product lookup
    # Key insight: Product name (e.g. "XG800-F0TP-0-K10") doesn't directly match 기종 name (e.g. "HM1000")
    # BUT: Same ModelCode can have multiple Products (variants)
    # AND: Same 기종 can have multiple ModelCodes
    #
    # The hint is: "같은 기종은 같은 Model 값을 가짐" -> Model column in "생산배포용" has 기종 names!
    # In Row 6 of "생산배포용": Col3 = "Model" header, Data rows Col3 = 기종명 (HM1000, etc.)
    # In "생산기종별": Col1 = 기종군, Col2 = 기종명 (same format)
    # In "생산기종별" Row 6: "행 레이블 | Model | ..." -> here "Model" is actually the 기종명 column
    # 
    # Key: In "MPS" tab Col4, "Model" = SAP material code (ML0486)
    #      In "생산배포용/생산기종별" Col3, "Model" = 기종명 (HM1000)
    #
    # Link: Product (e.g., "XG800-F0TP-0-K10") -- the base model name is in the product code
    #       "XG800" is the 기종 base name; "ML0486" is the SAP model code
    #       We need to match by extracting the base from Product name
    
    # Build a reverse map: ModelCode -> List of Products (unique by ModelCode)
    $modCodeToProducts = @{}
    foreach ($entry in $mpsList) {
        if ($entry.C -ne "") {
            if (-not $modCodeToProducts.ContainsKey($entry.C)) {
                $modCodeToProducts[$entry.C] = @()
            }
            if ($entry.P -ne "" -and $modCodeToProducts[$entry.C] -notcontains $entry.P) {
                $modCodeToProducts[$entry.C] += $entry.P
            }
        }
    }
    
    # Function to find matching ModelCode(s)/Product(s) for a given 기종명
    # Matching strategy: extract base model name from Product and compare with 기종명
    # Product format: "HM1000-F0TP-0-K10" -> base = "HM1000" -> matches 기종 "HM1000"
    # Note: some products are "XG800-F0TP..." -> 기종 might be "XG800" or "XG800 II" etc.
    function Get-ModelMappings($kikongName, $mpsList) {
        if ($kikongName -eq "") { return @() }
        $results = @()
        $seen = @{}
        foreach ($entry in $mpsList) {
            # Try exact prefix match: Product starts with 기종명
            if ($entry.P -ne "" -and $entry.P -match ("^" + [Regex]::Escape($kikongName) + "[-\s/]")) {
                $key = $entry.C + "|" + $entry.P
                if (-not $seen.ContainsKey($key)) {
                    $results += [PSCustomObject]@{ C = $entry.C; P = $entry.P }
                    $seen[$key] = $true
                }
            }
        }
        # If no match found, try contains match
        if ($results.Count -eq 0) {
            foreach ($entry in $mpsList) {
                if ($entry.P -ne "" -and $entry.P.ToUpper().Contains($kikongName.ToUpper().Replace(" ", "").Replace(".", ""))) {
                    $key = $entry.C + "|" + $entry.P
                    if (-not $seen.ContainsKey($key)) {
                        $results += [PSCustomObject]@{ C = $entry.C; P = $entry.P }
                        $seen[$key] = $true
                    }
                }
            }
        }
        return $results
    }
    
    Write-Log ("MPS entries: " + $mpsList.Count) "Gray"
    
    # ============================================================
    # 5. Extract to CSV
    # ============================================================
    $headers = "Site,Group,Model,RPM,Month,SerialNo,ModelCode,ProductName"
    $headers | Out-File $csvPath -Encoding UTF8
    
    $lastS = ""; $lastG = ""; $lastM = ""; $lastR = ""
    $maxRows = $targetSheet.UsedRange.Rows.Count + $targetSheet.UsedRange.Row
    
    # Cache for 기종->mapping lookups
    $kikongCache = @{}
    
    Write-Log ("Processing rows up to " + $maxRows) "Gray"
    for ($r = 7; $r -le $maxRows; $r++) {
        try {
            $cells = $targetSheet.Range("A$r:AZ$r").Value2
            if ($null -eq $cells) { continue }
            
            $v1 = $cells[1, 1]; $v2 = $cells[1, 2]; $v3 = $cells[1, 3]; $v4 = $cells[1, 4]
            
            $site = if ($null -eq $v1 -or $v1.ToString().Trim() -eq "") { $lastS } else { $v1.ToString().Trim() }
            $group = if ($null -eq $v2 -or $v2.ToString().Trim() -eq "") { $lastG } else { $v2.ToString().Trim() }
            $model = if ($null -eq $v3 -or $v3.ToString().Trim() -eq "") { $lastM } else { $v3.ToString().Trim() }
            $rpm = if ($null -eq $v4) { $lastR } else { $v4.ToString().Trim() }

            $lastS = $site; $lastG = $group; $lastM = $model
            if ($v4 -ne $null -and $v4.ToString().Trim() -ne "") { $lastR = $rpm }
            
            if ($model -eq "" -and $rpm -eq "") { continue }
            if ($site -match $strTotal -or $site -match "Total") { continue }
            
            # Lookup ModelCode/Product for this 기종
            if (-not $kikongCache.ContainsKey($model)) {
                $kikongCache[$model] = Get-ModelMappings $model $mpsList
            }
            $mappings = $kikongCache[$model]
            
            foreach ($mc in $monthCols) {
                $qv = $cells[1, $mc.Col]
                $qty = 0
                if ($qv -is [double] -or $qv -is [int]) { $qty = [int]$qv }
                elseif ($qv -as [double]) { $qty = [int][double]$qv }
                
                if ($qty -gt 0) {
                    $mDisplay = $mc.Name + $strMonth
                    
                    if ($mappings.Count -eq 0) {
                        # No mapping found: output one row with empty ModelCode/Product
                        for ($idx = 1; $idx -le $qty; $idx++) {
                            $fields = @($site, $group, $model, $rpm, $mDisplay, $idx, "", "")
                            $line = ($fields | ForEach-Object { 
                                    $val = if ($_ -eq $null) { "" } else { $_.ToString() }
                                    '"{0}"' -f ($val -replace '"', '""') 
                                }) -join ","
                            $line | Out-File $csvPath -Append -Encoding UTF8
                        }
                    }
                    else {
                        # Each unit gets assigned a ModelCode/Product in round-robin if multiple
                        for ($idx = 1; $idx -le $qty; $idx++) {
                            $mapIdx = ($idx - 1) % $mappings.Count
                            $resC = $mappings[$mapIdx].C
                            $resP = $mappings[$mapIdx].P
                            $fields = @($site, $group, $model, $rpm, $mDisplay, $idx, $resC, $resP)
                            $line = ($fields | ForEach-Object { 
                                    $val = if ($_ -eq $null) { "" } else { $_.ToString() }
                                    '"{0}"' -f ($val -replace '"', '""') 
                                }) -join ","
                            $line | Out-File $csvPath -Append -Encoding UTF8
                        }
                    }
                }
            }
            if ($r % 100 -eq 0) { Write-Log "Processed Row $r..." "Gray" }
        }
        catch {
            Write-Log ("Error at Row $r : " + $_.Exception.Message) "Yellow"
        }
    }

    $workbook.Close($false)
    $excel.Quit()
    
    # Report mapping stats
    $totalRows = (Get-Content $csvPath).Count - 1
    $mappedRows = (Get-Content $csvPath | Select-String '","[A-Z][A-Z]\d+","' | Measure-Object).Count
    Write-Log ("Success: " + $csvPath) "Green"
    Write-Log ("Total rows: $totalRows, Rows with ModelCode: $mappedRows") "Green"
    exit 0
}
catch {
    Write-Log ("Critical Error: " + $_.Exception.Message) "Red"
    if ($excel) { $excel.Quit() }
    exit 1
}
