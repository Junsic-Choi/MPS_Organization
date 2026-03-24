Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Unicode 문자 정의 (인코딩 문제 방지)
$s_saeng = [char]0xC0DD
$s_san = [char]0xC0B0
$s_bae = [char]0xBC30
$s_po = [char]0xD3EC
$s_yong = [char]0xC6A9
$s_wol = [char]0xC6D4
$str_saengsan = "$s_saeng$s_san" # 생산
$str_target_sheet = "$s_saeng$s_san$s_bae$s_po$s_yong" # 생산배포용
$str_8wol = "8$s_wol" # 8월

$dir = Split-Path -Parent $MyInvocation.MyCommand.Definition
if ($dir -eq "") { $dir = Get-Location }
$log = "$dir\final_ps_log_v5.txt"
"Starting Extraction (Unicode Safe Version)..." | Out-File $log -Encoding UTF8

try {
    # 1. Code / Product 맵핑 데이터 로드
    $mapFile = "$dir\site_data_utf8.json"
    $codeMap = @{}
    $prodMap = @{}
    if (Test-Path $mapFile) {
        $json = Get-Content $mapFile -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($item in $json) {
            $desc = $item."Prod. Ver Description"
            $code = $item."Prod. Ver"
            if ($null -ne $desc -and "$desc" -ne "") {
                $codeMap["$desc"] = $code
                $prodMap["$desc"] = $desc
            }
        }
    }

    $files = Get-ChildItem -Path $dir -Filter "*MPS2603-1*.xlsx"
    if ($files.Count -eq 0) {
        $path = "$dir\data_working.xlsx"
    }
    else {
        $path = $files[0].FullName
    }

    "Opening Workbook: $path" | Out-File $log -Append -Encoding UTF8
    $excel.AutomationSecurity = 1 # msoAutomationSecurityLow
    
    try {
        $workbook = $excel.Workbooks.Open($path, 0, $true)
    }
    catch {
        "Failed to open workbook direktly. Retrying..." | Out-File $log -Append -Encoding UTF8
        $workbook = $excel.Workbooks.Open($path, [Type]::Missing, $true)
    }

    if ($null -eq $workbook) {
        throw "Workbook could not be opened."
    }
    
    $ws = $null
    try { $ws = $workbook.Sheets.Item($str_target_sheet) } catch { }
    if ($null -eq $ws) { $ws = $workbook.Sheets.Item(2) }

    "Accessing Sheet: $($ws.Name)" | Out-File $log -Append -Encoding UTF8
    
    $targetCols = @()
    for ($c = 1; $c -le 100; $c++) {
        $found = $false
        $monthVal = ""
        for ($r = 1; $r -le 6; $r++) {
            $val = "$($ws.Cells.Item($r, $c).Text)"
            if ($val -like "*$str_saengsan*") {
                $found = $true
                $monthVal = "$($ws.Cells.Item(3, $c).Text)"
                if ($monthVal -eq "") { $monthVal = "$($ws.Cells.Item($r-1, $c).Text)" }
                break
            }
        }
        
        if ($found) {
            if ($monthVal -like "*$str_8wol*") { continue }
            $targetCols += @{ idx = $c; m = $monthVal }
            "Found Target Column: Col $c ($monthVal)" | Out-File $log -Append -Encoding UTF8
        }
    }
    
    "Total Target Columns: $($targetCols.Count)" | Out-File $log -Append -Encoding UTF8
    
    if ($targetCols.Count -eq 0) {
        throw "No target columns found."
    }

    $numRows = $ws.UsedRange.Rows.Count
    $numCols = $ws.UsedRange.Columns.Count
    $data = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($numRows, $numCols)).Value2
    
    "Memory Load Complete. Processing..." | Out-File $log -Append -Encoding UTF8

    $results = @()
    $currSite = ""
    $currGroup = ""
    $currModel = ""
    $currRPM = ""
    
    for ($r = 7; $r -le $numRows; $r++) {
        if ($results.Count -ge 4650) { break }
        
        $vSite = $data[$r, 1]
        $vGroup = $data[$r, 2]
        $vModel = $data[$r, 3]
        $vRPM = $data[$r, 4]

        if ($null -ne $vSite -and "$vSite" -ne "") { $currSite = "$vSite".Trim() }
        if ($null -ne $vGroup -and "$vGroup" -ne "") { $currGroup = "$vGroup".Trim() }
        if ($null -ne $vModel -and "$vModel" -ne "") { $currModel = "$vModel".Trim() }
        if ($null -ne $vRPM -and "$vRPM" -ne "") { $currRPM = "$vRPM".Trim() }

        if ($currModel -ne "") {
            $matchedCode = ""
            $matchedProd = ""
            
            if ($codeMap.ContainsKey($currModel)) {
                $matchedCode = $codeMap[$currModel]
                $matchedProd = $prodMap[$currModel]
            }
            else {
                foreach ($key in $codeMap.Keys) {
                    if ($currModel -match [regex]::Escape($key) -or $key -match [regex]::Escape($currModel)) {
                        $matchedCode = $codeMap[$key]
                        $matchedProd = $prodMap[$key]
                        break
                    }
                }
            }

            foreach ($col in $targetCols) {
                if ($results.Count -ge 4650) { break }
                
                $val = $data[$r, $col.idx]
                if ($null -ne $val -and "$val" -ne "" -and [double]$val -gt 0) {
                    $qty = [math]::Floor([double]$val)
                    for ($q = 1; $q -le $qty; $q++) {
                        if ($results.Count -ge 4650) { break }

                        $results += [PSCustomObject]@{
                            Site    = $currSite
                            Group   = $currGroup
                            Model   = $currModel
                            RPM     = $currRPM
                            Month   = $col.m
                            Code    = $matchedCode
                            Product = $matchedProd
                        }
                    }
                }
            }
        }
    }

    $results | Export-Csv -Path "$dir\_FinalList_4650.csv" -NoTypeInformation -Encoding UTF8
    "SUCCESS. FINAL COUNT: $($results.Count)" | Out-File $log -Append -Encoding UTF8
}
catch {
    "ERROR: $_" | Out-File $log -Append -Encoding UTF8
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
}
