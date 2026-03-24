Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$excel.DisplayAlerts = $false

$s_saeng = [char]0xC0DD; $s_san = [char]0xC0B0; $s_bae = [char]0xBC30; $s_po = [char]0xD3EC; $s_yong = [char]0xC6A9; $s_wol = [char]0xC6D4
$str_saengsan = "$s_saeng$s_san"
$str_target_sheet = "$s_saeng$s_san$s_bae$s_po$s_yong"
$str_8wol = "8$s_wol"

$dir = Split-Path -Parent $MyInvocation.MyCommand.Definition
if ($dir -eq "") { $dir = Get-Location }
$log = "$dir\final_ps_log_v6.txt"
"Starting Extraction (Sanitized Version V8 - Safe Rows)..." | Out-File $log -Encoding UTF8

try {
    $mapFile = "$dir\site_data_utf8.json"
    $codeMap = @{}; $prodMap = @{}
    if (Test-Path $mapFile) {
        $json = Get-Content $mapFile -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($item in $json) {
            $desc = $item."Prod. Ver Description"; $code = $item."Prod. Ver"
            if ($null -ne $desc -and "$desc" -ne "") { $codeMap["$desc"] = $code; $prodMap["$desc"] = $desc }
        }
    }

    $files = Get-ChildItem -Path $dir -Filter "*MPS2603-1*.xlsx"
    if ($files.Count -eq 0) { $path = "$dir\data_working.xlsx" } else { $path = $files[0].FullName }
    
    $tmpPath = "$dir\data_tmp.xlsx"
    Copy-Item $path $tmpPath -Force
    "Opening Workbook: $tmpPath" | Out-File $log -Append -Encoding UTF8
    
    $excel.AutomationSecurity = 1
    $workbook = $excel.Workbooks.Open($tmpPath, 0, $true)
    
    $ws = $null
    try { $ws = $workbook.Sheets.Item($str_target_sheet) } catch { }
    if ($null -eq $ws) { $ws = $workbook.Sheets.Item(2) }
    "Accessing Sheet: $($ws.Name)" | Out-File $log -Append -Encoding UTF8
    
    $targetCols = @()
    for ($c = 1; $c -le 80; $c++) {
        $v3 = "$($ws.Cells.Item(3, $c).Text)"; $v4 = "$($ws.Cells.Item(4, $c).Text)"
        if (($v3 -like "*$s_wol*" -or $v4 -like "*$s_wol*") -and ($v4 -like "*$str_saengsan*" -or $v3 -like "*$str_saengsan*")) {
            if ($v3 -like "*$str_8wol*") { continue }
            $mName = if ($v3 -like "*$s_wol*") { $v3 } else { $v4 }
            $targetCols += @{ idx = $c; m = $mName }
            "Found Col: $c ($mName)" | Out-File $log -Append -Encoding UTF8
        }
    }
    
    if ($targetCols.Count -eq 0) { throw "No target columns" }

    # 안전하게 행 개수 찾기 (xlUp = -4162)
    $lastRow = $ws.Cells.Item(1048576, 3).End(-4162).Row
    if ($lastRow -lt 7) { $lastRow = 2000 } # 최소값
    if ($lastRow -gt 10000) { $lastRow = 10000 } # 최대값 캡
    
    "Data extent: 1 to $lastRow" | Out-File $log -Append -Encoding UTF8
    "Loading Data..." | Out-File $log -Append -Encoding UTF8
    $data = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($lastRow, 80)).Value2
    "Data Loaded. Processing..." | Out-File $log -Append -Encoding UTF8

    $results = @()
    $currSite = ""; $currGroup = ""; $currModel = ""; $currRPM = ""
    for ($r = 7; $r -le $lastRow; $r++) {
        if ($results.Count -ge 4650) { break }
        $vSite = $data[$r, 1]; $vGroup = $data[$r, 2]; $vModel = $data[$r, 3]; $vRPM = $data[$r, 4]
        if ($null -ne $vSite -and "$vSite" -ne "") { $currSite = "$vSite".Trim() }
        if ($null -ne $vGroup -and "$vGroup" -ne "") { $currGroup = "$vGroup".Trim() }
        if ($null -ne $vModel -and "$vModel" -ne "") { $currModel = "$vModel".Trim() }
        if ($null -ne $vRPM -and "$vRPM" -ne "") { $currRPM = "$vRPM".Trim() }

        if ($currModel -ne "") {
            $matchedCode = ""; $matchedProd = ""
            if ($codeMap.ContainsKey($currModel)) {
                $matchedCode = $codeMap[$currModel]; $matchedProd = $prodMap[$currModel]
            }
            else {
                foreach ($key in $codeMap.Keys) {
                    if ($currModel -match [regex]::Escape($key)) {
                        $matchedCode = $codeMap[$key]; $matchedProd = $prodMap[$key]; break
                    }
                }
            }

            foreach ($col in $targetCols) {
                if ($results.Count -ge 4650) { break }
                $val = $data[$r, $col.idx]
                $qty = 0
                if ($null -ne $val -and "$val" -ne "") {
                    if ([double]::TryParse("$val", [ref]$qty) -and $qty -gt 0) {
                        $qty = [math]::Floor($qty)
                        for ($q = 1; $q -le $qty; $q++) {
                            if ($results.Count -ge 4650) { break }
                            $results += [PSCustomObject]@{
                                Site = $currSite; Group = $currGroup; Model = $currModel; RPM = $currRPM
                                Month = $col.m; Code = $matchedCode; Product = $matchedProd
                            }
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
    if (Test-Path $tmpPath) { Remove-Item $tmpPath -Force }
}
