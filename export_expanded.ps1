$ErrorActionPreference = "Stop"

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $dir = "C:\Users\i0215099\Desktop\MPS_UPDATE"
    $files = Get-ChildItem -Path $dir -Filter "*MPS2603-1*.xlsx"
    if ($files.Count -eq 0) { throw "File not found" }
    $path = $files[0].FullName
    
    $workbook = $excel.Workbooks.Open($path)
    $targetSheet = $workbook.Sheets.Item(2)
    
    $range = $targetSheet.UsedRange
    $maxRows = $range.Rows.Count
    
    $csvPath = "$dir\생산배포용_리스트.csv"
    
    # Write UTF8 BOM manually
    [System.IO.File]::WriteAllBytes($csvPath, [byte[]](239, 187, 191))
    
    $header = "`"생산처`",`"기종분류`",`"기종`",`"RPM`",`"월`",`"순번`""
    Out-File -FilePath $csvPath -InputObject $header -Append -Encoding UTF8
    
    $last_site = ""
    $last_group = ""
    $last_model = ""
    
    $monthCols = @(
        @{ Col = 8; Name = "3월" },
        @{ Col = 9; Name = "4월" },
        @{ Col = 10; Name = "5월" },
        @{ Col = 11; Name = "6월" },
        @{ Col = 13; Name = "7월" }
    )
    
    for ($r = 8; $r -le $maxRows; $r++) {
        $site = ""
        $group = ""
        $model = ""
        $rpm = ""
        
        if ($null -ne $range.Item($r, 1)) { $site = $range.Item($r, 1).Text.Trim() }
        if ($null -ne $range.Item($r, 2)) { $group = $range.Item($r, 2).Text.Trim() }
        if ($null -ne $range.Item($r, 3)) { $model = $range.Item($r, 3).Text.Trim() }
        if ($null -ne $range.Item($r, 4)) { $rpm = $range.Item($r, 4).Text.Trim() }
        
        if ($site -ne "") { $last_site = $site } else { $site = $last_site }
        if ($group -ne "") { $last_group = $group } else { $group = $last_group }
        if ($model -ne "") { $last_model = $model } else { $model = $last_model }
        
        if ($site -match "총합계") { continue }
        
        foreach ($m in $monthCols) {
            $qtyText = $range.Item($r, $m.Col).Text
            if ($null -eq $qtyText) { continue }
            $qtyText = $qtyText.Replace(",", "").Trim()
            
            if ($qtyText -match "^\d+$") {
                $qty = [int]$qtyText
                if ($qty -gt 0) {
                    for ($i = 1; $i -le $qty; $i++) {
                        $rowCsv = "`"$site`",`"$group`",`"$model`",`"$rpm`",`"$($m.Name)`",`"$i`""
                        Out-File -FilePath $csvPath -InputObject $rowCsv -Append -Encoding UTF8
                    }
                }
            }
        }
    }
    
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    Out-File -FilePath "$dir\com_success.txt" -InputObject "Success" -Encoding UTF8
}
catch {
    Out-File -FilePath "C:\Users\i0215099\Desktop\MPS_UPDATE\com_success.txt" -InputObject "Error: $_" -Encoding UTF8
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
