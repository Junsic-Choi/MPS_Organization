$ErrorActionPreference = "Stop"

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $dir = "C:\Users\i0215099\Desktop\MPS_UPDATE"
    $files = Get-ChildItem -Path $dir -Filter "*MPS2603-1*.xlsx"
    $path = $files[0].FullName
    
    $workbook = $excel.Workbooks.Open($path)
    $targetSheet = $workbook.Sheets.Item(2)
    
    $maxRows = $targetSheet.UsedRange.Rows.Count + $targetSheet.UsedRange.Row
    
    $csvPath = "$dir\MPS2603-1_FinalList.csv"
    
    [System.IO.File]::WriteAllBytes($csvPath, [byte[]](239, 187, 191))
    
    $headerLine = "`"Site`",`"Group`",`"Model`",`"RPM`",`"Month`",`"SerialNo`""
    Out-File -FilePath $csvPath -InputObject $headerLine -Append -Encoding UTF8
    
    $last_site = ""
    $last_group = ""
    $last_model = ""
    
    $monthCols = @(
        @{ Col = 8; Name = "3" },
        @{ Col = 9; Name = "4" },
        @{ Col = 10; Name = "5" },
        @{ Col = 11; Name = "6" },
        @{ Col = 13; Name = "7" }
    )
    
    for ($r = 7; $r -le $maxRows; $r++) {
        $site = ""
        $group = ""
        $model = ""
        $rpm = ""
        
        $c1 = $targetSheet.Cells.Item($r, 1).Text
        $c2 = $targetSheet.Cells.Item($r, 2).Text
        $c3 = $targetSheet.Cells.Item($r, 3).Text
        $c4 = $targetSheet.Cells.Item($r, 4).Text
        
        if ($null -ne $c1) { $site = $c1.Trim() }
        if ($null -ne $c2) { $group = $c2.Trim() }
        if ($null -ne $c3) { $model = $c3.Trim() }
        if ($null -ne $c4) { $rpm = $c4.Trim() }
        
        if ($site.Length -gt 0) { $last_site = $site } else { $site = $last_site }
        if ($group.Length -gt 0) { $last_group = $group } else { $group = $last_group }
        if ($model.Length -gt 0) { $last_model = $model } else { $model = $last_model }
        
        if ($model.Length -eq 0 -and $rpm.Length -eq 0) { continue }
        
        foreach ($m in $monthCols) {
            $qtyText = $targetSheet.Cells.Item($r, $m.Col).Text
            if ($null -eq $qtyText) { continue }
            $qtyText = $qtyText.Replace(",", "").Trim()
            
            if ($qtyText -match "^-?\d+$") { 
                $qty = 0
                if ([int]::TryParse($qtyText, [ref]$qty) -and $qty -gt 0) {
                    for ($i = 1; $i -le $qty; $i++) {
                        $monthText = $m.Name + "월"
                        $rowCsv = "`"$site`",`"$group`",`"$model`",`"$rpm`",`"$monthText`",`"$i`""
                        Out-File -FilePath $csvPath -InputObject $rowCsv -Append -Encoding UTF8
                    }
                }
            }
        }
    }
    
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
}
catch {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
