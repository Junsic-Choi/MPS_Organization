$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\i0215099\Desktop\MPS_UPDATE"
$log = "$dir\missing_rows_analysis.txt"
"Missing Rows Analysis started..." | Out-File $log -Encoding UTF8

try {
    $path = "$dir\data_working.xlsx"
    $workbook = $excel.Workbooks.Open($path)
    $ws = $workbook.Sheets.Item(2)
    
    $targetCols = @(5, 8, 9, 10, 11, 13)
    
    for ($r = 7; $r -le 2000; $r++) {
        $m = "$($ws.Cells.Item($r, 3).Text)".Trim()
        
        $hasQty = $false
        foreach ($c in $targetCols) {
            $val = $ws.Cells.Item($r, $c).Value2
            if ($null -ne $val -and $val -gt 0) { $hasQty = $true; break }
        }
        
        if ($m -eq "" -and $hasQty) {
            "Row $r has qty but NO MODEL NAME. Site: $($ws.Cells.Item($r,1).Text)" | Out-File $log -Append -Encoding UTF8
        }
    }
}
catch {
    "ERROR: $_" | Out-File $log -Append -Encoding UTF8
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
}
