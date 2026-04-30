$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\i0215099\Desktop\MPS_UPDATE"
$log = "$dir\quick_sum_res.txt"
"Quick Diagnosis started..." | Out-File $log -Encoding UTF8

try {
    $path = "$dir\data_working.xlsx"
    $workbook = $excel.Workbooks.Open($path)
    $ws = $workbook.Sheets.Item(2)
    
    for ($c = 5; $c -le 50; $c++) {
        $r3 = "$($ws.Cells.Item(3, $c).Text)"
        $r4 = "$($ws.Cells.Item(4, $c).Text)"
        
        $range = $ws.Range($ws.Cells.Item(7, $c), $ws.Cells.Item(2000, $c))
        $sum = $excel.WorksheetFunction.Sum($range)
        
        if ($sum -gt 0) {
            "Col $c : R3=[$r3] R4=[$r4] SUM=[$sum]" | Out-File $log -Append -Encoding UTF8
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
