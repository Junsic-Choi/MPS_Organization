$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\i0215099\Desktop\MPS_UPDATE"
$log = "$dir\duplicate_analysis.txt"
"Duplicate Analysis started..." | Out-File $log -Encoding UTF8

try {
    $path = "$dir\data_working.xlsx"
    $workbook = $excel.Workbooks.Open($path)
    $ws = $workbook.Sheets.Item(2)
    
    for ($r = 1; $r -le 200; $r++) {
        $m = "$($ws.Cells.Item($r, 3).Text)"
        if ($m -like "*MH0013*") {
            $s = "$($ws.Cells.Item($r, 1).Text)"
            $g = "$($ws.Cells.Item($r, 2).Text)"
            $rpm = "$($ws.Cells.Item($r, 4).Text)"
            $q2 = "$($ws.Cells.Item($r, 5).Text)"
            "Row $r | Site: $s | Group: $g | Model: $m | RPM: $rpm | 2Feb: $q2" | Out-File $log -Append -Encoding UTF8
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
