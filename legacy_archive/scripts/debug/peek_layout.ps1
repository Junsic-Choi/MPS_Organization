$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\i0215099\Desktop\MPS_UPDATE"
$log = "$dir\layout_sample.txt"
"Sheet 2 Layout Sample (Cols 1-20, Rows 1-20)..." | Out-File $log -Encoding UTF8

try {
    $path = "$dir\data_working.xlsx"
    $workbook = $excel.Workbooks.Open($path)
    $ws = $workbook.Sheets.Item(2)
    
    for ($r = 1; $r -le 20; $r++) {
        $rowStr = "$r: "
        for ($c = 1; $c -le 20; $c++) {
            $v = "$($ws.Cells.Item($r, $c).Text)"
            $rowStr += "[$v] "
        }
        $rowStr | Out-File $log -Append -Encoding UTF8
    }
}
catch {
    "ERROR: $_" | Out-File $log -Append -Encoding UTF8
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
}
