$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $dir = $PSScriptRoot
    $path = "$dir\temp_processing_file.xls"
    $workbook = $excel.Workbooks.Open($path, 0, $true)
    $sh = $workbook.Sheets.Item(2)
    $line = ""
    for ($c = 1; $c -le 30; $c++) {
        $line += $sh.Cells.Item(6, $c).Text + "|"
    }
    Write-Output $line
    $workbook.Close($false)
}
finally {
    $excel.Quit()
}
