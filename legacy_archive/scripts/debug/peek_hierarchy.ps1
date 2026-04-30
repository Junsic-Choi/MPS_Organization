$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $dir = Get-Location
    $files = Get-ChildItem -Path "$dir" -Filter "*MPS*.xlsx"
    $path = $files[0].FullName
    $workbook = $excel.Workbooks.Open($path)
    $sh = $workbook.Sheets.Item(4) # MPS
    Write-Output "--- MPS Hierarchy Peek (Rows 800-830) ---"
    for ($r = 800; $r -le 830; $r++) {
        $row = "Row " + $r + ": "
        for ($c = 1; $c -le 40; $c++) { 
            $val = $sh.Cells.Item($r, $c).Text
            $row += "[" + $c + "]" + $val + "|" 
        }
        Write-Output $row
    }
    $workbook.Close($false)
}
finally {
    $excel.Quit()
}
