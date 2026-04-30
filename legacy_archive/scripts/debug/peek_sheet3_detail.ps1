$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $dir = Get-Location
    $files = Get-ChildItem -Path "$dir" -Filter "*MPS*.xlsx"
    $path = $files[0].FullName
    $workbook = $excel.Workbooks.Open($path)
    $sh = $workbook.Sheets.Item(3) # 생산기종별
    Write-Output "--- Sheet: $($sh.Name) Peek (Rows 1-30) ---"
    for ($r = 1; $r -le 30; $r++) {
        $row = "Row " + $r + ": "
        for ($c = 1; $c -le 20; $c++) { 
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
