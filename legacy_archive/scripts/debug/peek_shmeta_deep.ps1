$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $dir = Get-Location
    $files = Get-ChildItem -Path "$dir" -Filter "*MPS*.xlsx"
    $path = $files[0].FullName
    $workbook = $excel.Workbooks.Open($path)
    $sh = $workbook.Sheets.Item(2) # 생산배포용
    Write-Output "--- shMeta Peek (Rows 1-20, Cols 1-20) ---"
    for ($r = 1; $r -le 20; $r++) {
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
