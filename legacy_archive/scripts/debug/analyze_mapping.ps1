$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $dir = Get-Location
    $files = Get-ChildItem -Path "$dir" -Filter "*MPS*.xlsx"
    if ($files.Count -eq 0) { throw "MPS File NOT found" }
    $path = $files[0].FullName
    $workbook = $excel.Workbooks.Open($path)
    
    $sh2 = $workbook.Sheets.Item(2) # 생산배포용
    $sh4 = $workbook.Sheets.Item(4) # MPS
    
    Write-Output "--- Sheet: $($sh2.Name) (Index 2) ---"
    for ($r = 6; $r -le 15; $r++) {
        # Header and some data
        $row = ""
        for ($c = 1; $c -le 10; $c++) { $row += "[" + $c + "]" + $sh2.Cells.Item($r, $c).Text + "|" }
        Write-Output $row
    }
    
    Write-Output "--- Sheet: $($sh4.Name) (Index 4) ---"
    for ($r = 5; $r -le 15; $r++) {
        # Header and some data
        $row = ""
        for ($c = 1; $c -le 15; $c++) { $row += "[" + $c + "]" + $sh4.Cells.Item($r, $c).Text + "|" }
        Write-Output $row
    }
    
    $workbook.Close($false)
}
finally {
    $excel.Quit()
}
