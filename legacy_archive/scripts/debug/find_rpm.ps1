$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $dir = Get-Location
    $files = Get-ChildItem -Path "$dir" -Filter "*MPS*.xlsx"
    if ($files.Count -eq 0) { throw "MPS File NOT found" }
    $path = $files[0].FullName
    $workbook = $excel.Workbooks.Open($path)
    $sh = $workbook.Sheets.Item(4) # MPS sheet
    Write-Output "--- Sheet: $($sh.Name) ---"
    for ($r = 1; $r -le 10; $r++) {
        for ($c = 1; $c -le 100; $c++) {
            $val = $sh.Cells.Item($r, $c).Text
            if ($val -match "RPM") {
                Write-Output "Found 'RPM' at Row $r, Col $c"
            }
        }
    }
    # Also just dump row 5 completely up to col 100
    $rowStr = "Row 5: "
    for ($c = 1; $c -le 100; $c++) {
        $val = $sh.Cells.Item(5, $c).Text
        $rowStr = $rowStr + "[" + $c + "]" + $val + "|"
    }
    Write-Output $rowStr
    $workbook.Close($false)
}
catch {
    Write-Output "Error: $($_.Exception.Message)"
}
finally {
    $excel.Quit()
}
