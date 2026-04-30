$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $dir = Get-Location
    $files = Get-ChildItem -Path "$dir" -Filter "*MPS*.xlsx"
    if ($files.Count -eq 0) { throw "MPS File NOT found" }
    $path = $files[0].FullName
    $workbook = $excel.Workbooks.Open($path)
    $sh = $workbook.Sheets.Item(2) # 생산배포용
    Write-Output "--- Sheet: $($sh.Name) ---"
    # Search for ML0486
    $found = $sh.UsedRange.Find("ML0486")
    if ($null -ne $found) {
        Write-Output "Found ML0486 at Row $($found.Row), Col $($found.Column)"
    }
    else {
        Write-Output "ML0486 NOT found in sheet 2"
    }
    # Also peek at headers of row 6 up to col 100
    $rowStr = "Row 6: "
    for ($c = 1; $c -le 100; $c++) { $rowStr += "[" + $c + "]" + $sh.Cells.Item(6, $c).Text + "|" }
    Write-Output $rowStr
    $workbook.Close($false)
}
finally {
    $excel.Quit()
}
