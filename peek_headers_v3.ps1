$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$dir = "c:\Users\i0215099\Desktop\MPS_UPDATE"
$files = Get-ChildItem -Path $dir -Filter "*.xlsx" | Where-Object { $_.Name -like "*MPS*" -and $_.Name -like "*(생산배포용)*" }
$path = $files[0].FullName
$workbook = $excel.Workbooks.Open($path, 0, $true)
$targetSheet = $null
foreach ($s in $workbook.Sheets) {
    if ($s.Name -like "*생산배포용*") {
        $targetSheet = $s
        break
    }
}
if ($targetSheet) {
    "Headers in Row 7:"
    for ($c = 1; $c -le 30; $c++) {
        $val = $targetSheet.Cells.Item(7, $c).Text
        "Col $c: $val"
    }
}
else {
    "Sheet not found"
}
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
