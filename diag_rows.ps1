$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\data_working.xlsx"
$workbook = $excel.Workbooks.Open($path, 0, $true)
$ws = $workbook.Sheets.Item(2)
$lastRow = $ws.UsedRange.Rows.Count
$res = "LastRow: $lastRow`n"

# Check columns by letter if possible
function Get-CellValue($colLetter, $row) {
    return "$($ws.Range($colLetter + $row).Value2)"
}

$cols = @("I", "M", "R", "W", "AC", "AI")
foreach ($c in $cols) {
    $v3 = Get-CellValue $c 3
    $v4 = Get-CellValue $c 4
    $res += "Col $c : R3=[$v3] R4=[$v4]`n"
}

$res | Out-File -FilePath "$dir\row_count_diag.txt" -Encoding UTF8
$workbook.Close($false)
$excel.Quit()
