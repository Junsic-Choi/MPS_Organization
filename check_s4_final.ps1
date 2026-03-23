$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.DisplayAlerts = $false
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\data_working.xlsx"
$wb = $excel.Workbooks.Open($path, 0, $true)
$ws = $wb.Sheets.Item(4)
$res = "Sheet 4 (MPS) Row 3 Head:`n"
for ($c = 1; $c -le 50; $c++) {
    $v = "$($ws.Cells.Item(3, $c).Value2)"
    if ($v -ne "") { $res += "Col $c : [$v]`n" }
}
$res | Out-File -FilePath "$dir\sheet4_head_final.txt" -Encoding UTF8
$wb.Close($false)
$excel.Quit()
