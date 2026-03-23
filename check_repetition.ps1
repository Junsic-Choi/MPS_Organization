$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\data_working.xlsx"
$wb = $excel.Workbooks.Open($path, 0, $true)
$ws = $wb.Sheets.Item(2)
$res = "Col 6 (Sales) Row 10-30 Analysis:`n"
for ($r = 10; $r -le 30; $r++) {
    $vModel = "$($ws.Cells.Item($r, 3).Value2)"
    $v6 = "$($ws.Cells.Item($r, 6).Value2)"
    $res += "Row $r : Model=[$vModel] Col6=[$v6]`n"
}
$res | Out-File -FilePath "$dir\col6_repetition.txt" -Encoding UTF8
$wb.Close($false)
$excel.Quit()
