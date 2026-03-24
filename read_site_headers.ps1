$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\site.xlsx"
$workbook = $excel.Workbooks.Open($path, 0, $true)
$ws = $workbook.Sheets.Item(1)

$res = ""
for ($c = 1; $c -le 20; $c++) {
    $val = $ws.Cells.Item(1, $c).Text
    if ($val -ne "") {
        $res += "Col $c : $val`n"
    }
}
$res | Out-File "$dir\site_headers.txt" -Encoding UTF8
$workbook.Close($false)
$excel.Quit()
