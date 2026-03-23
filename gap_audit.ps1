$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\일반비_MPS2603-1(생산배포용).xlsx"
$wb = $excel.Workbooks.Open($path, 0, $true)

# 1. Audit Sales on Sheet 2
$ws2 = $wb.Sheets.Item(2)
$salesSum = 0
for ($c = 1; $c -le 50; $c++) {
    $v4 = "$($ws2.Cells.Item(4, $c).Value2)"
    if ($v4 -match "판매") {
        for ($r = 7; $r -le $ws2.UsedRange.Rows.Count; $r++) {
            $val = $ws2.Cells.Item($r, $c).Value2
            if ($null -ne $val -and [double]$val -gt 0) {
                $salesSum += [double]$val
            }
        }
    }
}

# 2. Check Sheet 4 Count
$ws4 = $wb.Sheets.Item(4)
$s4Rows = $ws4.UsedRange.Rows.Count

$res = "Sales Sum (Sheet 2): $salesSum`n"
$res += "Sheet 4 Row Count: $s4Rows`n"
$res | Out-File -FilePath "$dir\audit_final_gap.txt" -Encoding UTF8
$wb.Close($false)
$excel.Quit()
