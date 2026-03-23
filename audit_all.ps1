$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\일반비_MPS2603-1(생산배포용).xlsx"
$wb = $excel.Workbooks.Open($path, 0, $true)

$res = ""
foreach ($s in $wb.Sheets) {
    $res += "--- Sheet: " + $s.Name + " ---`n"
    for ($r = 3; $r -le 4; $r++) {
        for ($c = 1; $c -le 50; $c++) {
            $v = "$($s.Cells.Item($r, $c).Value2)"
            if ($v -ne "") {
                $res += "R$r C$c : [$v]`n"
            }
        }
    }
    $res += "`n"
}

$res | Out-File -FilePath "$dir\all_sheets_audit.txt" -Encoding UTF8
$wb.Close($false)
$excel.Quit()
