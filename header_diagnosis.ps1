$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$dir = Get-Location
$path = "$dir\일반비_MPS2603-1(생산배포용).xlsx"
if (!(Test-Path $path)) {
    $path = "$dir\data_working.xlsx"
}
Write-Host "Opening: $path"
try {
    $wb = $excel.Workbooks.Open($path, 0, $true)
    $ws = $wb.Sheets.Item("생산배포용")
    if ($null -eq $ws) {
        $ws = $wb.Sheets.Item(2)
    }
    Write-Host "Active Sheet: $($ws.Name)"
    $out = ""
    for ($r = 1; $r -le 10; $r++) {
        $rowStr = "Row $r: "
        for ($c = 1; $c -le 50; $c++) {
            $val = "$($ws.Cells.Item($r, $c).Text)"
            if ($val -ne "") {
                $rowStr += "C$c:[$val] "
            }
        }
        $out += $rowStr + "`r`n"
    }
    $out | Out-File "header_diagnosis.txt" -Encoding UTF8
    $wb.Close($false)
}
catch {
    Write-Host "Error: $_"
}
$excel.Quit()
