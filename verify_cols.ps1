$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\일반비_MPS2603-1(생산배포용).xlsx"
$temp = "$dir\temp_col_verify.xlsx"
Copy-Item $path $temp -Force
$wb = $excel.Workbooks.Open($temp, 0, $true)
$ws = $wb.Sheets.Item(2)
$res = "Verifying I, M, R, W, AC, AI on Sheet 2:`n"

$cols = @("I", "M", "R", "W", "AC", "AI")
foreach ($c in $cols) {
    $v3 = "$($ws.Range($c + "3").Value2)"
    $v4 = "$($ws.Range($c + "4").Value2)"
    $res += "Col $c : R3=[$v3] R4=[$v4]`n"
}

$res | Out-File -FilePath "$dir\specific_col_verify.txt" -Encoding UTF8
$wb.Close($false)
$excel.Quit()
Remove-Item $temp -ErrorAction SilentlyContinue
