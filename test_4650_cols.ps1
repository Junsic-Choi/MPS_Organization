$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = Get-Location
$path = "$dir\data_working.xlsx"
if (!(Test-Path $path)) {
    Copy-Item "$dir\일반비_MPS2603-1(생산배포용).xlsx" $path -Force
}

$workbook = $excel.Workbooks.Open($path, 0, $true)
$ws = $workbook.Sheets.Item(2)

$labelProd = [char]0xC0DD + [char]0xC0B0 # "생산"
$labelSales = [char]0xD310 + [char]0xB9E4 # "판매"

$results = ""

for ($c = 1; $c -le 50; $c++) {
    $v4 = "$($ws.Cells.Item(4, $c).Value2)"
    $v3 = "$($ws.Cells.Item(3, $c).Value2)"
    if ($v4 -match $labelProd -or $v4 -match $labelSales) {
        $sum = 0
        $sumProp = 0
        $currModel = ""
        for ($r = 7; $r -le $ws.UsedRange.Rows.Count; $r++) {
            $modelCell = $ws.Cells.Item($r, 3).Value2
            if ($null -ne $modelCell -and "$modelCell" -ne "") {
                $currModel = "$modelCell".Trim()
            }
            
            $val = $ws.Cells.Item($r, $c).Value2
            if ($null -ne $val -and [double]$val -gt 0) {
                $sum += [double]$val
                if ($currModel -ne "") {
                    $sumProp += [double]$val
                }
            }
        }
        $results += "Col $c ($v3 $v4): Raw Sum = $sum, Propagated Sum = $sumProp`n"
    }
}

$results | Out-File "$dir\col_sums_diag.txt" -Encoding UTF8
$workbook.Close($false)
$excel.Quit()
