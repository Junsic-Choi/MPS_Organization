$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$dir = Get-Location
$path = "$dir\data_working.xlsx"
$wb = $excel.Workbooks.Open($path, 0, $true)

$res = ""
foreach ($s in $wb.Sheets) {
    $found = $s.UsedRange.Find("4650")
    if ($null -ne $found) {
        $res += "Found 4650 in Sheet: " + $s.Name + " at " + $found.Address + "`n"
    }
    else {
        $res += "Not found in Sheet: " + $s.Name + "`n"
    }
}

$res | Out-File -FilePath "$dir\excel_search_4650.txt" -Encoding UTF8
$wb.Close($false)
$excel.Quit()
