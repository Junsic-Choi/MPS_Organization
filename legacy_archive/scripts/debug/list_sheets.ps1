$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $dir = Get-Location
    $files = Get-ChildItem -Path "$dir" -Filter "*MPS*.xlsx"
    if ($files.Count -eq 0) { throw "MPS File NOT found" }
    $path = $files[0].FullName
    $workbook = $excel.Workbooks.Open($path)
    $sh = $workbook.Sheets.Item(4)
    Write-Output "--- Sheet: $($sh.Name) ---"
    for ($r = 1; $r -le 10; $r++) {
        $rowStr = "Row " + $r + ": "
        for ($c = 1; $c -le 45; $c++) {
            $val = $sh.Cells.Item($r, $c).Text
            if ($null -eq $val) { $val = "" }
            $rowStr = $rowStr + "[" + $c + "]" + $val + "|"
        }
        Write-Output $rowStr
    }
    $workbook.Close($false)
}
catch {
    Write-Output ("Error: " + $_.Exception.Message)
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}
