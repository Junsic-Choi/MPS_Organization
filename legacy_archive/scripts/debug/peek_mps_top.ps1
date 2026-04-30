$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $dir = Get-Location
    $files = Get-ChildItem -Path "$dir" -Filter "*MPS*.xlsx"
    $path = $files[0].FullName
    $workbook = $excel.Workbooks.Open($path)
    $sh = $workbook.Sheets.Item(4) # MPS
    Write-Output "--- MPS Top Peek (Rows 1-100) ---"
    for ($r = 1; $r -le 100; $r++) {
        $row = "Row " + $r + ": "
        for ($c = 1; $c -le 45; $c++) { 
            $val = $sh.Cells.Item($r, $c).Text
            if ($null -ne $val -and $val.Trim() -ne "") {
                $row += "[" + $c + "]" + $val + "|" 
            }
        }
        if ($row.Trim().Length -gt 10) { Write-Output $row }
    }
    $workbook.Close($false)
}
finally {
    $excel.Quit()
}
