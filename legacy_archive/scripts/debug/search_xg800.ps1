$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $dir = Get-Location
    $files = Get-ChildItem -Path "$dir" -Filter "*MPS*.xlsx"
    $path = $files[0].FullName
    $workbook = $excel.Workbooks.Open($path)
    Write-Output "--- Global Search for 'XG800' in ALL SHEETS ---"
    foreach ($sh in $workbook.Sheets) {
        $found = $sh.UsedRange.Find("XG800")
        if ($null -ne $found) {
            Write-Output "Found in Sheet '$($sh.Name)' at Row $($found.Row), Col $($found.Column)"
        }
    }
    $workbook.Close($false)
}
finally {
    $excel.Quit()
}
