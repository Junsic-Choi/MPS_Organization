$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
try {
    $dir = $PSScriptRoot
    $path = "$dir\temp_processing_file.xls"
    if (-not (Test-Path $path)) { throw "File not found: $path" }
    Write-Output "Processing: $path"
    $workbook = $excel.Workbooks.Open($path, 0, $true)
    foreach ($sh in $workbook.Sheets) {
        Write-Output "--- Sheet: $($sh.Name) ---"
        for ($r = 1; $r -le 20; $r++) {
            $line = ""
            for ($c = 1; $c -le 15; $c++) {
                $v = $sh.Cells.Item($r, $c).Text
                if ($null -eq $v) { $v = "" }
                $line += $v + "|"
            }
            Write-Output $line
        }
    }
    $workbook.Close($false)
}
catch {
    Write-Output "Error: $_"
}
finally {
    $excel.Quit()
}
