$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $dir = Get-Location
    $files = Get-ChildItem -Path "$dir" -Filter "*MPS*.xlsx"
    $path = $files[0].FullName
    $workbook = $excel.Workbooks.Open($path)
    $sh = $workbook.Sheets.Item(4) # MPS
    Write-Output "--- RPM Search in MPS sheet ---"
    $target = "6K"
    $found = $sh.UsedRange.Find($target)
    if ($null -ne $found) {
        Write-Output "Found '$target' at Row $($found.Row), Col $($found.Column)"
    }
    else {
        $target = "8K"
        $found = $sh.UsedRange.Find($target)
        if ($null -ne $found) {
            Write-Output "Found '$target' at Row $($found.Row), Col $($found.Column)"
        }
        else {
            Write-Output "RPM keys NOT found in MPS sheet."
        }
    }
    $workbook.Close($false)
}
finally {
    $excel.Quit()
}
