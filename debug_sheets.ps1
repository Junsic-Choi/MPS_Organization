$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$dir = Get-Location
$files = Get-ChildItem -Path $dir -Filter "*MPS2603-1*.xlsx"
if ($files.Count -eq 0) {
    Write-Host "No target Excel file found."
    $excel.Quit()
    exit
}
$path = $files[0].FullName
Write-Host "Opening: $path"
try {
    $wb = $excel.Workbooks.Open($path, 0, $true)
    foreach ($ws in $wb.Sheets) {
        Write-Host "Sheet Name: $($ws.Name)"
        $v4 = $ws.Cells.Item(4, 18).Text # Looking at column 18 which was 4월 in previous scripts
        Write-Host "  Row 4, Col 18: [$v4]"
    }
    $wb.Close($false)
}
catch {
    Write-Host "Error: $_"
}
$excel.Quit()
