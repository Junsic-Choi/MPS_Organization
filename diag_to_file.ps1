$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$log = "c:\Users\i0215099\Desktop\MPS_UPDATE\diag_log.txt"
"Starting Diag" | Out-File $log
try {
    $dir = $PSScriptRoot
    $file = Get-ChildItem -Path $dir -Filter "*MPS*(생산배포용)*.xlsx" | Select-Object -First 1
    "Opening: $($file.Name)" | Out-File $log -Append
    $wb = $excel.Workbooks.Open($file.FullName, 0, $true)
    $sh = $null
    foreach ($s in $wb.Sheets) { if ($s.Name -eq '생산배포용') { $sh = $s; break } }
    if ($sh) {
        "Headers in Row 7:" | Out-File $log -Append
        for ($c = 1; $c -le 30; $c++) {
            $txt = $sh.Cells.Item(7, $c).Text
            "Col $c: [$txt]" | Out-File $log -Append
        }
    }
    $wb.Close($false)
}
catch {
    $_.Exception.Message | Out-File $log -Append
}
finally {
    $excel.Quit()
}
