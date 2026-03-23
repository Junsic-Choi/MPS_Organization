$excel = New-Object -ComObject Excel.Application -ErrorAction SilentlyContinue
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = Get-Location
$logPath = "$dir\debug_output.txt"
if (Test-Path $logPath) { Remove-Item $logPath }
Start-Transcript -Path $logPath -Force

try {
    Write-Output "Searching for Excel files..."
    $files = Get-ChildItem -Path "$dir" -Filter "*MPS*.xlsx" | Where-Object { $_.Name -notlike "~$*" }
    if ($files.Count -eq 0) { throw "Excel file not found" }
    
    $path = $files[0].FullName
    $tempPath = "$dir\temp_data_extraction.xlsx"
    Write-Output "Copying file to temp: $tempPath"
    Copy-Item $path $tempPath -Force
    
    Write-Output "Opening workbook (Read-Only): $tempPath"
    $workbook = $excel.Workbooks.Open($tempPath, 0, $true)
    $ws = $workbook.Sheets.Item(2)
    Write-Output "Sheet accessed: $($ws.Name)"
    
    $tCols = @(
        @{ i = 5; m = 2 },
        @{ i = 8; m = 3 },
        @{ i = 9; m = 4 },
        @{ i = 10; m = 5 },
        @{ i = 11; m = 6 },
        @{ i = 13; m = 7 }
    )
    
    $res = @()
    $suffix = [char]0xC6D4
    
    for ($r = 7; $r -le 5000; $r++) {
        $v = $ws.Cells.Item($r, 3).Value2
        if ($null -eq $v -and $r -gt 2500) { break }
        if ($null -eq $v) { continue }
        
        $mo = "$v".Trim()
        $si = "$($ws.Cells.Item($r, 1).Value2)"
        $gp = "$($ws.Cells.Item($r, 2).Value2)"
        $rp = "$($ws.Cells.Item($r, 4).Value2)"
        
        foreach ($c in $tCols) {
            $q = $ws.Cells.Item($r, $c.i).Value2
            if ($null -ne $q -and $q -gt 0) {
                $count = [int]$q
                for ($k = 1; $k -le $count; $k++) {
                    $res += [PSCustomObject]@{
                        Site    = $si
                        Group   = $gp
                        Model   = $mo
                        RPM     = $rp
                        Month   = "$($c.m)$suffix"
                        Code    = ""
                        Product = ""
                    }
                }
            }
        }
    }
    
    Write-Output "Extraction complete. Total rows: $($res.Count)"
    $res | Export-Csv -Path "$dir\_FinalList.csv" -NoTypeInformation -Encoding UTF8
}
catch {
    Write-Output "ERROR: $_"
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
    Stop-Transcript
    if (Test-Path $tempPath) { Remove-Item $tempPath -ErrorAction SilentlyContinue }
}
