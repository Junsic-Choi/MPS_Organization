$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = Split-Path -Parent $MyInvocation.MyCommand.Definition
if ($dir -eq "") { $dir = Get-Location }
$log = "$dir\final_ps_log_v4.txt"
"Starting Extraction (Safe Encoding)..." | Out-File $log -Encoding UTF8

try {
    $path = "$dir\data_working.xlsx"
    $workbook = $excel.Workbooks.Open($path)
    $ws = $workbook.Sheets.Item(2)
    
    $targetCols = @(
        @{ idx = 5; m = "2" },
        @{ idx = 8; m = "3" },
        @{ idx = 9; m = "4" },
        @{ idx = 10; m = "5" },
        @{ idx = 11; m = "6" },
        @{ idx = 13; m = "7" }
    )
    
    $results = @()
    
    for ($r = 7; $r -le 2000; $r++) {
        $cellVal = $ws.Cells.Item($r, 3).Value2
        if ($null -eq $cellVal -and $r -gt 1500) { break }
        if ($null -eq $cellVal) { continue }
        $model = "$cellVal".Trim()
        
        $site = "$($ws.Cells.Item($r, 1).Value2)"
        $group = "$($ws.Cells.Item($r, 2).Value2)"
        $rpm = "$($ws.Cells.Item($r, 4).Value2)"
        
        foreach ($col in $targetCols) {
            $val = $ws.Cells.Item($r, $col.idx).Value2
            if ($null -ne $val -and $val -gt 0) {
                $qty = [int]$val
                for ($q = 1; $q -le $qty; $q++) {
                    $results += [PSCustomObject]@{
                        Site    = $site
                        Group   = $group
                        Model   = $model
                        RPM     = $rpm
                        Month   = $col.m
                        Code    = ""
                        Product = ""
                    }
                }
            }
        }
    }
    
    $results | Export-Csv -Path "$dir\_FinalList_4650.csv" -NoTypeInformation -Encoding UTF8
    "SUCCESS. Count: $($results.Count)" | Out-File $log -Append -Encoding UTF8
}
catch {
    "ERROR: $_" | Out-File $log -Append -Encoding UTF8
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
}
