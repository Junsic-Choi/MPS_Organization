# audit_mps_ps.ps1
$ErrorActionPreference = "Stop"
try {
    $xl = New-Object -ComObject Excel.Application
    $wb = $xl.Workbooks.Open("C:\Users\i0215099\Desktop\MPS_UPDATE\일반비_MPS2603-1(생산배포용).xlsx", $null, $true, 5, "dnpc1234")
    $ws = $wb.Sheets.Item(4)
    $range = $ws.Range("A1:E2000")
    $vals = $range.Value2
    $wb.Close($false)
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
    
    $output = ""
    for ($r = 1; $r -le 2000; $r++) {
        $row = ""
        for ($c = 1; $c -le 5; $c++) {
            $row += $vals[$r, $c].ToString() + "|"
        }
        $output += "$r|$row`n"
    }
    $output | Out-File "C:\Users\i0215099\Desktop\MPS_UPDATE\mps_sheet_dump_verified.txt" -Encoding utf8
    Write-Host "SUCCESS_DUMPED"
} catch {
    Write-Error $_.Exception.Message
}
