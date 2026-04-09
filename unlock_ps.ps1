# unlock_ps.ps1
$ErrorActionPreference = "Stop"
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open("C:\Users\i0215099\Desktop\MPS_UPDATE\일반비_MPS2603-1(생산배포용).xlsx", $null, $true, 5, "dnpc1234")
    $wb.SaveAs("C:\Users\i0215099\Desktop\MPS_UPDATE\temp_mps_unlocked.xlsx", 51)
    $wb.Close($false)
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
    Write-Host "UNLOCK_SUCCESS"
} catch {
    Write-Error $_.Exception.Message
}
