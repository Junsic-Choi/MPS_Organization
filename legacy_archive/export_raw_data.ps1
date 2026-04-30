# export_raw_data.ps1
$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
try {
    $wb = $xl.Workbooks.Open("C:\Users\i0215099\Desktop\MPS_UPDATE\일반비_MPS2603-1(생산배포용).xlsx", $null, $true, 5, "dnpc1234")
    
    # 1. Export Sheet 2 (Production Summary)
    $ws2 = $wb.Sheets.Item(2)
    $ws2.SaveAs("C:\Users\i0215099\Desktop\MPS_UPDATE\temp_prod_raw.csv", 6) # 6 = CSV
    
    # 2. Export Sheet 4 (MPS Reference)
    $ws4 = $wb.Sheets.Item(4)
    $ws4.SaveAs("C:\Users\i0215099\Desktop\MPS_UPDATE\temp_mps_raw.csv", 6)
    
    $wb.Close($false)
    Write-Host "EXPORT_SUCCESS"
} catch {
    Write-Error $_.Exception.Message
} finally {
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
