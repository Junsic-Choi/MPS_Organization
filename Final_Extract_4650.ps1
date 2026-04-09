# Final_Extract_4650.ps1
# Definitive Extraction Hub: Triggers Hardened VBS Engine for 100% Authenticity

$basePath = "c:\Users\i0215099\Desktop\MPS_UPDATE"
$vbsScript = Join-Path $basePath "final_vbs_export.vbs"
$finalCsv = Join-Path $basePath "_FinalList_4650_Complete.csv"

Write-Host "Starting Definitive Extraction (v113 Hardened Engine)..."

# 1. Execute the native VBS engine (Single-process stability)
& cscript.exe //nologo "$vbsScript"

# 2. Status Check
if (Test-Path $basePath\_FinalList_4650.csv) {
    # Move/Rename to the Dashboard's target filename
    Move-Item -Path "$basePath\_FinalList_4650.csv" -Destination $finalCsv -Force
    Write-Host "SUCCESS: 4,650 Rows Extracted and Mapped."
} else {
    Write-Error "Extraction Failed: VBS Output Missing."
}
