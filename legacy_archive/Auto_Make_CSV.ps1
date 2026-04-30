# Auto_Make_CSV.ps1
# This is a wrapper for 1.추출기_실행하기.bat to run the final extraction script

$dir = Get-Location
$script = "$dir\Final_Extract_4650.ps1"

if (Test-Path $script) {
    powershell -ExecutionPolicy Bypass -File $script
} else {
    Write-Host "Error: $script not found!" -ForegroundColor Red
    pause
}
