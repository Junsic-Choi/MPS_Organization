$ErrorActionPreference = "Stop"
$logFile = "run_debug_log_v3.txt"
"--- Start v3 ---" | Out-File $logFile -Encoding utf8

function Log-Message($msg, $color = "White") {
    Write-Host $msg -ForegroundColor $color
    $msg | Out-File $logFile -Append -Encoding utf8
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    $dir = $PSScriptRoot
    $files = Get-ChildItem -Path $dir -Filter "*.xlsx" | Where-Object { $_.Name -like "*MPS*" }
    
    if ($files.Count -eq 0) { exit }
    
    $oldPath = $files[0].FullName
    # 확장자를 .xls로 임시 변경 (서명이 XLS이므로)
    $tempPath = "$dir\temp_mps_file.xls"
    Log-Message "파일 복사 중: $tempPath"
    Copy-Item $oldPath $tempPath -Force
    
    Log-Message "파일 오픈 시도 중 (ReadOnly)..."
    try {
        # Open(Filename, UpdateLinks, ReadOnly, Format, Password, ...)
        $workbook = $excel.Workbooks.Open($tempPath, 0, $true)
        Log-Message "오픈 성공! 시트 수: $($workbook.Sheets.Count)" "Green"
        for ($i = 1; $i -le $workbook.Sheets.Count; $i++) {
            $s = $workbook.Sheets.Item($i)
            Log-Message "[$i] 시트: $($s.Name)"
        }
        
        $workbook.Close($false)
    }
    catch {
        Log-Message "오픈 실패: $_" "Red"
    }
    finally {
        $excel.Quit()
        if (Test-Path $tempPath) { Remove-Item $tempPath }
    }
}
catch {
    Log-Message "치명적 오류: $_" "Red"
}
