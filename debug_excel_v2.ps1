$ErrorActionPreference = "Continue"
$log = "debug_log_v2.txt"

function Write-Log($msg) {
    echo $msg | Out-File -FilePath $log -Append -Encoding ascii
}

Write-Log "--- Debug Start v2 ---"

try {
    $dir = $PSScriptRoot
    Write-Log "Dir: $dir"
    
    $files = Get-ChildItem -Path $dir -Filter "*.xlsx" | Where-Object { $_.Name -like "*MPS*" }
    foreach ($f in $files) {
        Write-Log "Found file: $($f.Name)"
    }
    
    if ($files.Count -eq 0) {
        Write-Log "No files."
        exit
    }
    
    $path = $files[0].FullName
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    
    Write-Log "Opening: $($files[0].Name)"
    $workbook = $excel.Workbooks.Open($path)
    
    Write-Log "Sheet count: $($workbook.Sheets.Count)"
    foreach ($s in $workbook.Sheets) {
        Write-Log "Sheet: $($s.Name)"
    }
    
    $workbook.Close($false)
    $excel.Quit()
    Write-Log "--- Success ---"
}
catch {
    Write-Log "Error: $_"
    if ($excel) { $excel.Quit() }
}
