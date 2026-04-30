$ErrorActionPreference = "Continue"
$log = "debug_log.txt"
"--- Debug Start ---" | Out-File $log

try {
    $dir = $PSScriptRoot
    "Current Dir: $dir" | Out-File $log -Append
    
    $files = Get-ChildItem -Path $dir -Filter "*.xlsx" | Where-Object { 
        $_.Name -like "*MPS*" -and $_.Name -like "*(생산배포용)*"
    }
    
    if ($null -eq $files -or $files.Count -eq 0) {
        "No matching files found." | Out-File $log -Append
        exit
    }
    
    $path = $files[0].FullName
    "Target File: $path" | Out-File $log -Append

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    
    "Opening workbook..." | Out-File $log -Append
    $workbook = $excel.Workbooks.Open($path)
    
    "Sheets found:" | Out-File $log -Append
    foreach ($s in $workbook.Sheets) {
        "- $($s.Name)" | Out-File $log -Append
    }
    
    $workbook.Close($false)
    $excel.Quit()
    "--- Debug Success ---" | Out-File $log -Append
}
catch {
    "Error occurred: $_" | Out-File $log -Append
    if ($excel) { $excel.Quit() }
}
