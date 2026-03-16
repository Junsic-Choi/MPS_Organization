$ErrorActionPreference = "Stop"
$logFile = "run_debug_log.txt"
"--- Start ---" | Out-File $logFile -Encoding utf8

function Log-Message($msg, $color = "White") {
    Write-Host $msg -ForegroundColor $color
    $msg | Out-File $logFile -Append -Encoding utf8
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    # Run in the directory where the script is located
    $dir = $PSScriptRoot
    Log-Message "실행 경로: $dir"
    
    # 1. MPS와 (생산배포용)이 포함된 엑셀 파일을 찾습니다.
    $files = Get-ChildItem -Path $dir -Filter "*.xlsx" | Where-Object { 
        $_.Name -like "*MPS*" -and $_.Name -like "*(생산배포용)*"
    }

    if ($null -eq $files -or $files.Count -eq 0) {
        $files = Get-ChildItem -Path $dir -Filter "*.xlsx" | Where-Object { $_.Name -like "*MPS*" }
    }

    if ($files.Count -eq 0) { throw "MPS 엑셀 파일을 이 폴더에서 찾을 수 없습니다." }
    
    # Use the most recently modified file if there are multiple
    $files = $files | Sort-Object LastWriteTime -Descending
    $path = $files[0].FullName
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($path)
    
    Log-Message "[$($files[0].Name)] 파일을 처리중입니다..." "Green"
    
    # [중요] 파일 형식이 .xls인데 확장자가 .xlsx인 경우 오픈 시 멈춤(Hang) 현상이 발생할 수 있습니다.
    # 이를 방지하기 위해 임시 .xls 파일로 복사하여 엽니다.
    $tempPath = "$dir\temp_processing_file.xls"
    Copy-Item $path $tempPath -Force
    
    Log-Message "워크북을 여는 중..." "Gray"
    $workbook = $excel.Workbooks.Open($tempPath, 0, $true) # ReadOnly로 오픈
    
    # 2. "생산배포용" 시트 찾기
    $targetSheet = $null
    Log-Message "시트 목록 확인 중..." "Gray"
    foreach ($s in $workbook.Sheets) {
        Log-Message "- 시트명: $($s.Name)" "Gray"
        if ($s.Name -like "*생산배포용*") {
            $targetSheet = $s
            Log-Message ">> '$($s.Name)' 시트를 사용합니다." "Cyan"
            break
        }
    }
    
    if (-not $targetSheet) { 
        Log-Message "!! '생산배포용' 이름의 시트를 찾지 못해 2번째 시트를 선택합니다." "Yellow"
        $targetSheet = $workbook.Sheets.Item(2) 
    }
    
    Log-Message "데이터 범위를 확인 중..." "Gray"
    $usedRange = $targetSheet.UsedRange
    $maxRows = $usedRange.Rows.Count + $usedRange.Row
    Log-Message "총 예상 행 수: $maxRows" "Gray"
    
    # Name the output file based on the found Excel file dynamic name
    $csvPath = "$dir\${baseName}_FinalList.csv"
    
    [System.IO.File]::WriteAllBytes($csvPath, [byte[]](239, 187, 191))
    
    $headerLine = "`"Site`",`"Group`",`"Model`",`"RPM`",`"Month`",`"SerialNo`""
    Out-File -FilePath $csvPath -InputObject $headerLine -Append -Encoding UTF8
    
    $last_site = ""
    $last_group = ""
    $last_model = ""
    
    # Read the month columns dynamically based on row 7
    $monthCols = @()
    # Safely scan columns 8 to 20 to find numeric month headers like "3", "4", "5", etc.
    for ($c = 6; $c -le 20; $c++) {
        $headerText = $targetSheet.Cells.Item(7, $c).Text
        if ($null -ne $headerText -and $headerText -match "^(\d+)$") {
            $monthObj = @{ Col = $c; Name = $headerText }
            $monthCols += $monthObj
        }
    }
    
    # If couldn't find dynamic headers, fallback to previous manual map
    if ($monthCols.Count -eq 0) {
        $monthCols = @(
            @{ Col = 8; Name = "3" },
            @{ Col = 9; Name = "4" },
            @{ Col = 10; Name = "5" },
            @{ Col = 11; Name = "6" },
            @{ Col = 13; Name = "7" }
        )
    }
    
    Log-Message "데이터 전개(Unroll) 중... 잠시만 기다려주세요." "Cyan"
    
    for ($r = 7; $r -le $maxRows; $r++) {
        $site = ""
        $group = ""
        $model = ""
        $rpm = ""
        
        $c1 = $targetSheet.Cells.Item($r, 1).Text
        $c2 = $targetSheet.Cells.Item($r, 2).Text
        $c3 = $targetSheet.Cells.Item($r, 3).Text
        $c4 = $targetSheet.Cells.Item($r, 4).Text
        
        if ($null -ne $c1) { $site = $c1.Trim() }
        if ($null -ne $c2) { $group = $c2.Trim() }
        if ($null -ne $c3) { $model = $c3.Trim() }
        if ($null -ne $c4) { $rpm = $c4.Trim() }
        
        if ($site.Length -gt 0) { $last_site = $site } else { $site = $last_site }
        if ($group.Length -gt 0) { $last_group = $group } else { $group = $last_group }
        if ($model.Length -gt 0) { $last_model = $model } else { $model = $last_model }
        
        if ($model.Length -eq 0 -and $rpm.Length -eq 0) { continue }
        
        foreach ($m in $monthCols) {
            $qtyText = $targetSheet.Cells.Item($r, $m.Col).Text
            if ($null -eq $qtyText) { continue }
            $qtyText = $qtyText.Replace(",", "").Trim()
            
            if ($qtyText -match "^-?\d+$") { 
                $qty = 0
                if ([int]::TryParse($qtyText, [ref]$qty) -and $qty -gt 0) {
                    for ($i = 1; $i -le $qty; $i++) {
                        $monthText = $m.Name + "월"
                        $rowCsv = "`"$site`",`"$group`",`"$model`",`"$rpm`",`"$monthText`",`"$i`""
                        Out-File -FilePath $csvPath -InputObject $rowCsv -Append -Encoding UTF8
                    }
                }
            }
        }
    }
    
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    Log-Message "`n✅ 추출 성공!" "Green"
    Log-Message "저장된 경로: $csvPath" "White"
    Start-Sleep -Seconds 3
    
}
catch {
    $err = "❌ 오류 발생: " + $_.ToString()
    if ($_.Exception) { $err += "`nException: " + $_.Exception.Message }
    if ($_.ScriptStackTrace) { $err += "`nStack: " + $_.ScriptStackTrace }
    Log-Message $err "Red"
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
