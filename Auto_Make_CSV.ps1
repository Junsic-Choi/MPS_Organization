$ErrorActionPreference = "Stop"
$logFile = "run_debug_log.txt"
"--- Start ---" | Out-File $logFile -Encoding utf8

function Write-Log($msg, $color = "White") {
    Write-Host $msg -ForegroundColor $color
    $msg | Out-File $logFile -Append -Encoding utf8
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    # Run in the directory where the script is located
    $dir = $PSScriptRoot
    Write-Log "실행 경로: $dir"
    
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
    
    Write-Log "[$($files[0].Name)] 파일을 처리중입니다..." "Green"
    
    # [중요] 파일 형식이 .xls인데 확장자가 .xlsx인 경우 오픈 시 멈춤(Hang) 현상이 발생할 수 있습니다.
    # 이를 방지하기 위해 임시 .xls 파일로 복사하여 엽니다.
    $tempPath = "$dir\temp_processing_file.xls"
    Copy-Item $path $tempPath -Force
    
    Write-Log "워크북을 여는 중..." "Gray"
    $workbook = $excel.Workbooks.Open($tempPath, 0, $true) # ReadOnly로 오픈
    
    # 2. "생산배포용" 시트 찾기
    $targetSheet = $null
    Write-Log "시트 목록 확인 중..." "Gray"
    foreach ($s in $workbook.Sheets) {
        Write-Log "- 시트명: $($s.Name)" "Gray"
        if ($s.Name -like "*생산배포용*") {
            $targetSheet = $s
            Write-Log ">> '$($s.Name)' 시트를 사용합니다." "Cyan"
            break
        }
    }
    
    if (-not $targetSheet) { 
        Write-Log "!! '생산배포용' 이름의 시트를 찾지 못해 2번째 시트를 선택합니다." "Yellow"
        $targetSheet = $workbook.Sheets.Item(2) 
    }
    
    Write-Log "데이터 범위를 확인 중..." "Gray"
    $usedRange = $targetSheet.UsedRange
    $maxRows = $usedRange.Rows.Count + $usedRange.Row
    Write-Log "총 예상 행 수: $maxRows" "Gray"
    
    # Name the output file based on the found Excel file dynamic name
    $csvPath = "$dir\${baseName}_FinalList.csv"
    
    [System.IO.File]::WriteAllBytes($csvPath, [byte[]](239, 187, 191))
    
    [System.IO.File]::WriteAllBytes($csvPath, [byte[]](239, 187, 191))
    
    # Removed initial header line to avoid duplicates
    
    $last_site = ""
    $last_group = ""
    $last_model = ""
    
    # Read the month columns dynamically based on row 7
    $monthCols = @()
    # Safely scan columns 6 to 20 to find numeric month headers like "3", "4", "5", etc.
    for ($c = 6; $c -le 20; $c++) {
        $headerText = $targetSheet.Cells.Item(7, $c).Text
        if ($null -ne $headerText -and $headerText -match "^(\d+)") {
            $monthNum = $Matches[1]
            $monthObj = [PSCustomObject]@{ Col = $c; Name = $monthNum }
            $monthCols += $monthObj
        }
    }
    
    # If couldn't find dynamic headers, fallback to previous manual map
    if ($monthCols.Count -eq 0) {
        $monthCols = @(
            [PSCustomObject]@{ Col = 8; Name = "3" },
            [PSCustomObject]@{ Col = 9; Name = "4" },
            [PSCustomObject]@{ Col = 10; Name = "5" },
            [PSCustomObject]@{ Col = 11; Name = "6" },
            [PSCustomObject]@{ Col = 13; Name = "7" }
        )
    }
    
    # 2.5 MPS 탭 데이터 로딩 (전체 범위를 한 번에 읽어 속도 최적화)
    Write-Log "MPS 탭 데이터 로딩 중 (메모리 최적화)..." "Gray"
    $mpsTab = $null
    foreach ($s in $workbook.Sheets) { if ($s.Name -eq "MPS") { $mpsTab = $s; break } }
    
    $mpsEntries = @()
    if ($null -ne $mpsTab) {
        $mpsRows = $mpsTab.UsedRange.Rows.Count + $mpsTab.UsedRange.Row
        # D열(4) ~ H열(8) 데이터 한 번에 가져오기
        $mpsRange = $mpsTab.Range("D6:H$mpsRows").Value2
        if ($null -ne $mpsRange) {
            $rowLimit = $mpsRange.GetUpperBound(0)
            for ($r = 1; $r -le $rowLimit; $r++) {
                $mCode = "$($mpsRange[$r, 1])"; # D
                $mProd = "$($mpsRange[$r, 2])"; # E
                $mSite = "$($mpsRange[$r, 4])"; # G
                $mVer = "$($mpsRange[$r, 5])"; # H
                
                if ($mCode -or $mProd) {
                    $mpsEntries += [PSCustomObject]@{
                        Code    = if ($mCode) { $mCode.Trim() } else { "" }
                        Product = if ($mProd) { $mProd.Trim() } else { "" }
                        Site    = if ($mSite) { $mSite.Trim() } else { "" }
                        Ver     = if ($mVer) { $mVer.Trim() } else { "" }
                    }
                }
            }
        }
        Write-Log "MPS 데이터 확보: $($mpsEntries.Count) 건" "Gray"
        if ($mpsEntries.Count -gt 0) {
            Write-Log "MPS 샘플: Code=$($mpsEntries[0].Code), Site=$($mpsEntries[0].Site)" "Gray"
        }
    }
    
    # 2.6 site.xlsx 마스터 데이터 로딩 (보조 매핑)
    $masterList = @()
    $siteXlsxPath = "$dir\site.xlsx"
    if (Test-Path $siteXlsxPath) {
        Write-Log "site.xlsx 마스터 로드 중..." "Gray"
        try {
            $siteWb = $excel.Workbooks.Open($siteXlsxPath, 0, $true)
            $siteSh = $siteWb.Sheets.Item(1)
            $siteRows = $siteSh.UsedRange.Rows.Count
            
            # Data starts from Row 3 (Row 2 is header: Plant, Prod. Ver, Prod. Ver Description)
            for ($r = 3; $r -le $siteRows; $r++) {
                $pPlant = "$($siteSh.Cells.Item($r, 3).Value2)".Trim() # Column 3: Plant
                $pCode = "$($siteSh.Cells.Item($r, 4).Value2)".Trim()  # Column 4: Prod. Ver (Code)
                $pDesc = "$($siteSh.Cells.Item($r, 5).Value2)".Trim()  # Column 5: Description
                
                if ($pPlant -and $pCode) {
                    $masterList += [PSCustomObject]@{ Plant = $pPlant; Code = $pCode; Desc = $pDesc }
                }
            }
            $siteWb.Close($false)
            Write-Log "site.xlsx 로드 완료: $($masterList.Count) 건" "Gray"
        }
        catch {
            Write-Log "!! site.xlsx 로드 실패: $_" "Yellow"
        }
    }

    Write-Log "데이터 전개 시작..." "Cyan"
    
    $headerLine = '"Site","Group","Model","RPM","Month","SerialNo","ModelCode","ProductName"'
    Out-File -FilePath $csvPath -InputObject $headerLine -Append -Encoding UTF8
    
    # 속도를 위해 생산배포용 데이터도 미리 메모리에 담기 (필요시)
    # 여기서는 기존 루프를 유지하되 매핑 로직을 최적화함
    
    # 루프 최적화: 모든 데이터를 메모리에 담아 처리
    $lastCol = 20 # Fallback max column
    if ($monthCols.Count -gt 0) {
        $lastCol = ($monthCols | Measure-Object -Property Col -Maximum).Maximum
    }
    
    for ($r = 7; $r -le $maxRows; $r++) {
        # 한 줄 전체 데이터를 한 번의 COM 호출로 가져옴
        $rowRange = $targetSheet.Range($targetSheet.Cells.Item($r, 1), $targetSheet.Cells.Item($r, $lastCol)).Value2
        if ($null -eq $rowRange) { continue }
        
        $site = ""; $group = ""; $model = ""; $rpm = ""
        
        $c1 = $rowRange[1, 1]; $c2 = $rowRange[1, 2]; $c3 = $rowRange[1, 3]; $c4 = $rowRange[1, 4]
        
        if ($null -ne $c1) { $site = "$c1".Trim() }
        if ($null -ne $c2) { $group = "$c2".Trim() }
        if ($null -ne $c3) { $model = "$c3".Trim() }
        if ($null -ne $c4) { $rpm = "$c4".Trim() }
        
        if ($site.Length -gt 0) { $last_site = $site } else { $site = $last_site }
        if ($group.Length -gt 0) { $last_group = $group } else { $group = $last_group }
        if ($model.Length -gt 0) { $last_model = $model } else { $model = $last_model }
        
        if ($model.Length -eq 0 -and $rpm.Length -eq 0) { continue }

        # --- 유사도 기반 매핑 (Similarity Heuristics) ---
        $resCode = ""; $resProd = ""
        
        # 이름 정규화 (공백 제거, 대문자 변환 등)
        $cleanModel = $model -replace '\s+', ''
        
        # 1. Site가 일치하는 MPS 항목 필터링
        $possible = $mpsEntries | Where-Object { $_.Site -eq $site }
        
        if ($possible.Count -gt 0) {
            # 1-A. Exact Match (Model/Product)
            $match = $possible | Where-Object { 
                $_.Product -eq $model -or $_.Code -eq $model -or
                ($_.Product -replace '\s+', '') -eq $cleanModel
            } | Select-Object -First 1
            
            # 1-B. Similarity Match (Contains / Reverse Contains)
            if ($null -eq $match) {
                $match = $possible | Where-Object { 
                    $_.Product -like "*$model*" -or $model -like "*$($_.Product)*" -or
                    ($_.Product -replace '\s+', '') -like "*$cleanModel*" -or $cleanModel -like "*$($_.Product -replace '\s+', '')*"
                } | Select-Object -First 1
            }
            
            # 1-C. site.xlsx 브릿지 (사이트 제약 없이 검색 시도)
            if ($null -eq $match -and $masterList.Count -gt 0) {
                # 먼저 사이트 일치하는 항목 찾기
                $bridge = $masterList | Where-Object { 
                    $_.Plant -eq $site -and (
                        $_.Desc -eq $model -or 
                        $_.Desc -like "*$model*" -or 
                        $model -like "*$($_.Desc)*" -or
                        ($_.Desc -replace '\s+', '') -like "*$cleanModel*"
                    )
                } | Select-Object -First 1
                
                # 사이트 일치 실패 시 전체에서 검색
                if ($null -eq $bridge) {
                    $bridge = $masterList | Where-Object { 
                        $_.Desc -eq $model -or 
                        $_.Desc -like "*$model*" -or 
                        $model -like "*$($_.Desc)*" -or
                        ($_.Desc -replace '\s+', '') -like "*$cleanModel*"
                    } | Select-Object -First 1
                }
                
                if ($null -ne $bridge) {
                    # 브릿지에서 찾은 Code로 MPS 탭 다시 검색 (역시 사이트 제약 없이 시도)
                    $match = $possible | Where-Object { $_.Code -eq $bridge.Code } | Select-Object -First 1
                    if ($null -eq $match) {
                        $match = $mpsEntries | Where-Object { $_.Code -eq $bridge.Code } | Select-Object -First 1
                    }
                }
            }

            if ($null -ne $match) {
                $resCode = $match.Code
                $resProd = $match.Product
            }
        }
        
        # --- 데이터 전개 기록 ---
        foreach ($m in $monthCols) {
            $qtyVal = $rowRange[1, $m.Col]
            if ($null -eq $qtyVal -or $qtyVal -le 0) { continue }
            
            $qty = [int]$qtyVal
            $monthText = "$($m.Name)월"
            for ($i = 1; $i -le $qty; $i++) {
                $rowCsv = '"{0}","{1}","{2}","{3}","{4}","{5}","{6}","{7}"' -f $site, $group, $model, $rpm, $monthText, $i, $resCode, $resProd
                Out-File -FilePath $csvPath -InputObject $rowCsv -Append -Encoding UTF8
            }
        }
        
        if ($r % 50 -eq 0) { Write-Log "진행 중: $r / $maxRows..." "Gray" }
    }
    
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    
    Write-Log "`n✅ 추출 성공!" "Green"
    Write-Log "저장된 경로: $csvPath" "White"
    Start-Sleep -Seconds 3
    exit 0
}
catch {
    $err = "❌ 오류 발생: " + $_.ToString()
    if ($_.Exception) { $err += "`nException: " + $_.Exception.Message }
    if ($_.ScriptStackTrace) { $err += "`nStack: " + $_.ScriptStackTrace }
    Write-Log $err "Red"
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    exit 1
}
