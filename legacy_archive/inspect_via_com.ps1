# inspect_via_com.ps1
# 사전 조건: Excel에서 일반비_MPS2603-1(생산배포용).xlsx 파일이 열려 있어야 합니다.

$outFile = Join-Path $PSScriptRoot "com_inspect_result.txt"
$lines = [System.Collections.Generic.List[string]]::new()

try {
    $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    Write-Host "ERROR: Excel이 실행 중이지 않습니다. 파일을 Excel에서 먼저 열어주세요."
    exit 1
}

# 대상 워크북 찾기
$wb = $null
foreach ($w in $xl.Workbooks) {
    if ($w.Name -like "*MPS2603*") { $wb = $w; break }
}
if ($null -eq $wb) {
    $names = ($xl.Workbooks | ForEach-Object { $_.Name }) -join ", "
    Write-Host "ERROR: MPS2603 파일을 못찾았습니다. 열린 파일들: $names"
    exit 1
}
$lines.Add("워크북: " + $wb.Name)

# 시트 목록
$sheetNames = ($wb.Worksheets | ForEach-Object { $_.Name }) -join ", "
$lines.Add("시트 목록: " + $sheetNames)

# ── 생산배포용 탭
$prodWs = $null
foreach ($ws in $wb.Worksheets) { if ($ws.Name -like "*생산배포용*") { $prodWs = $ws; break } }
if ($null -eq $prodWs) { $lines.Add("ERROR: 생산배포용 탭 없음") }
else {
    $lines.Add("`n=== 생산배포용 탭: " + $prodWs.Name + " ===")
    # 헤더 행 확인 (1~4행)
    for ($r = 1; $r -le 5; $r++) {
        $row = "R${r}: "
        foreach ($c in @(1,2,3,4,5,8,9,10,11,13)) {
            $val = $prodWs.Cells($r, $c).Text
            $col = [char](64+$c)
            $row += "$col=[$val] "
        }
        $lines.Add($row)
    }
    # 데이터 시작 행 찾기 (A열에 생산처 데이터가 있는 첫 행)
    $dataStart = 0
    for ($r = 1; $r -le 20; $r++) {
        $aVal = $prodWs.Cells($r, 1).Text
        if ($aVal -ne "" -and $aVal -notlike "*생산처*" -and $aVal -notlike "*합계*") {
            $dataStart = $r; break
        }
    }
    $lines.Add("데이터 시작 행: $dataStart")

    # 데이터 샘플 5행
    if ($dataStart -gt 0) {
        $lines.Add("--- 데이터 샘플 (A,B,C,D,E,H,I,J,K,M) ---")
        for ($r = $dataStart; $r -le ($dataStart + 7); $r++) {
            $aVal = $prodWs.Cells($r, 1).Text
            if ($aVal -eq "") { continue }
            $row = "R${r}: "
            foreach ($c in @(1,2,3,4,5,8,9,10,11,13)) {
                $val = $prodWs.Cells($r, $c).Text
                $col = [char](64+$c)
                $row += "$col=[$val] "
            }
            $lines.Add($row)
        }
    }

    # 마지막 행
    $lastRow = $prodWs.Cells($prodWs.Rows.Count, 1).End(-4162).Row  # xlUp = -4162
    $lines.Add("마지막 행: $lastRow")

    # E,H,I,J,K,M 열 월별 헤더 확인 (헤더행 위 행들에서)
    $lines.Add("--- 월 헤더 확인 (헤더 구간) ---")
    for ($r = 1; $r -le ($dataStart - 1); $r++) {
        $row = "R${r}: "
        foreach ($c in @(5,8,9,10,11,13)) {
            $val = $prodWs.Cells($r, $c).Text
            $col = [char](64+$c)
            $row += "$col=[$val] "
        }
        $lines.Add($row)
    }
}

# ── MPS 탭
$mpsWs = $null
foreach ($ws in $wb.Worksheets) { if ($ws.Name -eq "MPS") { $mpsWs = $ws; break } }
if ($null -eq $mpsWs) { $lines.Add("ERROR: MPS 탭 없음") }
else {
    $lines.Add("`n=== MPS 탭 ===")
    # 처음 6행 확인
    for ($r = 1; $r -le 6; $r++) {
        $row = "R${r}: "
        foreach ($c in @(4,5,7,8,9,13,18,23,29,35,41)) {
            $val = $mpsWs.Cells($r, $c).Text
            $col = [char](64+$c)
            $row += "$col=[$val] "
        }
        $lines.Add($row)
    }
    # I4+M4+R4+W4+AC4+AI4 합계
    $total = 0
    foreach ($c in @(9,13,18,23,29,35)) {
        $v = [double]($mpsWs.Cells(4, $c).Value2)
        $total += $v
    }
    $lines.Add("I4+M4+R4+W4+AC4+AI4 합계: $total")
    $lastRowMps = $mpsWs.Cells($mpsWs.Rows.Count, 4).End(-4162).Row
    $lines.Add("MPS 마지막 데이터 행: $lastRowMps")
    # 5~12행 샘플
    $lines.Add("--- MPS 데이터 샘플 R5~R12 ---")
    for ($r = 5; $r -le 12; $r++) {
        $row = "R${r}: "
        foreach ($c in @(4,5,7,8,9,13,18,23,29,35)) {
            $val = $mpsWs.Cells($r, $c).Text
            $col = [char](64+$c)
            $row += "$col=[$val] "
        }
        $lines.Add($row)
    }
}

# ── Site 탭
$siteWs = $null
foreach ($ws in $wb.Worksheets) { if ($ws.Name -like "*Site*" -or $ws.Name -like "*site*") { $siteWs = $ws; break } }
if ($null -eq $siteWs) { $lines.Add("Site 탭 없음") }
else {
    $lines.Add("`n=== Site 탭: " + $siteWs.Name + " === (첫 20행)")
    for ($r = 1; $r -le 20; $r++) {
        $row = "R${r}: "
        for ($c = 1; $c -le 10; $c++) {
            $val = $siteWs.Cells($r, $c).Text
            if ($val -ne "") { $row += "[$val] " }
        }
        if ($row -ne "R${r}: ") { $lines.Add($row) }
    }
}

$lines | Out-File -FilePath $outFile -Encoding UTF8
Write-Host "완료. 결과: $outFile"
