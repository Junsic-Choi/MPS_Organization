$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$dir = "c:\Users\i0215099\Desktop\MPS_UPDATE"
$log = "$dir\combination_search.txt"
"Starting Combination Search for 4650..." | Out-File $log -Encoding UTF8

try {
    $path = "$dir\data_working.xlsx"
    $workbook = $excel.Workbooks.Open($path)
    $ws = $workbook.Sheets.Item(2)
    
    $cols = @()
    for ($c = 5; $c -le 50; $c++) {
        $r3 = "$($ws.Cells.Item(3, $c).Text)"
        $r4 = "$($ws.Cells.Item(4, $c).Text)"
        $range = $ws.Range($ws.Cells.Item(7, $c), $ws.Cells.Item(2000, $c))
        $sum = $excel.WorksheetFunction.Sum($range)
        if ($sum -gt 0) {
            $cols += @{ idx = $c; sum = $sum; name = "$r3 $r4" }
        }
    }
    
    # Simple backtracking or recursive search for combination
    function Find-Combination($target, $currentSum, $startIndex, $currentList) {
        if ($currentSum -eq $target) {
            return $currentList
        }
        if ($currentSum -gt $target -or $startIndex -ge $cols.Count) {
            return $null
        }
        
        # Include current
        $res = Find-Combination $target ($currentSum + $cols[$startIndex].sum) ($startIndex + 1) ($currentList + $cols[$startIndex])
        if ($res) { return $res }
        
        # Exclude current
        $res = Find-Combination $target $currentSum ($startIndex + 1) $currentList
        if ($res) { return $res }
        
        return $null
    }
    
    $target = 4650
    "Target: $target" | Out-File $log -Append -Encoding UTF8
    
    $result = Find-Combination $target 0 0 @()
    
    if ($result) {
        "FOUND COMBINATION:" | Out-File $log -Append -Encoding UTF8
        foreach ($item in $result) {
            "Col $($item.idx) ($($item.name)) : Sum $($item.sum)" | Out-File $log -Append -Encoding UTF8
        }
    }
    else {
        "No exact combination found for $target" | Out-File $log -Append -Encoding UTF8
    }
}
catch {
    "ERROR: $_" | Out-File $log -Append -Encoding UTF8
}
finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    $excel.Quit()
}
