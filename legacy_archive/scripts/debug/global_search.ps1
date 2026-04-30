$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
try {
    $dir = Get-Location
    $files = Get-ChildItem -Path "$dir" -Filter "*MPS*.xlsx"
    $path = $files[0].FullName
    $workbook = $excel.Workbooks.Open($path)
    $sh = $workbook.Sheets.Item(4) # MPS
    Write-Output "--- Global Search in MPS sheet ---"
    
    # Search for "HM1000"
    $searchKey = "HM1000"
    $firstMatch = $sh.UsedRange.Find($searchKey)
    if ($null -ne $firstMatch) {
        Write-Output "Found '$searchKey' at Row $($firstMatch.Row), Col $($firstMatch.Column)"
        # Peek around this column
        $col = $firstMatch.Column
        $valAbove = $sh.Cells.Item(5, $col).Text
        Write-Output "Header for this column (Row 5): $valAbove"
    }
    else {
        Write-Output "'$searchKey' NOT found in MPS sheet."
    }
    
    # Also search for "01. 남산"
    $searchKey2 = "남산"
    $firstMatch2 = $sh.UsedRange.Find($searchKey2)
    if ($null -ne $firstMatch2) {
        Write-Output "Found '$searchKey2' at Row $($firstMatch2.Row), Col $($firstMatch2.Column)"
    }
    
    $workbook.Close($false)
}
finally {
    $excel.Quit()
}
