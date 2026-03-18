$csv = Get-ChildItem -Path . -Filter "*FinalList.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($null -eq $csv) { Write-Host "CSV not found"; exit }

Write-Host "--- Verifying Results: $($csv.Name) ---"
$data = Import-Csv $csv.FullName
if ($null -eq $data -or $data.Count -eq 0) { Write-Host "CSV is empty"; exit }

$grouped = $data | Group-Object Model | Select-Object Name, 
@{N = 'MatchedCount'; E = { ($_.Group | Where-Object { $_.ModelCode -ne '' }).Count } }, 
Count | Select-Object -First 50

$grouped | Format-Table -AutoSize

$totalMatched = ($data | Where-Object { $_.ModelCode -ne '' }).Count
$totalRows = $data.Count
$rate = [Math]::Round(($totalMatched / $totalRows) * 100, 2)

Write-Host "`n--- Summary ---"
Write-Host "Total Rows: $totalRows"
Write-Host "Matched Rows: $totalMatched ($rate%)"
