$csv = Get-ChildItem -Path . -Filter "*FinalList.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($null -eq $csv) { Write-Host "CSV not found"; exit }

$data = Import-Csv $csv.FullName
$unmatched = $data | Where-Object { $_.ModelCode -eq '' }
Write-Host "--- Top Unmatched Models ---"
$unmatched | Group-Object Model | Select-Object Name, Count | Sort-Object Count -Descending | Select-Object -First 20 | Format-Table -AutoSize
