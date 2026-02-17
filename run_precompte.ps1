chcp 65001 | Out-Null
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$csv = ".\Export_Planning_BetterStreet.csv"
$year = Read-Host "Quelle année veux-tu exporter (ex: 2025) ?"
$out = ".\Planning_Ouvriers_$year.xlsx"

python .\betterstreet_to_precompte_v5.py $csv $year $out
pause
