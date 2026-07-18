# Run arenda scrape/export from project root
$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $ProjectRoot

python .\scripts\etagi_commerce_scraper.py `
  --list-path /commerce/arenda/ `
  --html-cache-dir .\raw\arenda_325\html_cache `
  --detail-delay-min 8 `
  --detail-delay-max 18 `

python .\scripts\export_arenda_325_selected.py `
  --cache-dir .\raw\arenda_325\html_cache `
  --output .\raw\arenda_325\etagi_arenda_325_selected.xlsx

Write-Host "Done. HTML cache: .\raw\arenda_325\html_cache"
Write-Host "Excel output: .\raw\arenda_325\etagi_arenda_325_selected.xlsx"
Write-Host ("List HTML files: " + (Get-ChildItem .\raw\arenda_325\html_cache\list -File -ErrorAction SilentlyContinue | Measure-Object).Count)
Write-Host ("Detail HTML files: " + (Get-ChildItem .\raw\arenda_325\html_cache\detail -File -ErrorAction SilentlyContinue | Measure-Object).Count)
