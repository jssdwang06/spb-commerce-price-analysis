# St. Petersburg Commercial Real Estate — Price Factor Analysis

Data collection, cleaning and price-factor analysis for commercial property listings in Saint Petersburg (sale and rent).

## Repository layout

```
.
├── data/                 # Main Excel tables used for analysis
│   ├── Купить.xlsx       # Sale, 196 objects
│   └── Снять.xlsx        # Rent, 325 objects
├── figure/               # Analysis charts (PNG)
│   ├── 00_dashboard_overview.png
│   ├── 01_feature_importance.png
│   ├── …
│   └── 09_parking_walls.png
├── output/               # Deliverables
│   ├── Отчет_анализ_цен_коммерческой_недвижимости.pdf
│   ├── Отчет_анализ_цен_коммерческой_недвижимости.docx
│   └── analysis_summary.xlsx
├── raw/                  # Cached listing/detail HTML (for re-export only)
│   ├── arenda_325/html_cache/      # Rent: 11 list pages + 325 details
│   └── prodazha_196/html_cache/    # Sale: 7 list pages + 196 details
├── scripts/              # Scraper and analysis scripts
└── README.md
```

Large HTML under `raw/**/html_cache/` is ignored by Git (see `.gitignore`). Analysis runs from `data/*.xlsx`.

## HTML caches

| Path | Market | List pages | Detail pages | Excel |
|------|--------|------------|--------------|-------|
| `raw/prodazha_196/` | Sale | 7 | 196 | `data/Купить.xlsx` |
| `raw/arenda_325/` | Rent | 11 | 325 | `data/Снять.xlsx` |

- **list** — listing pages (many objects per page)
- **detail** — one object page each

## Scripts

**`generate_price_factors_fast.py`**  
Cleans fields and estimates feature importance (sale + rent).  
Input: `data/Купить.xlsx`, `data/Снять.xlsx`. Output: under `figure/`.

**`build_russian_report_pdf.py`**  
Builds the Russian report Word file with embedded figures.  
Input: `figure/*.png`. Output: `output/*.docx` (export PDF via Word if needed).

**`etagi_commerce_scraper.py`**  
Downloads list/detail HTML into a cache directory.

**`export_arenda_325_selected.py`**  
Parses rent HTML → `raw/arenda_325/etagi_arenda_325_selected.xlsx`.

**`analyze_arenda_325.py`**  
Summary stats for the rent selected table.

**`extract_cached_etagi_to_excel.py`**  
Parses sale HTML → `data/Купить.xlsx`.

**`run_arenda_325.ps1`**  
Optional one-shot: scrape rent + export selected Excel.

Run all commands from the **repository root**.

## Commands

```powershell
# Feature analysis
python .\scripts\generate_price_factors_fast.py data
python .\scripts\generate_price_factors_fast.py model

# Russian report (DOCX)
python .\scripts\build_russian_report_pdf.py

# Re-export sale table from HTML
python .\scripts\extract_cached_etagi_to_excel.py

# Re-export rent table and summarize
python .\scripts\export_arenda_325_selected.py
python .\scripts\analyze_arenda_325.py

# Optional: scrape rent again (slow)
.\scripts\run_arenda_325.ps1
```

## Deliverables

- Report PDF: `output/Отчет_анализ_цен_коммерческой_недвижимости.pdf`
- Editable DOCX: `output/Отчет_анализ_цен_коммерческой_недвижимости.docx`
- Charts: `figure/00`–`09`
- Summary table: `output/analysis_summary.xlsx`
- Main data: `data/Купить.xlsx`, `data/Снять.xlsx`

## Method (short)

1. Parse prices, area, metro time/distance, floor, categories from listing text  
2. For plots: trim price / price-per-m² / area at 1st–99th percentiles  
3. Group comparisons use **medians**; groups with very few objects may be hidden  
4. Spearman correlations + Random Forest permutation importance on price per m²  
5. Area elasticity of total price via log–log fit  

Label **«объектов: N»** on charts = number of objects in that group (sums can be less than 196/325 after trimming or missing fields).
