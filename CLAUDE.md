# CLAUDE.md

## Project Overview

This is a German-language real estate investment analysis tool that calculates the **DSCR (Debt Service Coverage Ratio)** for apartment listings and exports results to Excel. It targets 100%-financed properties in affordable German markets (East Germany, Ruhr region).

**Language note:** All variable names, comments, and docstrings in the source code are in German.

---

## Repository Structure

```
Adam613/
├── apartment_dscr_search.py   # Main script: data, DSCR logic, Excel export
├── requirements.txt           # Python dependencies
├── wohnungen_dscr_analyse.xlsx # Generated output (committed to repo)
└── .gitignore                 # Ignores __pycache__, venv/, .env
```

---

## Running the Script

```bash
# Install dependencies
pip install -r requirements.txt

# Run the analysis
python3 apartment_dscr_search.py
```

Output: `wohnungen_dscr_analyse.xlsx` in the working directory.

---

## Key Financing Parameters (apartment_dscr_search.py, lines 20–26)

| Parameter | Value | Description |
|-----------|-------|-------------|
| `ZINSSATZ` | 3.8% | Annual interest rate |
| `TILGUNG` | 1.0% | Initial amortization rate |
| `NICHT_UMLEGBARE_KOSTEN` | 20% | Non-billable overhead (% of rent) |
| `INSTANDHALTUNGSRUECKLAGE` | €10/m²/year | Maintenance reserve |
| `LEERSTAND_RISIKO` | 2% | Vacancy risk |
| `MIN_DSCR` | 1.2 | Minimum DSCR threshold for filtering |

To adjust financing assumptions, modify these constants at the top of the file.

---

## Core Functions

### `berechne_dscr(kaufpreis, kaltmiete_monat, flaeche_qm)` (line 94)
Calculates DSCR and related financial metrics for a property.

**Returns a dict with:**
- `dscr` — Debt Service Coverage Ratio (NOI / annual debt service)
- `noi` — Net Operating Income after deductions
- `jahrliche_schuldendienstleistung` — Annual debt service
- `brutto_mietrendite` — Gross rental yield
- `netto_mietrendite` — Net rental yield
- `kaufpreis_pro_qm` — Purchase price per m²
- `kaufpreis_jahresmiete_faktor` — Price-to-annual-rent multiplier
- `monatliche_rate` — Monthly mortgage payment

### `erstelle_excel(angebote_mit_dscr, dateiname)` (line 133)
Creates a formatted Excel workbook with:
- Title row with financing parameters and creation date
- 17 data columns with currency/percentage formatting
- Color-coded DSCR column (green = good/great)
- Alternating row colors, frozen header, auto-filter

### `main()` (line 281)
Orchestrates: filter by `MIN_DSCR` → sort by DSCR descending → take top 10 → print summary → generate Excel.

---

## Data

The script currently uses **simulated listing data** (`ANGEBOTE` list, lines 30–91): 15 hardcoded apartments across German cities. Each entry has:

```python
{
    "titel": str,          # Listing title
    "stadt": str,          # City
    "plz": str,            # Postal code
    "flaeche": float,      # Area in m²
    "kaufpreis": float,    # Purchase price in €
    "kaltmiete": float,    # Monthly cold rent in €
    "zimmer": float,       # Number of rooms
    "baujahr": int,        # Year built
    "etage": str,          # Floor
    "url": str             # Listing URL
}
```

To integrate live data, replace `ANGEBOTE` with data fetched via the `requests` library (already in `requirements.txt`).

---

## Development Conventions

- **Language:** German for all domain-specific names and comments; keep this consistent when adding new code
- **No tests exist** — when adding logic, consider adding `pytest` tests for `berechne_dscr()` with known inputs
- **No build system** — this is a single-file script; avoid over-engineering the structure unless the project grows significantly
- **Excel output is committed** — `wohnungen_dscr_analyse.xlsx` is tracked in git (the `.gitignore` was explicitly updated to allow this)

---

## Dependencies

```
openpyxl>=3.1.0   # Used for Excel generation
requests>=2.31.0  # Present but not yet used (reserved for future live data fetching)
```

---

## Branch & Git Workflow

- Development branch: `claude/add-claude-documentation-e5YTm`
- Push with: `git push -u origin <branch-name>`
- There is no CI/CD pipeline — all validation is manual

---

## Potential Improvements

1. Replace hardcoded `ANGEBOTE` with live ImmobilienScout24 API/scraping via `requests`
2. Add `pytest` unit tests for `berechne_dscr()`
3. Extract financing parameters to a config file or CLI arguments
4. Add error handling and input validation
5. Add a `README.md` with usage instructions and sample output screenshot
