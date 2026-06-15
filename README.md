# crucible_of_belonging-US_BLS_Data

Source data, extraction scripts, and generated tables and figures for the U.S. Bureau of Labor Statistics (BLS) engineering-employment analysis used in the dissertation *The Crucible of Belonging: How Gateway Courses Shape Motivation, Identity, and Persistence in Engineering* (South Dakota State University).

This repository is the reproducibility companion to the BLS data-analysis white paper (VerSteeg, 2026). The dissertation places undergraduate engineering education in its national labor-market context using two artifacts — a discipline-level table of current and projected engineering employment, and a figure tracing total engineering employment over time against successive BLS ten-year projections. Both are generated programmatically from BLS source workbooks. This repository holds the source data, the scripts, and their outputs so that anyone can refresh the figures against a newer BLS edition, audit a single cell, or reuse the approach for a different occupational group.

## Repository structure

```
.
├── README.md
├── LICENSE
├── .gitignore
├── scripts/
│   ├── bls.py                  # builds the projections table (dissertation Table 3.1)
│   └── bls_historical.py       # builds the historical-trend figure + supporting tables
├── data/
│   ├── occupation.xlsx         # current edition (2024–2034), read by bls.py
│   ├── occupation_2008.xlsx    # historical editions, read by bls_historical.py
│   ├── occupation_2010.xlsx
│   ├── ...
│   └── occupation_2024.xlsx
└── outputs/
    ├── BLS_table.csv  / BLS_table.tex
    ├── BLS_actuals.csv / BLS_actuals.tex
    ├── BLS_projections.csv / BLS_projections.tex
    └── BLS_trend.tex
```

**Note on paths.** Both scripts resolve paths relative to the repository root (the parent of `scripts/`): they read input workbooks from `data/` and write generated files to `outputs/`, creating `outputs/` automatically if it does not exist. You can run them from any working directory — e.g. `python scripts/bls.py` — and the locations resolve correctly.

## Data source

All employment figures derive from the BLS *Occupational Projections and Worker Characteristics* release, sheet `Table 1.2` ("Occupational projections, YYYY–YYYY"). Employment is reported in thousands in the source. The current edition projects 2024 employment to 2034; the historical figure draws on every edition from 2008 onward present in the working directory. The two earliest editions, distributed by BLS as legacy `.xls` files, are pre-converted to `.xlsx` and stored under the same `occupation_YYYY.xlsx` naming scheme.

BLS data are a U.S. Government work and are in the public domain. No human-subjects data are involved.

## Occupational filter and discipline naming

Engineering occupations are selected by Standard Occupational Classification (SOC) code: every detailed engineering occupation (`17-2****`), plus Sales Engineers (`41-9031`) and Architectural & Engineering Managers (`11-9041`, labeled "Engineering Managers" with a footnote noting the aggregation). Each occupation is relabeled to a short discipline vocabulary shared across the dissertation's external-data companions — for example, "Electrical Engineers" → "Electrical." `bls.py` maps on the occupational title (case-insensitive); `bls_historical.py` maps on the SOC code directly, which is more robust across editions whose title wording drifts.

## Requirements

- Python 3.8+
- `bls.py`: `pandas`, `openpyxl`
- `bls_historical.py`: `pandas`, `openpyxl`, `numpy`, `scipy`

```bash
python -m pip install --upgrade pandas openpyxl numpy scipy
```

## Usage

```bash
cd scripts
python bls.py              # -> BLS_table.csv, BLS_table.tex
python bls_historical.py   # -> BLS_actuals.*, BLS_projections.*, BLS_trend.tex
```

`bls.py` reads the current edition's `Table 1.2`, detects the base/projection years from the workbook, selects the engineering occupations, converts thousands to actual counts, computes ten-year growth, sorts disciplines by base-year employment, and appends two summary rows ("Above Disciplines" and "Total US Workforce").

`bls_historical.py` discovers every `occupation_YYYY.xlsx` beside it, handles three historical workbook layouts (Formats A/B/C), de-duplicates editions that share a base/projection year (keeping the later release), and emits the historical actuals table, the projected ten-year-growth table, and a self-contained `pgfplots` trend figure with OLS fits through the actual and projected series.

**Adding a new edition:** drop a new `occupation_YYYY.xlsx` into the directory and rerun. It is discovered, de-duplicated, and incorporated into every output automatically.

## Outputs

`bls.py` writes a semicolon-separated data file (`BLS_table.csv`) and a complete LaTeX table (`BLS_table.tex`) that the dissertation `\input`s directly. `bls_historical.py` writes paired CSV/LaTeX files for the historical actuals and projected growth, plus a standalone `BLS_trend.tex` figure. The outputs are committed so the chain from the published BLS workbook to the typeset artifact is fully visible.

## License

The extraction scripts, the derived tables and figures, and any prose original to this repository are released under the Creative Commons Attribution 4.0 International (CC BY 4.0) license. The underlying BLS employment figures are published by the U.S. Bureau of Labor Statistics and are in the public domain as a U.S. Government work; the CC BY grant extends only to the derived and original material, not to the public-domain source data. See [`LICENSE`](LICENSE).

## Citation

> VerSteeg, J. (2026). *Analysis of U.S. Bureau of Labor Statistics engineering employment data for The Crucible of Belonging: Source data, extraction scripts, and generated tables and figures* (White paper). South Dakota State University, Open PRAIRIE. https://doi.org/[DOI assigned at deposit]

**Author:** John VerSteeg — ORCID [0009-0000-1704-7890](https://orcid.org/0009-0000-1704-7890) — Department of Mechanical Engineering, South Dakota State University.
