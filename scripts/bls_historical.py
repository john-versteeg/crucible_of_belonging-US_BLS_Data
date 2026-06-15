#!/usr/bin/env python3
# Copyright (c) 2026 John VerSteeg, South Dakota State University
# SPDX-License-Identifier: CC-BY-4.0
"""
bls_historical.py
------------------------------------------------------------
Author: John VerSteeg (South Dakota State University)

Purpose -- Generates three LaTeX documents + companion CSVs from all
  occupation_YYYY.xlsx files found in the data/ directory:

  1) outputs/BLS_actuals.tex / outputs/BLS_actuals.csv
       Historical actual employment by engineering discipline (in thousands),
       one column per BLS edition baseline year.

  2) outputs/BLS_projections.tex / outputs/BLS_projections.csv
       Projected growth % for each BLS edition (YYYY -> YYYY+10),
       one column per edition.  Editions where the projection target year
       is now in the past (actual data available) are marked with a dagger.

  3) outputs/BLS_trend.tex
       A self-contained pgfplots figure showing:
         - Actual total engineering employment over time (solid line)
         - Projected total engineering employment per edition (dashed line)
         - Linear regression fits through each dataset

AUTO-UPDATES: Drop a new occupation_YYYY.xlsx in the data/ directory
  and rerun -- it will be picked up automatically.

FORMATS HANDLED:
  Format A (2008, 2010, 2012):  col 0=title, col 1=code, col 2=base, col 3=proj
  Format B (2014-2019):         col 0=title, col 1=code, col 2=type, col 3=base, col 4=proj
  Format C (2020+):             named header row, "Employment, YYYY" columns

PATHS:
  Resolves paths relative to the repository root (the parent of scripts/):
  reads data/occupation_YYYY.xlsx and writes outputs/BLS_*.{tex,csv},
  creating outputs/ if needed. Run from anywhere, e.g.
  `python scripts/bls_historical.py`.

DEPENDENCIES:
  Python 3.8+, pandas, openpyxl, numpy, scipy
  Install via:
      python -m pip install --upgrade pandas openpyxl numpy scipy
"""

import glob
import os
import re

import numpy as np
import pandas as pd
from scipy import stats

# ── paths ─────────────────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT  = os.path.dirname(SCRIPT_DIR)
DATA_DIR   = os.path.join(REPO_ROOT, "data")
OUT_DIR    = os.path.join(REPO_ROOT, "outputs")

OUT_TEX1 = os.path.join(OUT_DIR, "BLS_actuals.tex")
OUT_CSV1 = os.path.join(OUT_DIR, "BLS_actuals.csv")
OUT_TEX2 = os.path.join(OUT_DIR, "BLS_projections.tex")
OUT_CSV2 = os.path.join(OUT_DIR, "BLS_projections.csv")
OUT_TEX3 = os.path.join(OUT_DIR, "BLS_trend.tex")

# ── discipline code map (SOC code → short label) ─────────────────────────────
CODE_MAP = {
    "11-9041": r"Eng.\ Managers\textsuperscript{a}",
    "17-2011": "Aerospace",
    "17-2021": "Agricultural",
    "17-2031": "Biomedical",
    "17-2041": "Chemical",
    "17-2051": "Civil",
    "17-2061": "Computer Hardware",
    "17-2071": "Electrical",
    "17-2072": "Electronics",
    "17-2081": "Environmental",
    "17-2111": r"Health \& Safety",
    "17-2112": "Industrial",
    "17-2121": "Marine",
    "17-2131": "Materials",
    "17-2141": "Mechanical",
    "17-2151": "Mining",
    "17-2161": "Nuclear",
    "17-2171": "Petroleum",
    "17-2199": "Engineers, All Other",
    "41-9031": "Sales Engineers",
}

# Display order: largest to smallest (will be sorted by mean employment later)
CODES_ORDERED = list(CODE_MAP.keys())

# ── helpers ───────────────────────────────────────────────────────────────────

def norm(s):
    return re.sub(r"\s+", " ", str(s).strip())



def _read_excel(path: str, **kwargs) -> pd.DataFrame:
    """Read an Excel .xlsx file using openpyxl (standard pandas dependency)."""
    return pd.read_excel(path, **kwargs)


def extract_years_from_title(title_str: str):
    """
    Pull base and projection year from a BLS table title.
    Handles both:
      "2008 and projected 2018"
      "2023–33"   (abbreviated projection year, last 2 digits only)
    """
    s = str(title_str)
    # Try abbreviated format: YYYY[–-]YY  e.g. "2023–33" or "2023-33"
    m_abbr = re.search(r"\b(20\d{2})[–\-]\s*(\d{2})\b", s)
    if m_abbr:
        base = int(m_abbr.group(1))
        suffix = int(m_abbr.group(2))
        proj = (base // 100) * 100 + suffix
        if proj <= base:
            proj += 100
        return base, proj
    # Standard: find all 4-digit years
    years = sorted(set(int(y) for y in re.findall(r"\b(20\d{2})\b", s)))
    if len(years) < 2:
        raise ValueError(f"Could not find two years in title: '{title_str}'")
    return years[0], years[-1]


def detect_format(raw: pd.DataFrame):
    """
    Detect which of the three column layouts this file uses.
    Returns 'A', 'B', or 'C'.
    """
    # Format C: clean header in row 1 with named Employment columns
    row1 = [str(x) for x in raw.iloc[1].tolist()]
    if any(re.match(r"Employment,\s*\d{4}", h) for h in row1):
        return "C"
    # Format A vs B: check whether a 'Summary' / 'Line item' type column exists at col 2
    # Find first row that looks like a real data row (code in col 1 matches XX-XXXX)
    for i in range(2, 20):
        row = raw.iloc[i].tolist()
        if re.match(r"\d{2}-\d{4}", str(row[1]).strip()):
            val = str(row[2]).strip()
            if val in ("Summary", "Line item", "Detailed"):
                return "B"
            else:
                return "A"
    return "A"


def parse_file(path: str):
    """
    Parse one occupation_YYYY file.

    Returns:
      base_year  (int)
      proj_year  (int)
      records    list of (code, base_emp_thousands, proj_emp_thousands)
    """
    raw = _read_excel(path, sheet_name="Table 1.2", header=None)
    title_str = str(raw.iloc[0, 0])
    base_year, proj_year = extract_years_from_title(title_str)
    fmt = detect_format(raw)

    if fmt == "C":
        # ── Format C (2020+): named header row ────────────────────────────
        header = raw.iloc[1].tolist()
        df = raw.iloc[2:].copy()
        df.columns = header

        col_code = next(
            c for c in df.columns
            if re.match(r"\d{4}\s+National Employment Matrix code", str(c))
        )
        col_base = next(
            c for c in df.columns
            if re.match(rf"Employment,\s*{base_year}", str(c))
        )
        col_proj = next(
            c for c in df.columns
            if re.match(rf"Employment,\s*{proj_year}", str(c))
        )

        codes     = df[col_code].astype(str).str.strip()
        base_vals = pd.to_numeric(df[col_base], errors="coerce")
        proj_vals = pd.to_numeric(df[col_proj], errors="coerce")

    else:
        # ── Formats A and B: raw column indices ───────────────────────────
        # Data starts at row 4; col 1 = code, col (2 or 3) = base, col (3 or 4) = proj
        df = raw.iloc[4:].copy()
        df.columns = range(df.shape[1])
        codes = df[1].astype(str).str.strip()

        if fmt == "A":
            base_vals = pd.to_numeric(df[2], errors="coerce")
            proj_vals = pd.to_numeric(df[3], errors="coerce")
        else:   # fmt == "B"
            base_vals = pd.to_numeric(df[3], errors="coerce")
            proj_vals = pd.to_numeric(df[4], errors="coerce")

    records = []
    for code in CODE_MAP:
        mask = codes == code
        if mask.any():
            b = float(base_vals[mask].iloc[0])
            p = float(proj_vals[mask].iloc[0])
            if not (np.isnan(b) or np.isnan(p)):
                records.append((code, b, p))

    return base_year, proj_year, records


# ── main ──────────────────────────────────────────────────────────────────────

def main():
    # --- discover files -------------------------------------------------------
    # Only .xlsx files are supported. Use the pre-converted occupation_2008.xlsx
    # and occupation_2012.xlsx provided in data/ (not the .xls originals).
    os.makedirs(OUT_DIR, exist_ok=True)
    found = sorted(
        f for f in glob.glob(os.path.join(DATA_DIR, "occupation_*.xlsx"))
        if re.search(r"occupation_\d{4}\.xlsx$", f, re.I)
    )
    if not found:
        raise FileNotFoundError(
            f"No occupation_YYYY.xlsx files found in {DATA_DIR}\n"
            "Download the BLS Occupational Projections workbooks and place them "
            "in the data/ directory using the occupation_YYYY.xlsx naming scheme."
        )

    # --- parse each file ------------------------------------------------------
    editions = []   # list of dicts: file_year, base_year, proj_year, records
    for fpath in found:
        m = re.search(r"occupation_(\d{4})\.xlsx$", fpath, re.I)
        file_year = int(m.group(1))

        base_year, proj_year, records = parse_file(fpath)
        editions.append({
            "file_year": file_year,
            "base_year": base_year,
            "proj_year": proj_year,
            "records":   records,
        })
        print(f"  Parsed {os.path.basename(fpath):30s}  {base_year} -> {proj_year}  "
              f"({len(records)} disciplines)")

    editions.sort(key=lambda e: (e["base_year"], e["file_year"]))

    # Deduplicate: when two files share the same (base_year, proj_year), keep
    # the one with the higher file_year (later / more authoritative release).
    seen = {}
    for ed in editions:
        key = (ed["base_year"], ed["proj_year"])
        if key not in seen or ed["file_year"] > seen[key]["file_year"]:
            seen[key] = ed
    editions = sorted(seen.values(), key=lambda e: e["base_year"])

    n_raw = len(found)
    n_kept = len(editions)
    if n_raw != n_kept:
        print(f"  (Deduplicated {n_raw} files -> {n_kept} unique editions)")

    # --- build per-discipline dataframes ---------------------------------------
    # actuals:     code → {year: employment_thousands}
    # projections: code → {proj_year: (growth_pct, base_year)}

    actual_years = sorted({e["base_year"] for e in editions})
    proj_years   = sorted({e["proj_year"] for e in editions})

    actual_map = {code: {} for code in CODE_MAP}   # code → {year: emp_k}
    proj_map   = {code: {} for code in CODE_MAP}   # code → {proj_year: (g, base_yr)}

    agg_actuals = {}    # base_year → total_eng_employment_thousands
    agg_projs   = {}    # proj_year → total_proj_eng_employment_thousands

    for ed in editions:
        base_yr = ed["base_year"]
        proj_yr = ed["proj_year"]
        total_base = 0.0
        total_proj = 0.0
        for code, b, p in ed["records"]:
            actual_map[code][base_yr] = b
            growth = (p - b) / b * 100 if b else np.nan
            proj_map[code][proj_yr] = (growth, base_yr)
            total_base += b
            total_proj += p
        agg_actuals[base_yr] = total_base
        agg_projs[proj_yr]   = total_proj

    # Determine discipline display order (descending mean actual employment)
    mean_emp = {}
    for code in CODE_MAP:
        vals = [v for v in actual_map[code].values() if not np.isnan(v)]
        mean_emp[code] = np.mean(vals) if vals else 0
    ordered_codes = sorted(CODE_MAP.keys(), key=lambda c: -mean_emp[c])

    # Which proj years now have actual data available?
    realized_proj_years = set(proj_years) & set(actual_years)

    # ── TABLE 1: ACTUALS ──────────────────────────────────────────────────────
    def fmt_k(v):
        """Format a value in thousands to 1 decimal place."""
        if np.isnan(v):
            return "---"
        return f"{v:.1f}"

    def fmt_g(v):
        if np.isnan(v):
            return "---"
        return f"{v:.1f}\\%"

    # LaTeX: actual_years as columns (could be wide — use small font)
    col_spec = "@{}l" + "r" * len(actual_years) + "@{}"
    header_years = " & ".join(f"\\bfseries {y}" for y in actual_years)

    tex1_lines = [
        r"\begin{table}[htbp]",
        r"  \centering",
        r"  \small",
        r"  \setlength{\tabcolsep}{6pt}",
        r"  \captiondatasource{Historical US BLS Engineering Employment by Discipline (thousands)}"
        r"{\cite{u.s.bureauoflaborstatisticsOccupationalProjectionsWorker2024}}",
        r"  \label{tab:litReview-employmentActuals}",
        f"  \\begin{{tabular}}{{{col_spec}}}",
        r"    \toprule",
        f"    \\bfseries Discipline & {header_years} \\\\",
        r"    \midrule",
    ]

    totals_actual = {y: 0.0 for y in actual_years}
    for code in ordered_codes:
        label = CODE_MAP[code]
        cells = []
        for y in actual_years:
            v = actual_map[code].get(y, np.nan)
            cells.append(fmt_k(v))
            if not np.isnan(v):
                totals_actual[y] += v
        tex1_lines.append(f"    {label} & " + " & ".join(cells) + r" \\")

    tex1_lines += [
        r"    \midrule",
        "    \\bfseries Total (above) & " +
        " & ".join(f"\\bfseries {fmt_k(totals_actual[y])}" for y in actual_years) +
        r" \\",
        r"    \bottomrule",
        r"  \end{tabular}",
        r"  \vspace{4pt}",
        r"  \parbox{0.98\linewidth}{",
        r"    \footnotesize\textit{\textsuperscript{a}\,Data correspond to"
        r" ``Architectural and Engineering Managers'' (SOC~11-9041), which"
        r" aggregates both architectural and engineering management roles. More information in companion white paper \cite{versteeg2026bls}.}",
        r"  }",
        r"\end{table}",
        "",
    ]
    with open(OUT_TEX1, "w", encoding="utf-8") as f:
        f.write("\n".join(tex1_lines))

    # CSV1
    csv1_rows = []
    for code in ordered_codes:
        row = {"Discipline": CODE_MAP[code].replace("\\", "").replace("{", "").replace("}", "")}
        for y in actual_years:
            row[str(y)] = actual_map[code].get(y, "")
        csv1_rows.append(row)
    pd.DataFrame(csv1_rows).to_csv(OUT_CSV1, index=False, sep=";")

    # ── TABLE 2: PROJECTIONS (growth %) ───────────────────────────────────────
    # Columns: one per BLS edition, labeled "YYYY→YYYY"
    edition_labels = [f"{e['base_year']}$\\to${e['proj_year']}" for e in editions]
    ed_proj_years  = [e["proj_year"] for e in editions]

    col_spec2 = "@{}l" + "r" * len(editions) + "@{}"
    header_eds = " & ".join(f"\\bfseries {lbl}" for lbl in edition_labels)

    # Footnote: dagger on realized editions
    realized_marks = []
    for i, ed in enumerate(editions):
        if ed["proj_year"] in realized_proj_years:
            realized_marks.append(f"{ed['base_year']}\\to{ed['proj_year']}")

    tex2_lines = [
        r"\begin{table}[htbp]",
        r"  \centering",
        r"  \small",
        r"  \setlength{\tabcolsep}{5pt}",
        r"  \captiondatasource{BLS Projected 10-Year Growth \% by Engineering Discipline and Edition}"
        r"{\cite{u.s.bureauoflaborstatisticsOccupationalProjectionsWorker2024}}",
        r"  \label{tab:litReview-employmentProjections}",
        f"  \\begin{{tabular}}{{{col_spec2}}}",
        r"    \toprule",
        f"    \\bfseries Discipline & {header_eds} \\\\",
    ]

    # Realized row marker
    realized_row = "    Realized? & " + " & ".join(
        r"\checkmark" if ed["proj_year"] in realized_proj_years else "---"
        for ed in editions
    ) + r" \\"
    tex2_lines += [realized_row, r"    \midrule"]

    # Per-discipline growth rows
    total_growths = {}  # ed_proj_year → (total_base, total_proj) for sum growth
    for ed in editions:
        b_total = sum(b for _, b, _ in ed["records"])
        p_total = sum(p for _, _, p in ed["records"])
        total_growths[ed["proj_year"]] = (b_total, p_total)

    for code in ordered_codes:
        label = CODE_MAP[code]
        cells = []
        for ed in editions:
            entry = proj_map[code].get(ed["proj_year"])
            if entry is not None:
                cells.append(fmt_g(entry[0]))
            else:
                cells.append("---")
        tex2_lines.append(f"    {label} & " + " & ".join(cells) + r" \\")

    # Summary total growth row
    summary_cells = []
    for ed in editions:
        b, p = total_growths.get(ed["proj_year"], (0, 0))
        g = (p - b) / b * 100 if b else np.nan
        summary_cells.append(fmt_g(g))
    tex2_lines += [
        r"    \midrule",
        "    \\bfseries Total (above) & " +
        " & ".join(f"\\bfseries {c}" for c in summary_cells) + r" \\",
        r"    \bottomrule",
        r"  \end{tabular}",
        r"  \vspace{4pt}",
        r"  \parbox{0.98\linewidth}{",
        r"    \footnotesize\textit{\checkmark~Projection target year has elapsed;"
        r" actual employment data are available for comparison with Table~\ref{tab:litReview-employmentActuals}.}",
        r"    \\[2pt]",
        r"    \textit{\textsuperscript{a}\,Data correspond to"
        r" ``Architectural and Engineering Managers'' (SOC~11-9041).}",
        r"  }",
        r"\end{table}",
        "",
    ]
    with open(OUT_TEX2, "w", encoding="utf-8") as f:
        f.write("\n".join(tex2_lines))

    # CSV2
    csv2_rows = []
    for code in ordered_codes:
        row = {"Discipline": CODE_MAP[code].replace("\\", "").replace("{", "").replace("}", "")}
        for ed in editions:
            entry = proj_map[code].get(ed["proj_year"])
            key = f"{ed['base_year']}_to_{ed['proj_year']}_growth_pct"
            row[key] = round(entry[0], 1) if entry else ""
        csv2_rows.append(row)
    pd.DataFrame(csv2_rows).to_csv(OUT_CSV2, index=False, sep=";")

    # ── FIGURE: TREND PLOT ────────────────────────────────────────────────────
    # Build sorted coordinate lists for actuals and projections
    act_pts = sorted(agg_actuals.items())   # [(year, emp_k), ...]
    prj_pts = sorted(agg_projs.items())     # [(proj_year, emp_k), ...]

    act_x = np.array([p[0] for p in act_pts], dtype=float)
    act_y = np.array([p[1] for p in act_pts], dtype=float)
    prj_x = np.array([p[0] for p in prj_pts], dtype=float)
    prj_y = np.array([p[1] for p in prj_pts], dtype=float)

    # Linear regression
    act_slope, act_intercept, *_ = stats.linregress(act_x, act_y)
    prj_slope, prj_intercept, *_ = stats.linregress(prj_x, prj_y)

    # Regression line domain: from earliest actual to latest projection
    x_min = int(min(act_x))
    x_max = int(max(prj_x))
    reg_xs = [x_min, x_max]

    def reg_coords(xs, slope, intercept):
        return " ".join(f"({x},{slope*x+intercept:.1f})" for x in xs)

    def pgf_coords(pts):
        return " ".join(f"({x},{y:.1f})" for x, y in pts)

    tex3_lines = [
        r"\begin{figure}[htbp]",
        r"  \centering",
        r"  \begin{tikzpicture}",
        r"    \begin{axis}[",
        r"      width=\linewidth,",
        r"      height=0.55\linewidth,",
        r"      xlabel={Year},",
        r"      ylabel={Total Engineering Employment (millions)},",
        r"      xmin=" + str(x_min - 1) + r", xmax=" + str(x_max + 1) + ",",
        r"      xtick distance=4,",
        r"      minor tick num=3,",
        r"      x tick label style={/pgf/number format/1000 sep={}},",
        r"      y filter/.code={\pgfmathparse{#1/1000}\pgfmathresult},",
        r"      yticklabel={\pgfmathprintnumber[fixed,precision=1,zerofill]{\tick}},",
        r"      ymajorgrids=true,",
        r"      grid style={dotted,gray!40},",
        r"      legend pos=south east,",
        r"      legend style={font=\footnotesize, draw=gray!50, fill=white, fill opacity=0.9},",
        r"      tick label style={font=\small},",
        r"      label style={font=\small},",
        r"    ]",
        # Actual data points + line
        r"    % --- Actual employment ---",
        r"    \addplot[",
        r"      black, solid, very thick,",
        r"      mark=*, mark size=2pt,",
        r"    ] coordinates {" + pgf_coords(act_pts) + "};",
        r"    \addlegendentry{Actual}",
        # Projected data points + line (dashed)
        r"    % --- Projected employment (one point per BLS edition) ---",
        r"    \addplot[",
        r"      black, dashed, thick,",
        r"      mark=square*, mark size=2pt,",
        r"    ] coordinates {" + pgf_coords(prj_pts) + "};",
        r"    \addlegendentry{Projected}",
        # Regression lines
        r"    % --- Linear regression: actuals ---",
        r"    \addplot[",
        r"      black, dotted, very thick,",
        r"    ] coordinates {" + reg_coords(reg_xs, act_slope, act_intercept) + "};",
        r"    \addlegendentry{Actual trend}",
        r"    % --- Linear regression: projections ---",
        r"    \addplot[",
        r"      gray, dashdotted, thick,",
        r"    ] coordinates {" + reg_coords(reg_xs, prj_slope, prj_intercept) + "};",
        r"    \addlegendentry{Projected trend}",
        r"    \end{axis}",
        r"  \end{tikzpicture}",
        r"  \captiondatasource{US Engineering Employment, Actual and Projected}"
        r"{\cite{u.s.bureauoflaborstatisticsOccupationalProjectionsWorker2024}}",
        r"  \label{fig:litReview-employmentTrend}",
        r"\end{figure}",
        "",
    ]
    with open(OUT_TEX3, "w", encoding="utf-8") as f:
        f.write("\n".join(tex3_lines))

    print()
    print(f"[ok]  BLS_actuals.tex       ({len(actual_years)} actual years)")
    print(f"[ok]  BLS_actuals.csv")
    print(f"[ok]  BLS_projections.tex   ({len(editions)} editions)")
    print(f"[ok]  BLS_projections.csv")
    print(f"[ok]  BLS_trend.tex         (actuals {int(min(act_x))}-{int(max(act_x))},"
          f" projections -> {int(max(prj_x))})")
    print(f"\n     Actual years:      {actual_years}")
    print(f"     Projection years:  {ed_proj_years}")
    print(f"     Already realized:  {sorted(realized_proj_years)}")


if __name__ == "__main__":
    main()