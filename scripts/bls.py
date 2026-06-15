#!/usr/bin/env python3
# Copyright (c) 2026 John VerSteeg, South Dakota State University
# SPDX-License-Identifier: CC-BY-4.0
"""
bls.py
------------------------------------------------------------
Author: John VerSteeg (South Dakota State University)
Date:   2025-10-15

Purpose -- Generates:
  1) outputs/BLS_table.tex  (complete LaTeX table)
  2) outputs/BLS_table.csv  (semicolon-separated, pgfplotstable-friendly or other use)

SOURCE DATA:
  Download "occupation.xlsx" directly from BLS and place it in data/:
    https://www.bls.gov/emp/tables/occupational-projections-and-characteristics.htm
  Uses "Table 1.2: Occupational projections, YYYY--YYYY" inside that workbook.

WHAT IT DOES:
  - Auto-detects baseline/projection years (e.g., 2024 & 2034).
  - Selects engineering-related occupations (SOC 17-2****)
    + Sales Engineers (41-9031)
    + Architectural & Engineering Managers (11-9041)
  - Converts employment (thousands) -> actual counts
  - Computes growth % from base/projection
  - Adds summary rows (Above Disciplines, Total US Workforce)
  - Writes both LaTeX and CSV outputs for flexible reuse.

DEPENDENCIES:
  Python 3.8+, pandas, openpyxl
  Install via:
      python -m pip install --upgrade pandas openpyxl

PATHS:
  Resolves paths relative to the repository root (the parent of scripts/):
  reads data/occupation.xlsx and writes outputs/BLS_table.{tex,csv},
  creating outputs/ if needed. Run from anywhere, e.g. `python scripts/bls.py`.
"""

import os
import re

import pandas as pd

# ---------- paths ----------
# Resolve everything relative to the repository root (the parent of scripts/),
# so the script can be run from any working directory.
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(SCRIPT_DIR)
DATA_DIR = os.path.join(REPO_ROOT, "data")
OUT_DIR = os.path.join(REPO_ROOT, "outputs")

XLSX = os.path.join(DATA_DIR, "occupation.xlsx")
OUT_TEX = os.path.join(OUT_DIR, "BLS_table.tex")
OUT_CSV = os.path.join(OUT_DIR, "BLS_table.csv")  # semicolon-separated


# ---------- helpers ----------
def norm(s: str) -> str:
    """Collapse whitespace and strip."""
    return re.sub(r"\s+", " ", str(s).strip())


def latex_escape(text: str) -> str:
    """
    Escape LaTeX-sensitive characters unless text already
    contains LaTeX markup (e.g., \\textsuperscript{a}).
    """
    t = str(text)
    if "\\" in t:
        return t
    repl = {
        "&": r"\&", "%": r"\%", "_": r"\_", "#": r"\#",
        "$": r"\$", "{": r"\{", "}": r"\}",
    }
    return "".join(repl.get(ch, ch) for ch in t)


def detect_year_columns(columns) -> tuple[str, str, int, int]:
    """Find two 'Employment, YYYY' columns and return (col_base, col_proj, base_year, proj_year)."""
    year_cols = []
    for c in columns:
        m = re.match(r"\s*Employment,\s*(\d{4})\s*$", str(c))
        if m:
            year_cols.append((int(m.group(1)), str(c)))
    if len(year_cols) < 2:
        raise ValueError("Could not find two 'Employment, YYYY' columns in Table 1.2.")
    year_cols.sort(key=lambda x: x[0])
    base_year, col_base = year_cols[0]
    proj_year, col_proj = year_cols[-1]
    return col_base, col_proj, base_year, proj_year


def detect_matrix_columns(columns) -> tuple[str, str]:
    """
    Find the 'YYYY National Employment Matrix title' and '... code' columns
    regardless of the edition year embedded in the header.
    """
    col_title = col_code = None
    for c in columns:
        s = str(c).strip()
        if re.match(r"\d{4}\s+National Employment Matrix title", s):
            col_title = c
        elif re.match(r"\d{4}\s+National Employment Matrix code", s):
            col_code = c
    if col_title is None or col_code is None:
        raise ValueError(
            "Could not find 'YYYY National Employment Matrix title/code' columns.\n"
            f"Available columns: {list(columns)}"
        )
    return col_title, col_code


def fmt_int(n: int) -> str:
    return f"{n:,}"


def fmt_pct(x) -> str:
    return f"{x:.1f}\\%"


# ---------- discipline mapping ----------
# Keys are lowercase for case-insensitive matching against BLS titles.
TITLE_MAP = {
    "architectural and engineering managers": "Engineering Managers",
    "aerospace engineers": "Aerospace",
    "agricultural engineers": "Agricultural",
    "bioengineers and biomedical engineers": "Biomedical",
    "chemical engineers": "Chemical",
    "civil engineers": "Civil",
    "computer hardware engineers": "Computer Hardware",
    "electrical engineers": "Electrical",
    "electronics engineers, except computer": "Electronics",
    "engineers, all other": "Engineers, All Other",
    "environmental engineers": "Environmental",
    "health and safety engineers, except mining safety engineers and inspectors": "Health & Safety",
    "industrial engineers": "Industrial",
    "marine engineers and naval architects": "Marine",
    "materials engineers": "Materials",
    "mechanical engineers": "Mechanical",
    "mining and geological engineers, including mining safety engineers": "Mining",
    "nuclear engineers": "Nuclear",
    "petroleum engineers": "Petroleum",
    "sales engineers": "Sales Engineers",
}


# ---------- main ----------
def main():
    # --- guards ---
    try:
        import openpyxl  # noqa: F401
    except ModuleNotFoundError:
        raise SystemExit(
            "Missing dependency: openpyxl\n"
            "Install with: python -m pip install openpyxl"
        )

    if not os.path.exists(XLSX):
        raise FileNotFoundError(
            f"Cannot find input workbook: {XLSX}\n"
            "Download 'occupation.xlsx' from the BLS Occupational Projections page "
            "and place it in the data/ directory."
        )

    os.makedirs(OUT_DIR, exist_ok=True)

    # --- load & prepare ---
    raw = pd.read_excel(XLSX, sheet_name="Table 1.2", header=None)
    header = raw.iloc[1].tolist()
    df = raw.iloc[2:].copy()
    df.columns = header

    col_title, col_code = detect_matrix_columns(df.columns)
    col_base, col_proj, base_year, proj_year = detect_year_columns(df.columns)

    df["title_norm"] = df[col_title].map(norm)
    codes = df[col_code].astype(str)

    # --- filter engineering occupations ---
    subset = df[
        codes.str.startswith("17-2") | codes.eq("41-9031") | codes.eq("11-9041")
    ].copy()
    subset["Discipline"] = subset["title_norm"].str.lower().map(TITLE_MAP)
    subset = subset[subset["Discipline"].notna()].copy()

    # --- compute employment & growth ---
    e_base = pd.to_numeric(subset[col_base], errors="coerce") * 1000
    e_proj = pd.to_numeric(subset[col_proj], errors="coerce") * 1000
    subset["Employment_base"] = e_base.round().astype("Int64")
    subset["Employment_proj"] = e_proj.round().astype("Int64")
    subset["Growth_pct"] = ((e_proj - e_base) / e_base * 100).round(1)

    total_base = int(subset["Employment_base"].sum())
    total_proj = int(subset["Employment_proj"].sum())
    growth_total = round((total_proj - total_base) / total_base * 100, 1)

    # --- total US workforce ---
    us_row = df.loc[codes.eq("00-0000")]
    if not us_row.empty:
        us_base = int(round(float(us_row[col_base].iloc[0]) * 1000))
        us_proj = int(round(float(us_row[col_proj].iloc[0]) * 1000))
        us_growth = round((us_proj - us_base) / us_base * 100, 1)
    else:
        us_base = us_proj = 0
        us_growth = 0.0

    # --- display names (LaTeX footnote marker on Eng Managers) ---
    subset["Discipline_disp"] = subset["Discipline"].replace(
        {"Engineering Managers": "Engineering Managers\\textsuperscript{a}"}
    )
    ordered = (
        subset[["Discipline_disp", "Employment_base", "Employment_proj", "Growth_pct"]]
        .sort_values("Employment_base", ascending=False)
        .reset_index(drop=True)
    )

    # --- write TEX ---
    lines = [
        r"\begin{table}[htbp]",
        r"  \centering",
        r"  \setlength{\tabcolsep}{14pt}",
        r"  \captiondatasource{US BLS Engineering Employment Projections by Discipline}"
        r"{\cite{u.s.bureauoflaborstatisticsOccupationalProjectionsWorker2024}}",
        r"  \label{tab:litReview-employmentTrends}",
        r"  \begin{tabular}{@{}lrrr@{}}",
        r"    \toprule",
        f"    \\bfseries Engineering Discipline"
        f" & \\bfseries {base_year} Employment"
        f" & \\bfseries {proj_year} Projection"
        f" & \\bfseries Growth \\% \\\\",
        r"    \midrule",
    ]

    for _, row in ordered.iterrows():
        discipline = latex_escape(row["Discipline_disp"])
        e0 = fmt_int(int(row["Employment_base"]))
        e1 = fmt_int(int(row["Employment_proj"]))
        g = fmt_pct(float(row["Growth_pct"]))
        lines.append(f"    {discipline} & {e0} & {e1} & {g} \\\\")

    lines += [
        r"    \midrule",
        f"    Above Disciplines & {fmt_int(total_base)} & {fmt_int(total_proj)} & {fmt_pct(growth_total)} \\\\",
        f"    Total US Workforce & {fmt_int(us_base)} & {fmt_int(us_proj)} & {fmt_pct(us_growth)} \\\\",
        r"    \bottomrule",
        r"  \end{tabular}",
        r"  \vspace{4pt}",
        r"",
        r"  \parbox{0.98\linewidth}{",
        r"    \footnotesize\textit{\textsuperscript{a}\,Data correspond to the Bureau of Labor"
        r" Statistics occupation ``Architectural and Engineering Managers''"
        r" (SOC~11-9041), which aggregates both architectural and engineering management roles. More information in companion white paper \cite{versteeg2026bls}.}",
        r"  }",
        r"\end{table}",
        "",
    ]
    with open(OUT_TEX, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    # --- write CSV ---
    csv_df = ordered.copy()
    csv_df["Discipline_disp"] = csv_df["Discipline_disp"].str.replace(
        r"\\textsuperscript\{a\}", "", regex=True
    )
    csv_df = csv_df.rename(columns={
        "Discipline_disp": "Discipline",
        "Employment_base": f"Employment_{base_year}",
        "Employment_proj": f"Employment_{proj_year}",
    })
    csv_df = pd.concat([
        csv_df,
        pd.DataFrame([
            {"Discipline": "Above Disciplines",
             f"Employment_{base_year}": total_base,
             f"Employment_{proj_year}": total_proj,
             "Growth_pct": growth_total},
            {"Discipline": "Total US Workforce",
             f"Employment_{base_year}": us_base,
             f"Employment_{proj_year}": us_proj,
             "Growth_pct": us_growth},
        ])
    ], ignore_index=True)
    csv_df.to_csv(OUT_CSV, index=False, sep=";")

    print(f"[ok] Wrote {OUT_TEX}")
    print(f"[ok] Wrote {OUT_CSV} (sep=';')  | Years detected: {base_year} -> {proj_year}")


if __name__ == "__main__":
    main()
