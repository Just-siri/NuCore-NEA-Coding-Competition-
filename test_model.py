"""
Test and Validation Script
OECD NEA Coding Competition — NuCore

Runs model.py against the three training input files, compares each
generated output to its reference Final workbook, and reports:

  - Whether all mandatory output columns are present
  - Whether the generated row count matches the reference
  - The value match rate across all mandatory fields

Usage:
    python test_model.py

Prerequisites:
    - Input files must be present in the input/ directory
    - Reference Final workbooks must be accessible (searched automatically)
    - ANTHROPIC_API_KEY must be set in the environment or a .env file
"""

from __future__ import annotations

import os
import sys
import traceback
from pathlib import Path
from typing import Optional

import pandas as pd
from dotenv import load_dotenv

load_dotenv()

# ──────────────────────────────────────────────────────────────
#  PATHS
# ──────────────────────────────────────────────────────────────

BASE_DIR   = Path(__file__).parent
INPUT_DIR  = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "test_outputs"

sys.path.insert(0, str(BASE_DIR))
from model import RiskRegisterStandardizer

# ──────────────────────────────────────────────────────────────
#  CONFIGURATION
# ──────────────────────────────────────────────────────────────

# Mandatory columns — every generated output must contain all of these
MANDATORY_COLUMNS = [
    "Risk ID",
    "Risk Description",
    "Project Stage",
    "Project Category",
    "Risk Owner",
    "Mitigating Action",
    "Likelihood",
    "Impact",
    "Risk Priority",
]

# Test cases — the three training input–output pairs
TEST_CASES = [
    {
        "name":          "IVC DOE R2",
        "input_name":    "1. IVC DOE R2 (Input).xlsx",
        "final_pattern": "1. IVC DOE",
        "output_name":   "1. IVC DOE R2 (Test Final).xlsx",
    },
    {
        "name":          "City of York Council",
        "input_name":    "2. City of York Council (Input).xlsx",
        "final_pattern": "2. City of York Council",
        "output_name":   "2. City of York Council (Test Final).xlsx",
    },
    {
        "name":          "Digital Security IT Sample Register",
        "input_name":    "3. Digital Security IT Sample Register (Input).xlsx",
        "final_pattern": "3. Digital Security IT Sample Register",
        "output_name":   "3. Digital Security IT Sample Register (Test Final).xlsx",
    },
]

# Directories searched when looking for reference Final workbooks
REFERENCE_SEARCH_DIRS = [
    INPUT_DIR,
    BASE_DIR / "output",
    BASE_DIR / "finals",
    BASE_DIR,
]

# ──────────────────────────────────────────────────────────────
#  FILE UTILITIES
# ──────────────────────────────────────────────────────────────

def find_reference_file(pattern: str) -> Optional[Path]:
    """
    Search REFERENCE_SEARCH_DIRS for an .xlsx file whose name contains
    both `pattern` and the word "final" (case-insensitive).

    Returns the first match found, or None if no match exists.
    """
    pattern_lower = pattern.lower()
    for directory in REFERENCE_SEARCH_DIRS:
        if not directory.exists():
            continue
        for file in sorted(directory.iterdir()):
            name = file.name.lower()
            if (
                file.suffix.lower() == ".xlsx"
                and pattern_lower in name
                and "final" in name
            ):
                return file
    return None


def pick_data_sheet(xl: pd.ExcelFile) -> str:
    """
    Return the name of the data sheet in a workbook, skipping any sheet
    whose name suggests it is a requirements or output-reference sheet.
    """
    for name in xl.sheet_names:
        nl = name.lower()
        if "requirement" in nl or "output" in nl:
            continue
        if any(k in nl for k in ("simplified", "register", "risk")):
            return name
    return xl.sheet_names[0]


def read_register(file_path: Path) -> pd.DataFrame:
    """
    Load the data sheet from a workbook into a DataFrame.

    Automatically strips the hidden Row 2 reference-letter row used in
    the pre/post mitigation layout (Layout A), if present.
    """
    xl          = pd.ExcelFile(file_path)
    sheet_name  = pick_data_sheet(xl)
    df          = pd.read_excel(file_path, sheet_name=sheet_name, header=0)

    if len(df.columns) > 0:
        first_col = df.columns[0]
        # Row 2 in Layout A contains single letters (A, N, G, …) as column refs
        hidden_row_mask = (
            df[first_col].astype(str).str.fullmatch(r"[A-Za-z]", na=False)
        )
        df = df[~hidden_row_mask].reset_index(drop=True)

    return df

# ──────────────────────────────────────────────────────────────
#  COMPARISON UTILITIES
# ──────────────────────────────────────────────────────────────

def find_column(df: pd.DataFrame, keyword: str) -> Optional[str]:
    """
    Return the first column name in `df` that contains `keyword`
    (case-insensitive), or None if no such column exists.
    """
    keyword_lower = keyword.lower()
    for col in df.columns:
        if keyword_lower in str(col).lower():
            return col
    return None


def normalize(value) -> str:
    """
    Normalise a cell value to a stripped string for comparison.
    Whole-number floats are converted to integer strings (e.g. 3.0 → "3").
    NaN and None become empty strings.
    """
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def compare_workbooks(generated: Path, reference: Path) -> dict:
    """
    Load both workbooks and compare them across all mandatory columns.

    Returns a result dictionary containing:
        missing       — list of mandatory columns absent from the generated output
        row_match     — True if row counts are equal
        rows_gen      — number of data rows in the generated output
        rows_ref      — number of data rows in the reference
        match_rate    — percentage of matching cells across mandatory columns
        mismatches    — list of (column, row_index, generated, reference) tuples
                        for the first 10 mismatches found per column
    """
    df_gen = read_register(generated)
    df_ref = read_register(reference)

    missing   = [kw for kw in MANDATORY_COLUMNS if find_column(df_gen, kw) is None]
    row_match = len(df_gen) == len(df_ref)
    n_rows    = min(len(df_gen), len(df_ref))

    total_cells   = 0
    matched_cells = 0
    mismatches    = []

    for keyword in MANDATORY_COLUMNS:
        col_gen = find_column(df_gen, keyword)
        col_ref = find_column(df_ref, keyword)
        if col_gen is None or col_ref is None:
            continue
        col_mismatches = 0
        for i in range(n_rows):
            gv = normalize(df_gen.iloc[i][col_gen]).lower()
            rv = normalize(df_ref.iloc[i][col_ref]).lower()
            total_cells += 1
            if gv == rv:
                matched_cells += 1
            elif col_mismatches < 10:
                mismatches.append((keyword, i + 1, gv, rv))
                col_mismatches += 1

    match_rate = (matched_cells / total_cells * 100) if total_cells > 0 else 0.0

    return {
        "missing":    missing,
        "row_match":  row_match,
        "rows_gen":   len(df_gen),
        "rows_ref":   len(df_ref),
        "match_rate": match_rate,
        "mismatches": mismatches,
    }

# ──────────────────────────────────────────────────────────────
#  REPORTING
# ──────────────────────────────────────────────────────────────

def print_result(name: str, result: dict, show_mismatches: bool = False) -> str:
    """
    Print the comparison result for one test case and return its status string.
    """
    no_missing  = not result["missing"]
    rows_ok     = result["row_match"]
    status      = "PASS" if (no_missing and rows_ok) else "PARTIAL"

    row_label = (
        f"{result['rows_gen']} rows  ✓"
        if rows_ok
        else f"{result['rows_gen']} generated / {result['rows_ref']} expected  ✗"
    )
    missing_label = "None" if no_missing else ", ".join(result["missing"])

    print(f"  Status            : {status}")
    print(f"  Rows              : {row_label}")
    print(f"  Missing columns   : {missing_label}")
    print(f"  Value match rate  : {result['match_rate']:.1f}%")

    if show_mismatches and result["mismatches"]:
        print("  Sample mismatches :")
        for col, row, got, exp in result["mismatches"][:5]:
            print(f"    Row {row:>2} [{col}]")
            print(f"      generated : {got!r}")
            print(f"      reference : {exp!r}")

    return status

# ──────────────────────────────────────────────────────────────
#  MAIN
# ──────────────────────────────────────────────────────────────

def main() -> None:
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("Error: ANTHROPIC_API_KEY is not set.")
        print("Set it in a .env file or export it as an environment variable.")
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)

    try:
        standardizer = RiskRegisterStandardizer(api_key=api_key)
    except Exception as exc:
        print(f"Failed to initialize model: {exc}")
        sys.exit(1)

    LINE = "=" * 72

    print(LINE)
    print("  RISK REGISTER MODEL — VALIDATION REPORT")
    print(LINE)

    summary: list[tuple[str, str]] = []

    for case in TEST_CASES:
        print(f"\n{case['name']}")
        print("-" * len(case["name"]))

        input_path  = INPUT_DIR  / case["input_name"]
        output_path = OUTPUT_DIR / case["output_name"]

        # ── Step 1: check input exists ────────────────────────
        if not input_path.exists():
            print(f"  ✗ Input file not found: {input_path}")
            summary.append((case["name"], "FAIL — input missing"))
            continue

        # ── Step 2: run the model ─────────────────────────────
        print(f"  Running model…")
        ok = standardizer.process_file(str(input_path), str(output_path))
        if not ok:
            print("  ✗ Model processing failed.")
            summary.append((case["name"], "FAIL — processing error"))
            continue
        print(f"  Output written → {output_path.name}")

        # ── Step 3: locate reference Final ────────────────────
        reference_path = find_reference_file(case["final_pattern"])
        if reference_path is None:
            print("  Reference Final not found — output generated only.")
            summary.append((case["name"], "GENERATED — no reference"))
            continue
        print(f"  Reference found → {reference_path.name}")

        # ── Step 4: compare outputs ───────────────────────────
        try:
            result = compare_workbooks(output_path, reference_path)
        except Exception as exc:
            print(f"  ✗ Comparison failed: {exc}")
            traceback.print_exc()
            summary.append((case["name"], f"FAIL — {exc}"))
            continue

        status = print_result(case["name"], result, show_mismatches=True)
        summary.append((case["name"], f"{status}  ({result['match_rate']:.1f}% match)"))

    # ── Summary ───────────────────────────────────────────────
    print(f"\n{LINE}")
    print("  SUMMARY")
    print(LINE)
    for name, status in summary:
        print(f"  {name:<45} {status}")

    print(f"\n  Test outputs: {OUTPUT_DIR}\n")


if __name__ == "__main__":
    main()
