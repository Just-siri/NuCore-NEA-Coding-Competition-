"""
Validation script for the final Risk Register Standardization model.
OECD NEA Coding Competition — NuCore

What this script does:
1. Runs model.py on the three training input files.
2. Looks for matching reference Final workbooks.
3. Compares the generated Simplified Register sheet to the reference.
4. Reports row-count match, mandatory field presence, and value match rate.

This script is designed for the final calibrated model.
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

BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "test_outputs"

sys.path.insert(0, str(BASE_DIR))
from model import RiskRegisterStandardizer

MANDATORY_KEYWORDS = [
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

TEST_CASES = [
    {
        "name": "IVC DOE R2",
        "input_name": "1. IVC DOE R2 (Input).xlsx",
        "final_pattern": "1. IVC DOE",
        "output_name": "1. IVC DOE R2 (Test Final).xlsx",
    },
    {
        "name": "City of York Council",
        "input_name": "2. City of York Council (Input).xlsx",
        "final_pattern": "2. City of York Council",
        "output_name": "2. City of York Council (Test Final).xlsx",
    },
    {
        "name": "Digital Security IT Sample Register",
        "input_name": "3. Digital Security IT Sample Register (Input).xlsx",
        "final_pattern": "3. Digital Security IT Sample Register",
        "output_name": "3. Digital Security IT Sample Register (Test Final).xlsx",
    },
]

SEARCH_DIRS = [INPUT_DIR, BASE_DIR / "output", BASE_DIR, BASE_DIR / "finals", Path("/mnt/data")]


def find_final_file(pattern: str) -> Optional[Path]:
    pattern = pattern.lower()
    for directory in SEARCH_DIRS:
        if not directory.exists():
            continue
        for file in directory.iterdir():
            name = file.name.lower()
            if file.suffix.lower() == ".xlsx" and pattern in name and "final" in name:
                return file
    return None


def pick_data_sheet(xl: pd.ExcelFile) -> str:
    for name in xl.sheet_names:
        nl = name.lower()
        if "requirement" in nl or "output" in nl:
            continue
        if "simplified" in nl or "register" in nl or "risk" in nl:
            return name
    return xl.sheet_names[0]


def read_register(file_path: Path) -> pd.DataFrame:
    xl = pd.ExcelFile(file_path)
    sheet_name = pick_data_sheet(xl)
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)

    # Remove hidden row-2 reference-letter row if present
    if len(df.columns) > 0:
        first_col = df.columns[0]
        mask = df[first_col].astype(str).str.fullmatch(r"[A-Za-z]", na=False)
        df = df[~mask].reset_index(drop=True)

    return df


def find_col(df: pd.DataFrame, keyword: str) -> Optional[str]:
    keyword = keyword.lower()
    for col in df.columns:
        if keyword in str(col).lower():
            return col
    return None


def normalize_value(value):
    if pd.isna(value):
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value).strip()
    return str(value).strip()


def compare_outputs(generated_file: Path, expected_file: Path) -> dict:
    df_gen = read_register(generated_file)
    df_exp = read_register(expected_file)

    missing = [kw for kw in MANDATORY_KEYWORDS if find_col(df_gen, kw) is None]
    row_match = len(df_gen) == len(df_exp)

    total_cells = 0
    matched_cells = 0
    rows_to_compare = min(len(df_gen), len(df_exp))

    for kw in MANDATORY_KEYWORDS:
        gc = find_col(df_gen, kw)
        ec = find_col(df_exp, kw)
        if gc is None or ec is None:
            continue
        for i in range(rows_to_compare):
            gv = normalize_value(df_gen.iloc[i][gc]).lower()
            ev = normalize_value(df_exp.iloc[i][ec]).lower()
            total_cells += 1
            if gv == ev:
                matched_cells += 1

    value_match_rate = (matched_cells / total_cells * 100) if total_cells else 0.0

    return {
        "missing": missing,
        "row_match": row_match,
        "rows_generated": len(df_gen),
        "rows_expected": len(df_exp),
        "value_match_rate": value_match_rate,
    }


def main() -> None:
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        print("Error: ANTHROPIC_API_KEY is not set.")
        sys.exit(1)

    OUTPUT_DIR.mkdir(exist_ok=True)

    try:
        standardizer = RiskRegisterStandardizer(api_key=api_key)
    except Exception as exc:
        print(f"Failed to initialize model: {exc}")
        sys.exit(1)

    print("=" * 80)
    print("FINAL MODEL VALIDATION")
    print("=" * 80)

    results = []

    for case in TEST_CASES:
        print(f"
--- {case['name']} ---")
        input_path = INPUT_DIR / case["input_name"]
        output_path = OUTPUT_DIR / case["output_name"]

        if not input_path.exists():
            print(f"Missing input: {input_path}")
            results.append((case["name"], "FAIL", "input missing"))
            continue

        ok = standardizer.process_file(str(input_path), str(output_path))
        if not ok:
            results.append((case["name"], "FAIL", "processing failed"))
            continue

        final_path = find_final_file(case["final_pattern"])
        if final_path is None:
            print("Reference final not found; output generated only.")
            results.append((case["name"], "GENERATED", "no reference final"))
            continue

        try:
            comparison = compare_outputs(output_path, final_path)
        except Exception as exc:
            traceback.print_exc()
            results.append((case["name"], "FAIL", str(exc)))
            continue

        status = "PASS" if (not comparison["missing"] and comparison["row_match"]) else "PARTIAL"
        print(f"Rows: generated={comparison['rows_generated']} expected={comparison['rows_expected']}")
        print(f"Mandatory columns missing: {comparison['missing'] or 'None'}")
        print(f"Value match rate: {comparison['value_match_rate']:.2f}%")
        results.append((case["name"], status, comparison))

    print("
" + "=" * 80)
    print("SUMMARY")
    print("=" * 80)
    for name, status, details in results:
        print(f"- {name}: {status}")
        if isinstance(details, dict):
            print(f"  value match rate = {details['value_match_rate']:.2f}%")
        else:
            print(f"  {details}")

    print(f"
Test outputs written to: {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
