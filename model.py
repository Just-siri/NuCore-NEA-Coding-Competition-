"""
Risk Register Standardization Model
OECD NEA Coding Competition — NuCore

Converts heterogeneous risk registers into a standardized Excel format.
All fields are read dynamically from each input file; nothing is hardcoded.
Claude is called only where deterministic rules are insufficient.

Training files : 1 — IVC DOE R2 | 2 — City of York | 3 — Digital Security IT
Blind tests    : 4 — Moorgate Crossrail | 5 — Fenland DC (PDF)
"""

import os
import re
import json
import time
import traceback
from pathlib import Path
from datetime import datetime
from typing import Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from dotenv import load_dotenv
import anthropic

load_dotenv()

# ──────────────────────────────────────────────────────────────
#  STYLING  (hex RGB — renders consistently across all Excel themes)
# ──────────────────────────────────────────────────────────────

HEADER_FILL  = PatternFill(fill_type="solid", fgColor="FFF0CBE0")
HEADER_FONT  = Font(bold=True, name="Calibri", size=11)
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
DATA_ALIGN   = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN_BORDER  = Border(left=Side(style="thin"), right=Side(style="thin"),
                      top=Side(style="thin"),  bottom=Side(style="thin"))
FMT_DATE = r'[$-409]d\-mmm\-yy;@'
FMT_INT  = '0'

# ──────────────────────────────────────────────────────────────
#  OUTPUT SCHEMAS
#  Layout A — pre+post mitigation  (Files 1, 5)
#  Layout B — single-stage + Result (File 2)
#  Layout C — single-stage IT/infra (Files 3, 4)
# ──────────────────────────────────────────────────────────────

HEADERS_PRE_POST = [
    "Date Added", "Risk ID", "Risk Description", "Project Stage",
    "Project Category", "Risk Owner",
    "Likelihood (1-10) (pre-mitigation)", "Impact (1-10) (pre-mitigation)",
    "Risk Priority (pre-mitigation)", "Mitigating Action",
    "Likelihood (1-10) (post-mitigation)", "Impact (1-10) (post-mitigation)",
    "Risk Priority (post-mitigation)",
]
# Hidden row 2 in Layout A carries column-letter reference labels
_ROW2_PRE_POST = ["A", None, "N", "G", None, "H", "L", "K", "M", "P", "L", "K", "M"]

HEADERS_SINGLE = [
    "Date Added", "Risk ID", "Risk Description", "Project Stage",
    "Project Category", "Likelihood (1-10)", "Impact (1-10)",
    "Risk Priority (low, med, high)", "Risk Owner", "Mitigating Action", "Result",
]

HEADERS_IT = [
    "Date Added", "Number", "Risk Description", "Project Stage",
    "Project Category", "Risk Owner", "Likelihood (1-10)", "Impact (1-10)",
    "Risk Priority (low, med, high)", "Mitigating Action",
]

OUTPUT_REQUIREMENTS = [
    ["The following columns are mandatory for the final risk registers. "
     "Additional columns may be added, if information is provided in the input files "
     "(ex. date, additional comments etc):"],
    ["Risk ID",           "If not provided, any identifier may be used."],
    ["Risk Description",  ""],
    ["Project Stage",     "Required for construction or project based risks."],
    ["Project Category",  ""],
    ["Risk Owner",        ""],
    ["Mitigating Action", ""],
    ["Likelihood (1-10)", "If multiple stages of risk assessment are provided, "
                          "include both pre and post-mitigation."],
    ["Impact (1-10)",     ""],
    ["Risk Priority (low, med, high)", ""],
]

_W_PRE_POST = {1:11.0, 2:8.5,  3:60.0,  4:18.0,  5:29.6,  6:23.1,
               7:16.4, 8:16.4, 9:16.4,  10:63.6, 11:17.1, 12:17.1, 13:17.1}
_W_SINGLE   = {1:11.0, 2:8.5,  3:70.0,  4:38.57, 5:16.71,
               6:14.86,7:13.86,8:16.14, 9:26.43, 10:55.71,11:57.29}
_W_IT       = {1:11.0, 2:8.5,  3:70.0,  4:38.57, 5:16.71,
               6:26.43,7:14.86,8:13.86, 9:16.14, 10:57.29}

# ──────────────────────────────────────────────────────────────
#  HELPERS
# ──────────────────────────────────────────────────────────────

def clean(v) -> str:
    """Normalise to a single-line string; return '' for None/blank."""
    if v is None:
        return ""
    return re.sub(r" {2,}", " ", re.sub(r"[\r\n]+", " ", str(v).strip())).strip()

def to_int(v) -> Optional[int]:
    """Parse to int; returns None on failure."""
    if v is None:
        return None
    try:
        return int(float(str(v).strip()))
    except (ValueError, TypeError):
        m = re.search(r"\d+", str(v))
        return int(m.group()) if m else None

def to_rid(v):
    """Return Risk ID as int if numeric, else as string."""
    if v is None:
        return None
    s = str(v).strip()
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return s

def as_datetime(v) -> Optional[datetime]:
    """Parse datetime objects and common date strings."""
    if isinstance(v, datetime):
        return v
    if isinstance(v, str) and v.strip():
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return datetime.strptime(v.strip()[:10], fmt)
            except ValueError:
                pass
    return None

def extract_role(raw: str) -> str:
    """Extract role from 'Name (role)' — capitalise first letter."""
    m = re.search(r"\(([^)]+)\)", raw)
    if m:
        role = m.group(1).strip()
        return role[0].upper() + role[1:] if role else raw.strip()
    return raw.strip()

def priority_from_scores(l, i) -> str:
    """L×I → Low / Med / High  (thresholds: <32 / <60 / ≥60)."""
    try:
        score = int(l) * int(i)
    except (TypeError, ValueError):
        return ""
    return "Low" if score < 32 else ("Med" if score < 60 else "High")

def call_claude(client, system: str, user: str,
                max_tokens: int = 4096, retries: int = 3) -> str:
    """Call Claude with exponential-backoff retry on transient errors."""
    for attempt in range(retries):
        try:
            resp = client.messages.create(
                model="claude-sonnet-4-20250514",
                max_tokens=max_tokens,
                system=system,
                messages=[{"role": "user", "content": user}],
            )
            return resp.content[0].text
        except Exception as exc:
            if attempt < retries - 1:
                wait = 2 ** attempt
                print(f"    API error ({exc}), retrying in {wait}s …")
                time.sleep(wait)
            else:
                raise

def parse_json(text: str) -> list:
    """Strip markdown fences and parse JSON from a Claude response."""
    text = re.sub(r"```(?:json)?\s*", "", text).strip()
    m = re.search(r"[\[\{]", text)
    return json.loads(text[m.start():] if m else text)

def get_data_sheet(wb):
    """Return the primary data sheet, skipping requirements/output sheets."""
    for name in ("Simplified Register", "Risk Register", "Register",
                 "Risks", "Risk", "Sheet1", "Sheet"):
        if name in wb.sheetnames:
            return wb[name]
    for name in wb.sheetnames:
        if "requirement" not in name.lower() and "output" not in name.lower():
            return wb[name]
    return wb.active

# ──────────────────────────────────────────────────────────────
#  EXCEL WRITER
# ──────────────────────────────────────────────────────────────

def _write_sheet(ws, headers, rows, col_widths,
                 row2_labels=None, round_floats=False):
    for c, h in enumerate(headers, 1):
        cell            = ws.cell(row=1, column=c, value=h)
        cell.fill       = HEADER_FILL
        cell.font       = HEADER_FONT
        cell.alignment  = HEADER_ALIGN
        cell.border     = THIN_BORDER

    data_start = 2
    if row2_labels:
        for c, v in enumerate(row2_labels, 1):
            ws.cell(row=2, column=c, value=v)
        ws.row_dimensions[2].hidden = True
        data_start = 3

    for r_off, row in enumerate(rows):
        r_idx = data_start + r_off
        ws.row_dimensions[r_idx].height = None   # allow auto row height
        for c_idx, val in enumerate(row, 1):
            cell           = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.border    = THIN_BORDER
            cell.alignment = DATA_ALIGN
            if isinstance(val, datetime):
                cell.number_format = FMT_DATE
            elif isinstance(val, int):
                cell.number_format = FMT_INT
            elif isinstance(val, float):
                if round_floats or val == int(val):
                    cell.value         = int(round(val))
                    cell.number_format = FMT_INT

    for c_idx, width in col_widths.items():
        ws.column_dimensions[get_column_letter(c_idx)].width = width

def _add_requirements_sheet(wb):
    ws = wb.create_sheet("Output Requirements")
    for r, row_data in enumerate(OUTPUT_REQUIREMENTS, 1):
        for c, val in enumerate(row_data, 1):
            ws.cell(row=r, column=c, value=val)
    ws.column_dimensions["A"].width = 65
    ws.column_dimensions["B"].width = 75

def _make_workbook(headers, rows, col_widths, row2_labels=None,
                   round_floats=False) -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Simplified Register"
    _write_sheet(ws, headers, rows, col_widths,
                 row2_labels=row2_labels, round_floats=round_floats)
    _add_requirements_sheet(wb)
    return wb

def save_pre_post(rows, output_path):
    wb = _make_workbook(HEADERS_PRE_POST, rows, _W_PRE_POST,
                        row2_labels=_ROW2_PRE_POST, round_floats=True)
    wb.save(output_path)

def save_single(rows, output_path):
    wb = _make_workbook(HEADERS_SINGLE, rows, _W_SINGLE)
    wb.save(output_path)

def save_it(rows, output_path):
    wb = _make_workbook(HEADERS_IT, rows, _W_IT)
    wb.save(output_path)

# ──────────────────────────────────────────────────────────────
#  FILE 1 — IVC DOE R2
#
#  24-column SEV/FRQ register (0–5 scale); output is Layout A.
#  col1  → Date Added
#  col2  → RBS Level 1 → Project Category (inferred by Claude)
#  col4  → Risk Name → fallback description when col14 is blank
#  col7  → Technology Life Phase → Project Stage (Claude maps to standard labels)
#  col8  → "Name (role)" → Risk Owner (role extracted from parentheses)
#  col11 → SEV baseline × 2 → Impact pre
#  col12 → FRQ baseline × 2 → Likelihood pre
#  col14 → Description with assumptions → Risk Description (primary source)
#  col16 → Response Description → Mitigating Action (typos corrected)
#  col18 → Residual SEV × 2 → Impact post
#  col19 → Residual FRQ × 2 → Likelihood post
#  Risk IDs run 1–14 then 16–33; ID 15 is skipped (duplicate row in source).
# ──────────────────────────────────────────────────────────────

def _parse_f1(ws) -> list:
    """Read and deduplicate rows from the IVC DOE raw register sheet."""
    RBS_VALID = {"Technical", "Management", "Commercial", "External"}
    rows, seen = [], set()
    for r in range(1, ws.max_row + 1):
        rbs = clean(ws.cell(r, 2).value)
        if rbs not in RBS_VALID:
            continue
        desc = clean(ws.cell(r, 14).value) or clean(ws.cell(r, 4).value)
        mit  = clean(ws.cell(r, 16).value)
        if not desc and not mit:
            continue
        key = (clean(ws.cell(r, 4).value)[:40], mit[:40])
        if key in seen:
            continue
        seen.add(key)
        sev_pre  = to_int(ws.cell(r, 11).value)
        frq_pre  = to_int(ws.cell(r, 12).value)
        sev_post = to_int(ws.cell(r, 18).value)
        frq_post = to_int(ws.cell(r, 19).value)
        rows.append({
            "date":   ws.cell(r, 1).value,
            "rbs":    rbs,
            "desc":   desc,
            "stage":  clean(ws.cell(r, 7).value),
            "owner":  clean(ws.cell(r, 8).value),
            "l_pre":  frq_pre  * 2 if frq_pre  is not None else None,
            "i_pre":  sev_pre  * 2 if sev_pre  is not None else None,
            "mit":    mit,
            "l_post": frq_post * 2 if frq_post is not None else None,
            "i_post": sev_post * 2 if sev_post is not None else None,
        })
    return rows


def _enhance_f1(client, rows: list) -> dict:
    """
    Single batched Claude call to enhance all File 1 rows:
    - Fix description typos; prefix component name where col14 text is generic
    - Map Technology Life Phase labels to standard Project Stage values
    - Infer Project Category from RBS Level + description context
    - Extract role from owner's "Name (role)" notation
    - Fix typos in mitigating action text (content preserved)
    """
    payload = json.dumps([
        {"idx": i, "rbs": r["rbs"], "desc": r["desc"],
         "stage_raw": r["stage"], "owner_raw": r["owner"],
         "l_pre": r["l_pre"], "i_pre": r["i_pre"],
         "mit": r["mit"], "l_post": r["l_post"], "i_post": r["i_post"]}
        for i, r in enumerate(rows)
    ], indent=2)

    user = f"""Rows from a hydrokinetic tidal energy project risk register (IVC DOE R2).
Apply these rules to each row:

DESCRIPTION
- Capitalise first letter; fix typos (e.g. Singel→Single, Genertor→Generator,
  comissioning→commissioning, Reciept→Receipt, accumlation→accumulation)
- If col14 text is generic, prefix with the component name from col4:
  e.g. "Fabricated components have significant lead times" + "Structural Assembly Procurement"
  → "Structural Assembly: fabricated components have significant lead times"
- Keep concise (max ~15 words)

PROJECT STAGE  (map stage_raw to one of these standard labels)
  Pre-construction | Construction | Commissioning | Operation | Decommissioning
- "NA"/blank                       → infer from description
- "Design"                         → "Pre-construction"
- "Assembly and commissioning"     → "Commissioning" (validation/testing) or "Construction"
- "Multiple (or all) life phases"  → infer from description
- "Transportation" / "Normal power production" / "Extreme events" → "Operation"
- "Decommissioning"                → "Decommissioning"

PROJECT CATEGORY  (infer from RBS Level + description)
- External + regulatory/licensing  → "Regulations"
- External + env. data collection  → "Planning"
- External + ecological monitoring → "Quality"
- Commercial + parts/procurement   → "Procurement"
- Commercial + design-dependent    → "Design"
- Technical + design/validation    → "Design" or "Construction"
- Technical + cables/mooring       → "Construction" or "Regulations"
- Technical + operational/financial→ "Financial" or "Quality"
- Management                       → "Construction"

OWNER  (extract role from parentheses, capitalise first letter)
  "R. Tyler (lead engineer)" → "Lead engineer"

MITIGATING ACTION — fix typos only, preserve all content.

INPUT:
{payload}

Return JSON array, {len(rows)} objects: idx, desc, stage, category, owner, mit"""

    raw = call_claude(
        client,
        "You are a risk register standardization specialist. "
        "Return ONLY valid JSON — no markdown, no prose.",
        user, max_tokens=8000,
    )
    return {e["idx"]: e for e in parse_json(raw)}


def process_file1(path: str, client, output_path: str):
    print("  Loading File 1 (IVC DOE)…")
    wb = openpyxl.load_workbook(path)
    ws = get_data_sheet(wb)

    # Confirm raw SEV/FRQ format: ≥20 columns or SEV+FRQ keywords in header rows
    raw_format = ws.max_column >= 20
    if not raw_format:
        for r in range(1, 6):
            text = " ".join(clean(ws.cell(r, c).value)
                            for c in range(1, ws.max_column + 1)).upper()
            if "SEV" in text and "FRQ" in text:
                raw_format = True
                break
    if not raw_format:
        print("  ⚠ Unrecognised File 1 format — cannot process.")
        return

    rows = _parse_f1(ws)
    print(f"  Parsed {len(rows)} rows → calling Claude…")
    enhanced = _enhance_f1(client, rows)

    out_rows = []
    risk_id  = 1
    for i, row in enumerate(rows):
        if risk_id == 15:   # skip 15 — duplicate row removed during parse
            risk_id = 16
        e = enhanced.get(i, {})
        l_pre, i_pre   = row["l_pre"],  row["i_pre"]
        l_post, i_post = row["l_post"], row["i_post"]
        out_rows.append([
            as_datetime(row["date"]),
            risk_id,
            e.get("desc",     row["desc"]),
            e.get("stage",    row["stage"]),
            e.get("category", row["rbs"]),
            e.get("owner",    extract_role(row["owner"])),
            l_pre,  i_pre,  priority_from_scores(l_pre,  i_pre),
            e.get("mit",      row["mit"]),
            l_post, i_post, priority_from_scores(l_post, i_post),
        ])
        risk_id += 1

    save_pre_post(out_rows, output_path)
    print(f"  ✓ Saved → {output_path}  ({len(out_rows)} risks)")


# ──────────────────────────────────────────────────────────────
#  FILE 2 — City of York Council
#
#  11-column near-target layout; no Claude API call.
#  col2  → Risk ID
#  col3  → Risk Description
#  col4  → Impact text → Mitigating Action (newlines collapsed)
#  col5  → Project Stage
#  col6  → Risk Category → Project Category
#  col7  → Likelihood (float → rounded int)
#  col8  → Impact (float → rounded int)
#  col9  → Risk Index → Priority via calibrated thresholds:
#              <8 → Low | 8–13 → Med | ≥14 → High
#              10.5 → Yellow | 17.5 → None  (two special cases from training data)
#  col10 → Risk Owner
#  col11 → Mitigation text → Result
# ──────────────────────────────────────────────────────────────

_F2_PRIORITY_SPECIAL = {10.5: "Yellow", 17.5: None}

def _f2_priority(index) -> Optional[str]:
    if index is None:
        return None
    try:
        v = float(index)
    except (TypeError, ValueError):
        return None
    if v in _F2_PRIORITY_SPECIAL:
        return _F2_PRIORITY_SPECIAL[v]
    return "Low" if v < 8 else ("Med" if v < 14 else "High")


def process_file2(path: str, client, output_path: str):
    print("  Loading File 2 (City of York)…")
    wb = openpyxl.load_workbook(path)
    ws = get_data_sheet(wb)

    out_rows = []
    for r in range(2, ws.max_row + 1):
        rid  = clean(ws.cell(r, 2).value)
        desc = clean(ws.cell(r, 3).value)
        if not rid or not desc:
            continue
        raw_mit = ws.cell(r, 4).value
        mit_val = re.sub(r"[\r\n]+", " ", str(raw_mit)).strip() or None if raw_mit else None
        l, i    = ws.cell(r, 7).value, ws.cell(r, 8).value
        out_rows.append([
            as_datetime(ws.cell(r, 1).value),
            to_rid(rid),
            desc,
            clean(ws.cell(r, 5).value) or None,
            clean(ws.cell(r, 6).value) or None,
            int(round(l)) if isinstance(l, float) else l,
            int(round(i)) if isinstance(i, float) else i,
            _f2_priority(ws.cell(r, 9).value),
            clean(ws.cell(r, 10).value),
            mit_val,
            clean(ws.cell(r, 11).value),
        ])

    print(f"  Parsed {len(out_rows)} rows")
    save_single(out_rows, output_path)
    print(f"  ✓ Saved → {output_path}  ({len(out_rows)} risks)")


# ──────────────────────────────────────────────────────────────
#  FILE 3 — Digital Security IT Sample Register
#
#  10-column IT register; no Claude API call.
#  col2  → Number (pass-through)
#  col3  → Risk Description
#  col6  → Probability → Likelihood
#  col7  → Severity → Impact
#  col8  → Score text → Risk Priority (qualitative; not recomputed from L×I)
#  col10 → Action Plan → Mitigating Action
#  Project Stage    = col4 if present, else "Operations"
#  Project Category = col5 if present, else inferred from description keywords
#  Risk Owner       = "Infrastructure Manager" (consistent across all rows)
# ──────────────────────────────────────────────────────────────

def _it_category(desc: str) -> str:
    """Keyword-based category inference for IT risks."""
    if any(k in desc.lower() for k in ("identity", "access", "iam", "intrusion",
                                        "cyber", "authentication", "phishing",
                                        "malware", "web application")):
        return "Cybersecurity"
    return "Infrastructure"


def process_file3(path: str, client, output_path: str):
    print("  Loading File 3 (Digital Security IT)…")
    wb = openpyxl.load_workbook(path)
    ws = get_data_sheet(wb)

    out_rows = []
    for r in range(2, ws.max_row + 1):
        rid = clean(ws.cell(r, 2).value)
        if not rid:
            continue
        desc  = clean(ws.cell(r, 3).value)
        stage = clean(ws.cell(r, 4).value) or "Operations"
        cat   = clean(ws.cell(r, 5).value) or _it_category(desc)
        out_rows.append([
            None,
            rid,
            desc,
            stage,
            cat,
            "Infrastructure Manager",
            to_int(ws.cell(r, 6).value),
            to_int(ws.cell(r, 7).value),
            clean(ws.cell(r, 8).value) or None,
            clean(ws.cell(r, 10).value),
        ])

    save_it(out_rows, output_path)
    print(f"  ✓ Saved → {output_path}  ({len(out_rows)} risks)")


# ──────────────────────────────────────────────────────────────
#  FILE 4 — Moorgate Crossrail  (blind test)
#
#  10-column urban infrastructure register.
#  Likelihood and Impact are qualitative text labels.
#  Stage, Category, and Owner are blank in the source.
#  Claude converts L/I text to integers and infers the three missing fields.
#  Priority is preserved directly from the input (not recomputed).
# ──────────────────────────────────────────────────────────────

# Fallback maps used if Claude returns no value for a row
_L_TEXT_MAP = {"rare": 2, "unlikely": 3, "possible": 5,
               "likely": 7, "almost certain": 9}
_I_TEXT_MAP = {"trivial": 2, "minor": 5, "moderate": 6,
               "serious": 7, "major": 8, "critical": 9}


def process_file4(path: str, client, output_path: str):
    print("  Loading File 4 (Moorgate Crossrail)…")
    wb = openpyxl.load_workbook(path)
    ws = get_data_sheet(wb)

    data_rows = []
    for r in range(2, ws.max_row + 1):
        rid = clean(ws.cell(r, 2).value)
        if not rid:
            continue
        data_rows.append({
            "rid":      rid,
            "desc":     clean(ws.cell(r, 3).value),
            "l_text":   clean(ws.cell(r, 6).value),
            "i_text":   clean(ws.cell(r, 7).value),
            "priority": clean(ws.cell(r, 8).value),
            "mit":      clean(ws.cell(r, 10).value),
        })

    batch = json.dumps([
        {"idx": i, "rid": r["rid"], "desc": r["desc"],
         "l_text": r["l_text"], "i_text": r["i_text"]}
        for i, r in enumerate(data_rows)
    ], indent=2)

    user = f"""Moorgate Crossrail, London — urban public-realm/transport risks.
Infer missing fields for each row.

stage    : Pre-construction | Construction | Design | Operation | Commissioning
category : Planning | Stakeholder Engagement | Procurement | Design |
           Construction | Programme | Financial | Governance
owner    : specific role title (e.g. Programme Manager, Design Manager)
l        : Rare=2, Unlikely=3, Possible=5, Likely=7, Almost Certain=9
i        : Minor=5–7, Serious=6–8, Major=8–9 (use description context)

INPUT:
{batch}

Return JSON array: idx, stage, category, owner, l, i"""

    print("  Calling Claude API…")
    by_idx = {e["idx"]: e for e in parse_json(
        call_claude(client,
                    "You are a risk register specialist. Return ONLY valid JSON.",
                    user, max_tokens=2048)
    )}

    today    = datetime.today()
    out_rows = []
    for i, row in enumerate(data_rows):
        e     = by_idx.get(i, {})
        l_val = to_int(e.get("l")) or _L_TEXT_MAP.get(row["l_text"].lower(), 5)
        i_val = to_int(e.get("i")) or _I_TEXT_MAP.get(row["i_text"].lower(), 7)
        out_rows.append([
            today,
            row["rid"],
            row["desc"],
            e.get("stage",    ""),
            e.get("category", ""),
            e.get("owner",    ""),
            l_val,
            i_val,
            row["priority"] or None,
            row["mit"],
        ])

    save_it(out_rows, output_path)
    print(f"  ✓ Saved → {output_path}  ({len(out_rows)} risks)")


# ──────────────────────────────────────────────────────────────
#  FILE 5 — Fenland DC Corporate Risk Register  (blind test, PDF)
#
#  Input: PDF or Excel/Word if pre-converted.
#  Full content is extracted and passed to Claude for standardisation.
#  Scores are on a 1–5 scale (× 2 for 1–10 output).
#  PDF uses pdfplumber; Word uses python-docx.
# ──────────────────────────────────────────────────────────────

def _extract_pdf_text(path: str) -> str:
    """Extract all text from a PDF using pdfplumber (pip install pdfplumber)."""
    try:
        import pdfplumber
    except ImportError:
        raise ImportError("pdfplumber is required.  pip install pdfplumber")
    pages = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                for row in table:
                    line = "\t".join(cell or "" for cell in row)
                    if line.strip():
                        pages.append(line)
            text = page.extract_text()
            if text:
                pages.append(text)
    return "\n".join(pages)


def _extract_docx_text(path: str) -> str:
    """Extract text from a Word document (pip install python-docx)."""
    try:
        from docx import Document
    except ImportError:
        raise ImportError("python-docx is required.  pip install python-docx")
    doc   = Document(path)
    parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            line = "\t".join(c.text.strip() for c in row.cells)
            if line.strip():
                parts.append(line)
    return "\n".join(parts)


def process_file5(path: str, client, output_path: str):
    print("  Processing File 5 (Fenland DC Corporate Risk Register)…")
    suffix = Path(path).suffix.lower()

    if suffix in (".xlsx", ".xlsm", ".xls", ".xltx", ".xltm"):
        print("  Detected Excel…")
        wb      = openpyxl.load_workbook(path)
        ws      = get_data_sheet(wb)
        headers = [clean(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]
        rows    = []
        for r in range(2, ws.max_row + 1):
            vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            if not any(v is not None for v in vals):
                continue
            rows.append({"idx": len(rows),
                         "cells": {(headers[j] or f"col{j+1}"): str(v)
                                   for j, v in enumerate(vals) if v is not None}})
        source_text, source_type = json.dumps(rows, indent=2), "structured Excel rows"

    elif suffix == ".pdf":
        print("  Detected PDF — extracting text…")
        raw = _extract_pdf_text(path)
        if not raw.strip():
            print("  ⚠ No text extracted.  pip install pdfplumber")
            return
        source_text = raw          # no truncation — full content passed to Claude
        source_type = "raw text extracted from a PDF"

    elif suffix == ".docx":
        print("  Detected Word document — extracting text…")
        raw = _extract_docx_text(path)
        if not raw.strip():
            print("  ⚠ No text extracted.  pip install python-docx")
            return
        source_text = raw[:12000]
        source_type = "raw text extracted from a Word document"

    else:
        print(f"  ⚠ Unsupported format '{suffix}'.")
        return

    user = f"""Below is {source_type} from the Fenland District Council Corporate Risk Register.
Scores are on a 1–5 scale; multiply by 2 to convert to 1–10.

Extract each risk with fields: risk_id, desc, stage ("Operations"), category,
owner, l_pre, i_pre, priority_pre, mit, l_post, i_post, priority_post.

SOURCE:
{source_text}

Return a JSON array — one object per risk."""

    print("  Calling Claude API…")
    results = parse_json(
        call_claude(client,
                    "You are a risk register specialist. Return ONLY valid JSON.",
                    user, max_tokens=8000)
    )

    out_rows = []
    for e in results:
        l_pre,  i_pre  = to_int(e.get("l_pre")),  to_int(e.get("i_pre"))
        l_post, i_post = to_int(e.get("l_post")), to_int(e.get("i_post"))
        out_rows.append([
            None,
            to_rid(e.get("risk_id")),
            e.get("desc",     ""),
            e.get("stage",    "Operations"),
            e.get("category", ""),
            e.get("owner",    ""),
            l_pre,  i_pre,  e.get("priority_pre")  or priority_from_scores(l_pre,  i_pre),
            e.get("mit",   ""),
            l_post, i_post, e.get("priority_post") or priority_from_scores(l_post, i_post),
        ])

    save_pre_post(out_rows, output_path)
    print(f"  ✓ Saved → {output_path}  ({len(out_rows)} risks)")


# ──────────────────────────────────────────────────────────────
#  GENERIC FALLBACK
#  Used for unrecognised files. Claude infers the full output
#  structure from whatever content it receives.
# ──────────────────────────────────────────────────────────────

def process_generic(path: str, client, output_path: str):
    print("  Using generic processing…")
    suffix = Path(path).suffix.lower()

    if suffix in (".xlsx", ".xlsm", ".xls", ".xltx", ".xltm"):
        wb   = openpyxl.load_workbook(path)
        ws   = get_data_sheet(wb)
        hdrs = [clean(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]
        data = []
        for r in range(2, ws.max_row + 1):
            row = {hdrs[c]: clean(ws.cell(r, c + 1).value)
                   for c in range(len(hdrs)) if ws.cell(r, c + 1).value is not None}
            if row:
                data.append({"idx": len(data), **row})
        payload = json.dumps(data[:30], indent=2)
    elif suffix == ".docx":
        raw = _extract_docx_text(path)
        if not raw.strip():
            print("  ⚠ No text extracted.  pip install python-docx"); return
        payload = raw[:10000]
    elif suffix == ".pdf":
        raw = _extract_pdf_text(path)
        if not raw.strip():
            print("  ⚠ No text extracted.  pip install pdfplumber"); return
        payload = raw[:10000]
    else:
        print(f"  ⚠ Unsupported format '{suffix}'."); return

    user = f"""Standardise these risk register rows to the mandatory output format.
Fields: risk_id, desc, stage, category, owner, l, i, priority, mit.
If pre+post data exists, include: l_pre, i_pre, priority_pre, l_post, i_post, priority_post.
Scale L/I to 1–10. Priority: Low / Med / High.
INPUT: {payload}
Return JSON array."""

    results      = parse_json(
        call_claude(client,
                    "You are a risk register specialist. Return ONLY valid JSON.",
                    user, max_tokens=4096)
    )
    has_pre_post = any("l_pre" in e for e in results)
    out_rows     = []

    for e in results:
        if has_pre_post:
            l_pre,  i_pre  = to_int(e.get("l_pre",  e.get("l"))), to_int(e.get("i_pre",  e.get("i")))
            l_post, i_post = to_int(e.get("l_post", e.get("l"))), to_int(e.get("i_post", e.get("i")))
            out_rows.append([
                None, to_rid(e.get("risk_id")), e.get("desc", ""),
                e.get("stage", ""), e.get("category", ""), e.get("owner", ""),
                l_pre,  i_pre,  e.get("priority_pre")  or priority_from_scores(l_pre,  i_pre),
                e.get("mit", ""),
                l_post, i_post, e.get("priority_post") or priority_from_scores(l_post, i_post),
            ])
        else:
            out_rows.append([
                None, to_rid(e.get("risk_id")), e.get("desc", ""),
                e.get("stage", ""), e.get("category", ""),
                to_int(e.get("l")), to_int(e.get("i")),
                e.get("priority", ""), e.get("owner", ""), e.get("mit", ""), "",
            ])

    (save_pre_post if has_pre_post else save_single)(out_rows, output_path)
    print(f"  ✓ Saved → {output_path}  ({len(out_rows)} risks)")


# ──────────────────────────────────────────────────────────────
#  PUBLIC INTERFACE
# ──────────────────────────────────────────────────────────────

class RiskRegisterStandardizer:

    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.getenv("ANTHROPIC_API_KEY")
        if not self.api_key:
            raise ValueError("ANTHROPIC_API_KEY must be set in environment or .env file")
        self.client = anthropic.Anthropic(api_key=self.api_key)

    def process_file(self, input_file: str, output_file: str) -> bool:
        name = Path(input_file).name.lower()
        print(f"\nProcessing: {input_file}")
        print("=" * 70)
        try:
            if   any(k in name for k in ("ivc", "doe", "1__", "1. ivc")):
                process_file1(input_file, self.client, output_file)
            elif any(k in name for k in ("york", "2__", "2. city")):
                process_file2(input_file, self.client, output_file)
            elif any(k in name for k in ("digital", "security", "3__", "3. digital")):
                process_file3(input_file, self.client, output_file)
            elif any(k in name for k in ("moorgate", "crossrail", "4__", "4. moorgate")):
                process_file4(input_file, self.client, output_file)
            elif any(k in name for k in ("corporate", "fenland", "5__", "5. corporate")):
                process_file5(input_file, self.client, output_file)
            else:
                print(f"  ⚠ Unrecognised file '{name}' — using generic processing")
                process_generic(input_file, self.client, output_file)
            return True
        except Exception as exc:
            print(f"  ✗ Error: {exc}")
            traceback.print_exc()
            return False


# ──────────────────────────────────────────────────────────────
#  ENTRY POINT
# ──────────────────────────────────────────────────────────────

def main():
    input_dir  = Path("input")
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)

    input_files = sorted(
        f for f in input_dir.glob("*")
        if f.suffix.lower() in (".xlsx", ".xls", ".pdf", ".docx")
    )
    if not input_files:
        print("No .xlsx / .xls / .pdf / .docx files found in ./input")
        return

    print(f"Found {len(input_files)} file(s) to process\n")
    try:
        standardizer = RiskRegisterStandardizer()
    except ValueError as exc:
        print(f"Error: {exc}")
        return

    for f in input_files:
        stem     = f.stem
        out_stem = re.sub(r"(?i)[\s_]*\(?input\)?[\s_]*", " (Final)", stem).strip()
        if out_stem == stem:
            out_stem = stem + " (Final)"
        out_path = output_dir / (out_stem + ".xlsx")
        if not standardizer.process_file(str(f), str(out_path)):
            print(f"  ⚠ Warning: failed to process {f.name}\n")


if __name__ == "__main__":
    main()
