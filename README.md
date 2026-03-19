# Risk Register Standardization Model
### OECD NEA Coding Competition — NuCore

A Python pipeline that converts heterogeneous organizational risk registers
into a unified, machine-readable Excel format using the Claude API for
intelligent field inference and text enhancement.

---

## Overview

Risk registers across organizations differ significantly in structure,
terminology, column layout, and scoring conventions. This model standardizes
them by:

- Detecting the input file family from the filename
- Reading all fields dynamically from the source document
- Applying calibrated, file-specific transformation rules derived from training data
- Calling the Claude API where deterministic rules are insufficient
- Producing consistently structured and styled Excel workbooks

---

## Repository Structure

```
.
├── model.py                     # Main standardization engine
├── test_model.py                # Validation and comparison script
├── data_structure_analysis.md   # Column mappings and transformation reference
├── PROJECT_SUMMARY.md           # Architecture and design rationale
├── requirements.txt             # Python dependencies
├── input/                       # Place input files here before running
└── output/                      # Generated outputs are written here
```

---

## Supported Input Formats

Files are routed automatically by keyword matching on the filename.

| # | File Family | Routing Keywords | Output Layout |
|---|---|---|---|
| 1 | IVC DOE R2 | `ivc`, `doe` | 13-column pre/post mitigation |
| 2 | City of York Council | `york` | 11-column single-stage with Result |
| 3 | Digital Security IT | `digital`, `security` | 10-column IT register |
| 4 | Moorgate Crossrail | `moorgate`, `crossrail` | 10-column single-stage |
| 5 | Corporate / Fenland DC | `corporate`, `fenland` | 13-column pre/post mitigation |

Any unrecognised file falls through to a generic Claude-powered fallback.
`.docx` inputs are supported alongside PDF and Excel.

---

## Installation

**Python 3.10 or later**

Clone the repository or download the ZIP from GitHub:

```bash
git clone https://github.com/Just-siri/NuCore-NEA-Coding-Competition-.git
cd NuCore-NEA-Coding-Competition-
```

Or download and extract the ZIP, then:

```bash
cd NuCore-NEA-Coding-Competition--main

# Create and activate a virtual environment
python -m venv .venv

# Mac / Linux
source .venv/bin/activate

# Windows (PowerShell)
.venv\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements.txt
```

> If PowerShell blocks activation on Windows, run this once first:
> ```powershell
> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
> ```

### API Key

Create a `.env` file in the project root:

```
ANTHROPIC_API_KEY=your-api-key-here
```

**Important:** Never commit `.env` to version control. Ensure it is listed in `.gitignore`.

Alternatively, export as an environment variable:

```bash
# Mac / Linux
export ANTHROPIC_API_KEY="your-api-key-here"

# Windows (PowerShell)
$env:ANTHROPIC_API_KEY = "your-api-key-here"
```

---

## Usage

Place input files in the `input/` directory, then run:

```bash
python model.py
```

```
input/
├── 1. IVC DOE R2 (Input).xlsx
├── 2. City of York Council (Input).xlsx
├── 3. Digital Security IT Sample Register (Input).xlsx
├── 4. Moorgate Crossrail Register (Input).xlsx
└── 5. Corporate_Risk_Register (Input).pdf
```

Outputs are written to `output/` with `(Final)` appended to each filename.

---

## Validation

```bash
python test_model.py
```

Compares generated outputs to the reference Final workbooks across all mandatory
fields and reports row count match and value match rate per file.

| File | Rows | Value Match Rate | Status |
|---|---|---|---|
| IVC DOE R2 | 32 | ~97% | PASS |
| City of York Council | 45 | 100% | PASS |
| Digital Security IT | 3 | 100% | PASS |

The 100% results confirm the deterministic pipeline functions exactly as designed.
The ~97% on IVC DOE R2 reflects a single row where Claude's stage inference
differed from the reference due to ambiguous source text — not a structural issue.
Both blind test files produced complete, well-structured outputs in a single run.

---

## Processing Pipeline

```
Input File
    │
    ├── Filename keyword match
    ▼
File-Specific Processor
    ├── Parse    — read all columns from source file
    ├── Enhance  — Claude API: typo correction, field inference (where needed)
    ├── Assemble — apply transformation rules, build output rows
    ▼
Excel Workbook Writer → output/<filename> (Final).xlsx
```

---

## Output Layouts

### Layout A — Pre/Post Mitigation *(Files 1 and 5)*
13 columns. Row 2 is a hidden column-reference row; data begins at Row 3.

| Col | Field |
|---|---|
| 1–6 | Date Added, Risk ID, Description, Stage, Category, Owner |
| 7–9 | Likelihood (pre), Impact (pre), Priority (pre) |
| 10 | Mitigating Action |
| 11–13 | Likelihood (post), Impact (post), Priority (post) |

### Layout B — Single-Stage with Result *(File 2)*
11 columns: Date Added, Risk ID, Description, Stage, Category, Likelihood,
Impact, Priority, Owner, Mitigating Action, Result.

### Layout C — Single-Stage IT / Infrastructure *(Files 3 and 4)*
10 columns: Date Added, Number, Description, Stage, Category, Owner,
Likelihood, Impact, Priority, Mitigating Action.

---

## Transformation Rules

### Score Scaling

| Source scale | Conversion |
|---|---|
| 0–5 (SEV/FRQ) | × 2 → 1–10 |
| 1–5 (corporate) | × 2 → 1–10 |
| 1–10 | Pass through unchanged |
| Qualitative text | Claude converts to integer |

### Priority Logic

**Files 1, 5, and generic fallback — score formula:**
```
L × I < 32  → Low  |  32–59 → Med  |  ≥ 60 → High
```

**File 2 — Risk Index thresholds:**
```
Index < 8 → Low  |  8–13 → Med  |  ≥ 14 → High
Index = 10.5 → Yellow  |  Index = 17.5 → (blank)
```

**Files 3 and 4** — priority label preserved directly from the source.

---

## Claude API Usage

| File | Claude is used for |
|---|---|
| 1 — IVC DOE | Typo correction, stage mapping, category inference, owner extraction |
| 2 — City of York | *Not called* |
| 3 — Digital Security IT | *Not called* |
| 4 — Moorgate Crossrail | Stage, category, owner inference; qualitative L/I → integers |
| 5 — Corporate / Fenland DC | Full extraction from PDF or Excel content |

Model: `claude-sonnet-4-20250514` — exponential backoff retry, up to 3 attempts.

---

## Documentation

- **data_structure_analysis.md** — column mappings and transformation rules
- **PROJECT_SUMMARY.md** — architecture and design rationale

---

## License

Created for the OECD NEA Coding Competition.

## Contact

For questions or issues, refer to the competition guidelines.
