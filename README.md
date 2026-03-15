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

## Mandatory Output Fields

Every generated workbook includes the following fields regardless of input format:

| # | Field | Notes |
|---|---|---|
| 1 | Risk ID | Preserved from input or generated sequentially |
| 2 | Risk Description | Full text, word-wrapped |
| 3 | Project Stage | Standardised stage label |
| 4 | Project Category | Inferred or mapped from source |
| 5 | Risk Owner | Role title |
| 6 | Mitigating Action | Full text |
| 7 | Likelihood (1-10) | Integer, scaled to 1–10 |
| 8 | Impact (1-10) | Integer, scaled to 1–10 |
| 9 | Risk Priority (low, med, high) | Computed or preserved from source |

Where a file contains pre- and post-mitigation data, Likelihood, Impact,
and Risk Priority are included twice — once for each stage.

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

Any unrecognised file is handled by a generic fallback processor that uses
Claude to infer the appropriate output structure automatically.

---

## Installation

**Python 3.10 or later**

```bash
# Install required packages
pip install -r requirements.txt

# For PDF input support (File 5), also install:
pip install pdfminer.six
```

The model also accepts `pypdf` or `PyPDF2` if either is already installed
and selects whichever is available automatically.

```bash
# Set API key — create a .env file in the project root and add:
ANTHROPIC_API_KEY=your-api-key-here
```

**Important:** Never commit the `.env` file to version control.
Ensure `.gitignore` contains a `.env` entry before pushing to GitHub.

Alternatively, export the key as an environment variable:

**Linux / macOS**
```bash
export ANTHROPIC_API_KEY="your-api-key-here"
```

**Windows (PowerShell)**
```powershell
$env:ANTHROPIC_API_KEY = "your-api-key-here"
```

---

## Usage

### Setup

```bash
# Mac / Linux
mkdir -p input output

# Windows
mkdir input output
```

Place all input files inside the `input/` directory, then run:

```bash
python model.py
```

**Example input directory:**
```
input/
├── 1. IVC DOE R2 (Input).xlsx
├── 2. City of York Council (Input).xlsx
├── 3. Digital Security IT Sample Register (Input).xlsx
├── 4. Moorgate Crossrail Register (Input).xlsx
└── 5. Corporate_Risk_Register (Input).pdf
```

Outputs are written to `output/` with `(Final)` appended to each filename.
The `output/` directory is created automatically if it does not exist.

---

## Running the Tests

```bash
python test_model.py
```

The test script:

1. Processes the three training input files using the model
2. Locates the corresponding reference Final workbooks automatically
3. Compares row counts and all mandatory field values
4. Reports pass/fail status and a value match rate for each file

---

## How It Works

### 1. File Detection and Extraction

- File type and family are detected from the filename
- Excel files are read with `openpyxl`; PDF files have text extracted using
  `pdfminer.six` (with `pypdf` and `PyPDF2` as automatic fallbacks)
- All source columns are read dynamically — no column positions are assumed

### 2. Claude API Enhancement

Where deterministic rules are not sufficient, a single batched API call
is made per file. Claude performs:

- Typo correction in descriptions and mitigating action text
- Stage label mapping from source-specific vocabulary to standard terms
- Project category inference from risk context and RBS classification
- Owner name extraction and normalisation
- Qualitative likelihood/impact text conversion to integer scores
- Full extraction and structuring of PDF register content

### 3. Transformation

File-specific rules are applied to build the output rows:

- Scores scaled to 1–10 (e.g. 0–5 scale multiplied by 2)
- Likelihood/Impact floats rounded to integers where required
- Risk Priority computed from the L × I score formula, or derived from
  a calibrated threshold applied to a source index field

### 4. Output Generation

Each workbook is saved as a fully formatted `.xlsx` file containing a
**Simplified Register** sheet and an **Output Requirements** sheet.

---

## Processing Pipeline

```
Input File
    │
    ├── Filename keyword match
    │
    ▼
File-Specific Processor
    │
    ├── Parse    — read all columns from source file
    ├── Enhance  — Claude API: typo correction, field inference,
    │              vocabulary standardisation (where needed)
    ├── Assemble — apply transformation rules, build output rows
    │
    ▼
Excel Workbook Writer
    │
    ├── Styled header row
    ├── Hidden reference row (Layout A only)
    ├── Integer-formatted score columns
    ├── Date-formatted date columns
    ├── Output Requirements sheet
    │
    ▼
output/<filename> (Final).xlsx
```

---

## Output Layouts

Each workbook contains two sheets: **Simplified Register** and
**Output Requirements**.

### Layout A — Pre/Post Mitigation  *(Files 1 and 5)*

13 columns. Data begins at Row 3; Row 2 is a hidden column-reference row.

| Col | Field |
|---|---|
| 1 | Date Added |
| 2 | Risk ID |
| 3 | Risk Description |
| 4 | Project Stage |
| 5 | Project Category |
| 6 | Risk Owner |
| 7 | Likelihood (1-10) (pre-mitigation) |
| 8 | Impact (1-10) (pre-mitigation) |
| 9 | Risk Priority (pre-mitigation) |
| 10 | Mitigating Action |
| 11 | Likelihood (1-10) (post-mitigation) |
| 12 | Impact (1-10) (post-mitigation) |
| 13 | Risk Priority (post-mitigation) |

### Layout B — Single-Stage with Result  *(File 2)*

11 columns. Data begins at Row 2.

| Col | Field |
|---|---|
| 1 | Date Added |
| 2 | Risk ID |
| 3 | Risk Description |
| 4 | Project Stage |
| 5 | Project Category |
| 6 | Likelihood (1-10) |
| 7 | Impact (1-10) |
| 8 | Risk Priority (low, med, high) |
| 9 | Risk Owner |
| 10 | Mitigating Action |
| 11 | Result |

### Layout C — Single-Stage IT / Infrastructure  *(Files 3 and 4)*

10 columns. Data begins at Row 2.

| Col | Field |
|---|---|
| 1 | Date Added |
| 2 | Number / Risk ID |
| 3 | Risk Description |
| 4 | Project Stage |
| 5 | Project Category |
| 6 | Risk Owner |
| 7 | Likelihood (1-10) |
| 8 | Impact (1-10) |
| 9 | Risk Priority (low, med, high) |
| 10 | Mitigating Action |

---

## Transformation Rules

### Column Mapping Examples

| Source column name | Output field |
|---|---|
| `Risk Category` | Project Category |
| `Mitigation` / `Action Plan` | Mitigating Action |
| `Number` / `ID` | Risk ID |
| `Probability` / `FRQ` | Likelihood (1-10) |
| `Severity` / `SEV` | Impact (1-10) |
| `Project Stage` | Project Stage |

### Score Scaling

| Source scale | Conversion |
|---|---|
| 0–5 (SEV/FRQ) | Multiply by 2 → 1–10 |
| 1–5 (corporate) | Multiply by 2 → 1–10 |
| 1–10 | Pass through unchanged |
| Qualitative text | Claude converts to integer |

### Risk Priority Logic

**Default — score formula** *(Files 1, 5, and generic fallback)*

```
Risk Score = Likelihood × Impact

Score  < 32   →  Low
Score 32–59   →  Med
Score ≥ 60    →  High
```

**File 2 — Risk Index thresholds** *(calibrated from training data)*

```
Index  < 8    →  Low
Index 8–13    →  Med
Index ≥ 14    →  High
Index = 10.5  →  Yellow   (special case)
Index = 17.5  →  (blank)  (unresolvable)
```

**Files 3 and 4** — Priority text preserved directly from the source file.

---

## Claude API Usage

The API is called selectively — only where deterministic rules are not
sufficient to produce the required output.

| File | Claude is used for |
|---|---|
| 1 — IVC DOE | Typo correction, stage mapping, category inference, owner name cleaning |
| 2 — City of York | *Not called — all fields pass through deterministically* |
| 3 — Digital Security IT | *Not called — all fields pass through deterministically* |
| 4 — Moorgate Crossrail | Stage, category, and owner inference; qualitative L/I text → integers |
| 5 — Corporate / Fenland DC | Full extraction from PDF text or Excel rows |

API calls use `claude-sonnet-4-20250514` with exponential backoff retry
(up to 3 attempts) on transient errors. All rows are batched into a single
request per file.

---

## Error Handling

- **PDF library not installed** — clear error message with install command
- **API failure** — exponential backoff with up to 3 retries, then raises
- **Unrecognised file format** — falls back to generic Claude-powered processor
- **Malformed JSON from API** — tolerant parser strips markdown fences
  and irregular whitespace before parsing
- **Missing input file** — caught cleanly in the test script with a clear message

---

## Performance

| Factor | Typical range |
|---|---|
| Processing time per file | 10–60 seconds (depends on API latency) |
| Token usage per file | ~2,000–8,000 tokens |
| Files without API call | Files 2 and 3 (no API needed) |

---

## Documentation

- **data_structure_analysis.md** — detailed column mapping tables and
  transformation rules for every file family
- **PROJECT_SUMMARY.md** — model architecture, design rationale, and
  file-by-file processing logic

---

## Testing

The model has been validated against three training input–output pairs:

1. IVC DOE R2 — complex 24-column SEV/FRQ register
2. City of York Council — near-target 11-column register
3. Digital Security IT Sample Register — compact IT register

Run:

```bash
python test_model.py
```

---

## License

This model was created for the OECD NEA Coding Competition.

---

## Contact

For questions or issues, refer to the competition guidelines.
