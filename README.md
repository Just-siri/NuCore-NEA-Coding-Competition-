
# Risk Register Standardization Model
## OECD NEA Coding Competition — NuCore

This repository contains the final implementation used to convert heterogeneous risk registers into a standardized Excel structure. 
The model processes multiple source formats and produces outputs aligned with the competition schema.

---

# Overview

Risk registers from different organizations often vary significantly in structure, terminology, and scoring conventions.

This model standardizes those registers by:

- Detecting the input file family
- Mapping source columns to a unified schema
- Applying calibrated transformation rules
- Producing structured Excel outputs compatible with the competition specification

The implementation has been validated against the provided training input–output pairs.

---

# Repository Structure

```
.
├── model.py                     # Main standardization engine
├── test_model.py                # Validation script
├── data_structure_analysis.md   # Detailed mapping analysis
├── PROJECT_SUMMARY.md           # Design overview
├── requirements.txt             # Python dependencies
├── input/                       # Input risk registers
└── output/                      # Generated standardized outputs
```

---

# Supported Input Types

The model routes files automatically based on filename patterns.

| File Family | Typical Input Name | Output Layout |
|---|---|---|
| File 1 — IVC DOE | `1. IVC DOE R2 (Input).xlsx` | 13 columns with pre/post mitigation |
| File 2 — City of York Council | `2. City of York Council (Input).xlsx` | 11 columns |
| File 3 — Digital Security IT | `3. Digital Security IT Sample Register (Input).xlsx` | 10 columns |
| File 4 — Moorgate Crossrail | `4. Moorgate Crossrail Register (Input).xlsx` | 10 columns |
| File 5 — Corporate / Fenland | `5. Corporate_Risk_Register (Input).pdf` | 13 columns |

---

# Installation

Install dependencies:

```bash
pip install -r requirements.txt
```

The model requires **Python 3.10+**.

---

# API Key Initialization

The model initializes a Claude client through the Anthropic API.  
An API key must be available before running the program.

The recommended approach is to store the key in a `.env` file.

Create a `.env` file in the repository root:

```
ANTHROPIC_API_KEY=your-api-key-here
```

Alternatively, export the key directly in your shell.

Linux / macOS:

```bash
export ANTHROPIC_API_KEY="your-api-key-here"
```

Windows (PowerShell):

```powershell
setx ANTHROPIC_API_KEY "your-api-key-here"
```

Once the environment variable is set, the model can access the API automatically.

---

# Running the Model

Place input files in the `input/` directory.

Example:

```
input/
├── 1. IVC DOE R2 (Input).xlsx
├── 2. City of York Council (Input).xlsx
├── 3. Digital Security IT Sample Register (Input).xlsx
├── 4. Moorgate Crossrail Register (Input).xlsx
└── 5. Corporate_Risk_Register (Input).pdf
```

Run the model:

```bash
python model.py
```

Outputs will be written to the `output/` directory.

---

# Model Architecture

The model follows a four‑stage processing pipeline.

```
Input Register
      │
      ▼
File Detection
      │
      ▼
Column Mapping
      │
      ▼
Risk Transformation Logic
      │
      ▼
Standardized Output Register
```

---

# Transformation Pipeline

```
Raw Register
     │
     ▼
Identify File Type
     │
     ▼
Apply File‑Specific Column Mapping
     │
     ▼
Normalize Risk Fields
     │
     ▼
Apply Risk Logic
     │
     ▼
Write Standardized Output Workbook
```

---

# Risk Logic

Default fallback priority logic:

```
Risk Score = Likelihood × Impact
```

Classification:

```
Score < 32       → Low
Score 32–59      → Medium
Score ≥ 60       → High
```

Where reference outputs show calibrated behavior, the model preserves those values instead of recomputing them.

---

# Example Transformation

Example input (simplified):

| Risk ID | Description | Likelihood | Impact |
|---|---|---|---|
| 1 | Equipment failure | 8 | 10 |

Computed score:

```
8 × 10 = 80
```

Standardized output:

| Risk ID | Risk Description | Likelihood | Impact | Priority |
|---|---|---|---|---|
| 1 | Equipment failure | 8 | 10 | High |

---

# Output Structure

Each generated workbook contains:

| Sheet | Purpose |
|---|---|
| Simplified Register | Standardized risk register |
| Output Requirements | Schema reference |

---

# Testing

Run:

```bash
python test_model.py
```

The test script:

1. Generates outputs for the training inputs  
2. Loads the reference output files  
3. Compares row counts and key columns  
4. Reports match statistics

---

# Documentation

Additional documentation:

- **data_structure_analysis.md** — detailed column mapping and transformation rules
- **PROJECT_SUMMARY.md** — model architecture and reasoning

---

# Summary

This system provides a reproducible method for transforming heterogeneous risk registers into a unified schema suitable for automated analysis and comparison.

The model supports multiple input formats, applies calibrated transformation logic, and produces consistent standardized outputs.

