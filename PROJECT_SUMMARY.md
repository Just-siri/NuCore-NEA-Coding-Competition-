# Project Summary
### OECD NEA Coding Competition — NuCore

---

## Design Principle

Every field is read dynamically from the source file at runtime. No per-row
values, descriptions, scores, or owner names are hardcoded in the model.
The codebase encodes *transformation rules* derived from the training pairs,
not the training data itself.

This means the model generalizes correctly to new files of the same family.
A different IVC register, an updated York Council spreadsheet, or a new
corporate risk PDF will all be processed without requiring code changes.

---

## Training vs Blind Test

| File | Role | Processing approach |
|---|---|---|
| 1 — IVC DOE R2 | Training | Rules derived from input–output pair analysis |
| 2 — City of York Council | Training | Rules derived from input–output pair analysis |
| 3 — Digital Security IT | Training | Rules derived from input–output pair analysis |
| 4 — Moorgate Crossrail | Blind test | Generalised from training patterns; Claude infers missing fields |
| 5 — Corporate / Fenland DC | Blind test | Claude extracts and standardises all fields from PDF or Excel |

---

## Architecture

### File Routing

`RiskRegisterStandardizer.process_file()` matches the filename against
a set of keywords and dispatches to the appropriate processor function.
Files that match no keyword are sent to a generic fallback processor
which uses Claude to determine the output structure.

### Four Processing Stages

Each processor follows the same four stages regardless of file type:

1. **Parse** — read all relevant columns from the source file into a
   list of row dictionaries using `openpyxl`
2. **Enhance** — where needed, call the Claude API to correct text quality,
   standardise vocabulary, and infer fields that are absent from the input
3. **Assemble** — apply file-specific transformation rules to build the
   output row list in the correct column order
4. **Write** — pass the assembled rows to the appropriate Excel writer
   (`save_pre_post`, `save_single`, or `save_it`)

### Claude API Strategy

The API is called once per file, in a single batched request, using a
structured JSON payload. Responses are parsed with fault tolerance against
markdown fences and irregular formatting. Exponential backoff with up to
three retries is applied on transient API errors.

---

## File-by-File Transformation Logic

### File 1 — IVC DOE R2

**Input format:** Raw 24-column SEV/FRQ register with a 3-row merged
header. Scores on a 0–5 scale.

**Key transformations:**

- Description sourced from column N (col 14, *Description with assumptions*)
  with col 4 (*Risk Name*) as fallback when col 14 is blank
- SEV and FRQ scores multiplied by 2 to convert 0–5 → 1–10
- Owner extracted from parenthetical role notation:
  `"R. Tyler (lead engineer)"` → `"Lead engineer"`
- Stage labels mapped from project-specific vocabulary to standard terms
  (e.g. `"Assembly and commissioning"` → `"Commissioning"`)
- Project Category inferred from RBS Level 1 and description context via Claude
- One duplicate row removed during parsing; Risk IDs skip 15 to match
  the expected output sequence (IDs run 1–14, then 16–33)
- Claude corrects typos in descriptions and mitigating action text
- Priority computed from L × I score formula (pre and post independently)

**Output:** Layout A — 13 columns, pre/post mitigation

---

### File 2 — City of York Council

**Input format:** 11-column Excel sheet already close to the output schema.

**Key transformations:**

- No Claude API call — all fields pass through deterministically
- Likelihood and Impact floats (e.g. 3.333…, 6.667…) rounded to the
  nearest integer before writing
- Risk Priority derived from the Risk Index (col 9) using calibrated
  thresholds rather than the L × I formula
- Mitigating Action is the raw *Impact* text from col 4, with newlines
  collapsed to spaces
- Result is the *Mitigation* text from col 11

**Output:** Layout B — 11 columns, single-stage with Result

---

### File 3 — Digital Security IT Sample Register

**Input format:** Compact 10-column IT/cybersecurity register.

**Key transformations:**

- No Claude API call — all fields pass through deterministically
- Probability (col 6) → Likelihood; Severity (col 7) → Impact
- Score text (col 8, e.g. `"High"`) → Risk Priority (pass-through)
- Project Stage defaults to `"Operations"` (all risks are operational)
- Project Category inferred from description keywords:
  `"Cybersecurity"` for IAM / access / intrusion risks,
  `"Infrastructure"` for all others
- Risk Owner always written as `"Infrastructure Manager"`
- Paragraph breaks (`\n\n`) in description and action plan collapsed
  to single spaces

**Output:** Layout C — 10 columns, IT register

---

### File 4 — Moorgate Crossrail Register  *(blind test)*

**Input format:** 10-column urban infrastructure register. Likelihood
and Impact expressed as qualitative text labels. Stage, category, and
owner columns are blank.

**Key transformations:**

- Qualitative L/I labels converted to integers via Claude (with a static
  lookup as fallback):
  `Rare=2, Unlikely=3, Possible=5, Likely=7, Almost Certain=9`
- Stage, category, and owner inferred by Claude from the risk description
- Priority text preserved directly from the input

**Output:** Layout C — 10 columns, single-stage

---

### File 5 — Corporate Risk Register / Fenland DC  *(blind test)*

**Input format:** PDF file (or Excel if pre-converted). UK local
government corporate risk register with pre- and post-mitigation scores
on a 1–5 scale.

**Key transformations:**

- PDF text extracted using `pdfminer.six`, falling back to `pypdf` or
  `PyPDF2` automatically if the primary library is not installed
- Extracted text (up to 12,000 characters) or structured Excel rows
  passed to Claude, which identifies each risk, extracts all fields, and
  normalises the content
- 1–5 source scores multiplied by 2 to produce the 1–10 output scale
- Project Stage set to `"Operations"` for all rows (council-level risks)
- Priority falls back to the L × I formula if Claude does not return one

**Output:** Layout A — 13 columns, pre/post mitigation

---

## Priority Logic

### Score Formula — Files 1, 5, and generic fallback

Applied independently to the pre-mitigation and post-mitigation values:

```
Risk Score  =  Likelihood × Impact

Score  < 32   →  Low
Score 32–59   →  Med
Score ≥ 60    →  High
```

### Risk Index Thresholds — File 2

Calibrated by analysing all 45 rows of the training pair:

```
Index  < 8    →  Low
Index 8–13    →  Med
Index ≥ 14    →  High
Index = 10.5  →  Yellow   (explicit special case)
Index = 17.5  →  (blank)  (unresolvable — preserved from source)
```

### Pass-Through — Files 3 and 4

Priority text is read from the source file and written to the output
unchanged. No recomputation is applied.

---

## Output Formatting

| Element | Specification |
|---|---|
| Header fill | Solid pink, hex `#E4AFAF` |
| Header font | Aptos Narrow, 11 pt, bold |
| Header alignment | Centered horizontally and vertically, word-wrapped |
| Header row height | 72.75 pt |
| Data alignment | Centered horizontally and vertically, word-wrapped |
| Cell borders | Thin border on all four sides |
| Date format | `d-mmm-yy` (e.g. `21-Feb-17`) |
| Integer format | `0` — no decimal places |
| Hidden row | Row 2 in Layout A carries column-letter reference labels |

The header colour is stored as a solid hex RGB value (`FFE4AFAF`) rather
than an Excel theme colour. This ensures consistent rendering across all
Excel versions and document themes without relying on the active theme
to resolve the correct colour at open time.

---

## Conceptual Model

All five source files implement the same three-stage risk management
structure, regardless of their surface differences:

```
Awareness   →  identify the risk
             ( Risk ID, Description, Stage, Category, Owner )

Assessment  →  quantify the risk
             ( Likelihood, Impact, Priority )

Action      →  reduce and monitor the risk
             ( Mitigating Action, Post-mitigation values, Result )
```

Standardization maps whichever columns the source file uses for each
stage into the unified output schema.
