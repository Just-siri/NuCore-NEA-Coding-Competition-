# Data Structure Analysis
### OECD NEA Coding Competition — NuCore

Detailed column mappings, transformation rules, and output specifications
for each file family processed by `model.py`.

---

## Conceptual Framework

All five source files implement the same three-stage risk management
structure under different column arrangements:

```
Awareness   →  Risk ID, Description, Project Stage, Category, Owner
Assessment  →  Likelihood, Impact, Risk Priority
Action      →  Mitigating Action, Result, Post-mitigation values
```

Standardization maps whichever source columns serve each role into the
unified output schema.

---

## Output Layouts

### Layout A — Pre/Post Mitigation
*Used by Files 1 and 5.*

13 columns. Row 1 is the styled header. Row 2 is a hidden column-reference
row (labels: A, —, N, G, —, H, L, K, M, P, L, K, M). Data begins at Row 3.

| Col | Field | Type |
|---|---|---|
| 1 | Date Added | Datetime (`d-mmm-yy`) |
| 2 | Risk ID | Integer |
| 3 | Risk Description | Text |
| 4 | Project Stage | Text |
| 5 | Project Category | Text |
| 6 | Risk Owner | Text |
| 7 | Likelihood (1-10) (pre-mitigation) | Integer 1–10 |
| 8 | Impact (1-10) (pre-mitigation) | Integer 1–10 |
| 9 | Risk Priority (pre-mitigation) | Low / Med / High |
| 10 | Mitigating Action | Text |
| 11 | Likelihood (1-10) (post-mitigation) | Integer 1–10 |
| 12 | Impact (1-10) (post-mitigation) | Integer 1–10 |
| 13 | Risk Priority (post-mitigation) | Low / Med / High |

---

### Layout B — Single-Stage with Result
*Used by File 2.*

11 columns. Data begins at Row 2.

| Col | Field | Type |
|---|---|---|
| 1 | Date Added | Datetime (often blank) |
| 2 | Risk ID | Integer |
| 3 | Risk Description | Text |
| 4 | Project Stage | Text |
| 5 | Project Category | Text |
| 6 | Likelihood (1-10) | Integer |
| 7 | Impact (1-10) | Integer |
| 8 | Risk Priority (low, med, high) | Low / Med / High / Yellow |
| 9 | Risk Owner | Text |
| 10 | Mitigating Action | Text |
| 11 | Result | Text |

---

### Layout C — Single-Stage IT / Infrastructure
*Used by Files 3 and 4.*

10 columns. Data begins at Row 2.

| Col | Field | Type |
|---|---|---|
| 1 | Date Added | Datetime (File 3: blank; File 4: today) |
| 2 | Number / Risk ID | Text or integer |
| 3 | Risk Description | Text |
| 4 | Project Stage | Text |
| 5 | Project Category | Text |
| 6 | Risk Owner | Text |
| 7 | Likelihood (1-10) | Integer |
| 8 | Impact (1-10) | Integer |
| 9 | Risk Priority (low, med, high) | Text |
| 10 | Mitigating Action | Text |

---

## File 1 — IVC DOE R2

### Source Characteristics

Raw 24-column register with a 3-row merged header. Scores use a 0–5 scale;
multiplied by 2 to produce the 1–10 output scale. Contains pre- and
post-mitigation values as separate column groups.

### Column Mapping

| Source Col | Source Header | Transformation | Output Field |
|---|---|---|---|
| col 1 | Revision Date | `as_datetime()` | Date Added |
| col 2 | RBS Level 1 | Context for Claude category inference | → Project Category |
| col 4 | Risk Name | Fallback if col 14 is blank | Risk Description (fallback) |
| col 7 | Technology Life Phase | Mapped to standard stage labels by Claude | Project Stage |
| col 8 | Risk Owner | Role extracted from parentheses | Risk Owner |
| col 11 | SEV baseline | `value × 2` | Impact (pre-mitigation) |
| col 12 | FRQ baseline | `value × 2` | Likelihood (pre-mitigation) |
| col 14 | Description with assumptions | Primary description; typos corrected by Claude | Risk Description |
| col 16 | Response Description | Typos corrected by Claude | Mitigating Action |
| col 18 | Residual SEV | `value × 2` | Impact (post-mitigation) |
| col 19 | Residual FRQ | `value × 2` | Likelihood (post-mitigation) |

### Stage Label Mapping

| Source value | Output value |
|---|---|
| `"Design"` | `"Pre-construction"` |
| `"Assembly and commissioning"` | `"Commissioning"` or `"Construction"` (context-dependent) |
| `"Multiple (or all) life phases"` | Inferred from description |
| `"Transportation"` | `"Operation"` |
| `"Normal power production"` | `"Operation"` |
| `"Extreme events"` | `"Operation"` |
| `"Decommissioning"` | `"Decommissioning"` |
| `"NA"` or blank | Inferred from description |

Allowed output values: `Pre-construction`, `Construction`, `Commissioning`,
`Operation`, `Decommissioning`

### Category Inference Rules

| RBS Level + Description context | Output category |
|---|---|
| External + regulatory / licensing | Regulations |
| External + environmental data collection | Planning |
| External + ecological monitoring | Quality |
| Commercial + parts availability / procurement | Procurement |
| Commercial + design-dependent custom components | Design |
| Technical + design validation | Design |
| Technical + structural or assembly | Construction |
| Technical + cable, mooring, or driveline systems | Construction or Regulations |
| Technical + operational risk or financial exposure | Financial |
| Management | Construction |

### Risk ID Sequence

Risk IDs run 1–14, then skip to 16 and continue to 33. The gap at 15
reflects a duplicate row in the source file that is removed during
deduplication (keyed on Risk Name + Mitigation text).

---

## File 2 — City of York Council

### Source Characteristics

11-column Excel sheet already close to the output schema. No Claude API
call — all transformation is deterministic.

### Column Mapping

| Source Col | Source Header | Transformation | Output Field |
|---|---|---|---|
| col 2 | Risk ID | `to_rid()` → integer | Risk ID |
| col 3 | Risk Description | Whitespace cleaned | Risk Description |
| col 4 | Impact description | Newlines → spaces | Mitigating Action (col 10) |
| col 5 | Project Stage | Pass-through | Project Stage |
| col 6 | Risk Category | Pass-through | Project Category |
| col 7 | Likelihood | Float → `int(round(...))` | Likelihood |
| col 8 | Impact | Float → `int(round(...))` | Impact |
| col 9 | Risk Index | Threshold lookup | Risk Priority |
| col 10 | Risk Owner | Pass-through | Risk Owner |
| col 11 | Mitigation | Pass-through | Result (col 11) |

### Likelihood / Impact Rounding

Source values are stored as recurring decimals
(e.g. 3.333…, 6.667…, 8.333…). The model rounds to the nearest integer
before writing: 3.333 → 3, 6.667 → 7, 8.333 → 8.

### Priority — Risk Index Thresholds

Calibrated by analysing all 45 rows of the training pair.

| Risk Index | Priority |
|---|---|
| < 8 | Low |
| 8 – 13.9 | Med |
| ≥ 14 | High |
| 10.5 (exact) | Yellow |
| 17.5 (exact) | (blank) |

---

## File 3 — Digital Security IT Sample Register

### Source Characteristics

10-column cybersecurity register. No Claude API call — all transformation
is deterministic.

### Column Mapping

| Source Col | Source Header | Transformation | Output Field |
|---|---|---|---|
| col 2 | Number | Pass-through | Number |
| col 3 | Risk Description | `\n\n` → single space | Risk Description |
| col 6 | Probability | `to_int()` | Likelihood |
| col 7 | Severity | `to_int()` | Impact |
| col 8 | Score | Text pass-through | Risk Priority |
| col 10 | Action Plan | `\n\n` → single space | Mitigating Action |
| — | (inferred) | Always `"Operations"` | Project Stage |
| — | (inferred) | Keyword lookup on description | Project Category |
| — | (inferred) | Always `"Infrastructure Manager"` | Risk Owner |

### Category Inference

`"Cybersecurity"` is assigned when the description (lowercased) contains
any of the following keywords:
`identity`, `access`, `iam`, `intrusion`, `cyber`, `authentication`,
`phishing`, `malware`, `web application`.

All other rows receive `"Infrastructure"`.

---

## File 4 — Moorgate Crossrail Register  *(blind test)*

### Source Characteristics

10-column urban infrastructure register. Likelihood and Impact are
qualitative text labels. Stage, category, and owner columns are blank.
Claude infers all four missing fields in a single batched API call.

### Column Mapping

| Source Col | Source Header | Transformation | Output Field |
|---|---|---|---|
| col 2 | Risk ID | Pass-through | Risk ID |
| col 3 | Risk Description | Pass-through | Risk Description |
| col 6 | Likelihood text | Claude → integer (lookup fallback) | Likelihood |
| col 7 | Impact text | Claude → integer (lookup fallback) | Impact |
| col 8 | Priority | Text pass-through | Risk Priority |
| col 10 | Mitigating Action | Pass-through | Mitigating Action |
| — | (blank in source) | Claude infers from description | Project Stage |
| — | (blank in source) | Claude infers from description | Project Category |
| — | (blank in source) | Claude infers from description | Risk Owner |

### Qualitative Text → Integer Conversion

**Likelihood:**

| Source text | Integer |
|---|---|
| Rare | 2 |
| Unlikely | 3 |
| Possible | 5 |
| Likely | 7 |
| Almost Certain | 9 |

**Impact (default lookup; Claude refines per context):**

| Source text | Integer |
|---|---|
| Minor | 5 |
| Serious | 7 |
| Major | 8 |

---

## File 5 — Corporate Risk Register / Fenland DC  *(blind test)*

### Source Characteristics

PDF file (or Excel if pre-converted). UK local government corporate risk
register with pre- and post-mitigation scores on a 1–5 scale. Claude
extracts and standardises all fields.

### Processing Flow

1. If the source is a PDF, text is extracted using `pdfminer.six`
   (falls back to `pypdf` or `PyPDF2` if not installed)
2. Extracted text (≤ 12,000 characters) or structured Excel rows sent
   to Claude in a single API call
3. Claude returns a JSON array with one object per risk, all fields populated
4. Model writes the results as Layout A
5. If Claude does not return a priority value, the L × I formula is applied
   as a fallback

### Scale Conversion

Source scores are on a 1–5 scale. The model multiplies all scores by 2
before writing: `1 → 2`, `2 → 4`, `3 → 6`, `4 → 8`, `5 → 10`.

---

## Priority Logic Summary

| File | Method | Formula / Source |
|---|---|---|
| 1 | L × I score formula | Applied to pre and post values independently |
| 2 | Risk Index thresholds | Calibrated from training pair analysis |
| 3 | Source text pass-through | Score column (High / Med) |
| 4 | Source text pass-through | Priority column (Low / Med / High) |
| 5 | Claude + formula fallback | Formula applied if Claude returns no priority |

**Score formula:**
```
L × I  < 32   →  Low
L × I 32–59   →  Med
L × I ≥ 60    →  High
```

---

## Formatting Specification

| Element | Value |
|---|---|
| Header fill | Solid `#E4AFAF` (dusty rose/pink) |
| Header font | Aptos Narrow, 11 pt, bold |
| Header alignment | Centered horizontally and vertically, word-wrapped |
| Header row height | 72.75 pt |
| Data alignment | Centered horizontally and vertically, word-wrapped |
| Cell borders | Thin on all four sides |
| Date number format | `[$-409]d\-mmm\-yy;@` → `21-Feb-17` |
| Integer number format | `0` — no decimal places |
| Hidden reference row | Row 2 in Layout A only |

The header colour is stored as a fixed hex RGB string (`FFE4AFAF`) rather
than an Excel theme index. This ensures consistent rendering across all
Excel installations regardless of the active document theme.
