# Data Structure Analysis
### OECD NEA Coding Competition — NuCore

Detailed column mappings, transformation rules, and output specifications
for each file family processed by `model.py`. Observations marked
**[docx]** are drawn directly from the input/output difference analysis.

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
unified output schema. However, the three training files represent three
distinct levels of input completeness and transformation complexity —
they cannot be processed with a single uniform rule set.

| File | Input completeness | Transformation complexity |
|---|---|---|
| 1 — IVC DOE R2 | Many fields absent or structured differently | High — inference, scale conversion, field derivation |
| 2 — City of York Council | Near-complete, close to target schema | Low — mostly column renames and threshold application |
| 3 — Digital Security IT | 4 fields entirely blank | Medium — inference only, no scale conversion |

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
| 1 | Date Added | Datetime (blank for all 45 rows) |
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

Raw 24-column register with a 3-row merged header, organised into four
process groups: Identify, Analyse, Respond, and Contingency. Scores use a
0–5 scale multiplied by 2 to produce the 1–10 output scale. The output
collapses this to 13 columns. **This file requires the most interpretation
— several output fields have no direct equivalent in the input and must be
derived. [docx]**

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

**Dropped columns [docx]:** RPN pre/post, +/– flag (all entries negative),
TYP code, TRL, TPL, Response Strategy, Response Timing, Residual Description,
Secondary Risks, Recommendations & Action Items, Contingency Plan — none
appear in the output.

### Awareness — Risk Description, Stage, Category, Owner

**Risk Description [docx]:** The output uses col 14 (*Baseline Description
with assumptions*), not col 4 (*Risk Name*). Col 4 contains short titles;
col 14 contains the full risk statement. Where col 14 is generic and col 4
provides a component name, Claude prefixes col 14 with the component name.

**Project Stage [docx]:** The input uses MHK-specific lifecycle vocabulary
that must be mapped to standard labels:

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

**Project Category [docx]:** No single category field exists in the input.
Derived from three fields: RBS Level 1, RBS Level 2, and TYP code, via
Claude interpretation:

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

**Risk Owner [docx]:** Input stores full name plus role in parentheses
(e.g. `"R. Tyler (lead engineer)"`). Output retains the role only, dropping
the personal name (e.g. `"Lead engineer"`).

**Risk ID [docx]:** No ID column exists in the input. Sequential integers
are auto-generated. IDs run 1–14, then skip to 16 and continue to 33 — the
gap at 15 reflects a duplicate row removed during deduplication.

### Assessment — Likelihood, Impact, Priority

**Scale conversion [docx]:** FRQ (Frequency) and SEV (Severity) are on a
0–5 integer scale. Output requires 1–10 even numbers only. Conversion:
`output = input × 2`.

**Post-mitigation scores [docx]:** For most rows the conversion is
mechanical. In approximately 13 high-consequence rows the post-mitigation
Impact score deviates upward from the formula, reflecting qualitative
judgement that the risk remains severe. Likelihood always decreases after
mitigation (minimum –2 on the 0–10 output scale). Impact is unchanged in
most rows; for SEV=5 (catastrophic) risks, Impact stays at 10 regardless
of strategy.

**Priority formula:**
```
L × I  < 32   →  Low
L × I 32–59   →  Med
L × I ≥ 60    →  High
```
Applied independently to pre- and post-mitigation values. Post-mitigation
priority is always ≤ pre-mitigation priority.

### Action — Mitigating Action

**[docx]:** The input's Response Description (col 16) maps directly to
Mitigating Action. It is copied verbatim with minor spelling corrections —
no summarisation. The input contains 10 action-related columns; only this
one is carried forward. Response Strategy (col 14), Response Timing (col
16), and all four Contingency columns are dropped.

---

## File 2 — City of York Council

### Source Characteristics

11-column Excel sheet with 45 rows. Input and output have the same Risk IDs
and the same row count. **Most of the transformation is a renaming exercise.
No Claude API call is made — all transformation is deterministic. [docx]**

### Column Mapping

| Source Col | Source Header | Transformation | Output Field |
|---|---|---|---|
| col 2 | Risk ID | `to_rid()` → integer | Risk ID |
| col 3 | Risk Description | Whitespace cleaned | Risk Description |
| col 4 | Impact | Newlines → spaces; passed through | Mitigating Action (col 10) |
| col 5 | Project Stage | Pass-through | Project Stage |
| col 6 | Risk Category | Column renamed | Project Category |
| col 7 | Likelihood | Float → `int(round(...))` | Likelihood |
| col 8 | Impact | Float → `int(round(...))` | Impact |
| col 9 | Risk Index | Threshold lookup | Risk Priority |
| col 10 | Risk Owner | Pass-through | Risk Owner |
| col 11 | Mitigation | Column renamed | Result (col 11) |

No fields are dropped — all 11 input columns are accounted for in the
11 output columns.

### Awareness Fields

**[docx]:** Risk Description, Project Stage, Project Owner all pass through
verbatim. Project Category is a pure rename of Risk Category — values are
identical. Row 27 has a joint owner (`"Project Manager / Environment Lead"`)
which is simplified to `"Environment Lead"` in the output.

### Assessment — Likelihood, Impact, Priority

**Likelihood / Impact rounding [docx]:** Values are stored as recurring
decimals reflecting an underlying 1–6 integer rating mapped to tenths of 10
(e.g. 1.67, 3.33, 5, 6.67, 8.33, 10). The model rounds to the nearest
integer: 3.333 → 3, 6.667 → 7, 8.333 → 8. No scale conversion is applied.

**Priority — Risk Index thresholds [docx]:** The register carries a
pre-scored Risk Index per row that does not equal Likelihood × Impact in any
of the 45 rows. It is independently assigned by the project team. The Index,
not the L×I product, determines the output priority label.

| Risk Index | Priority |
|---|---|
| ≤ 7.5 | Low (20 rows) |
| 8.0 – 13.0 | Med (13 rows) |
| ≥ 14.0 | High (10 rows) |
| 10.5 (exact) | Yellow (Risk 13 — one edge case) |
| 17.5 (exact) | (blank) — Risk 32, no priority assigned |

**Anomaly [docx]:** Risk 35 has Risk Index = 6.0, which falls in the Low
range by the threshold rule, but is assigned High in the output. This is
the only row where the threshold rule is overridden, with no documented
explanation.

### Action — Mitigating Action and Result

**[docx]:** The input column named `"Impact"` contains descriptive text of
the consequence if the risk materialises — not a mitigation plan. This is
renamed `"Mitigating Action"` in the output. The input column named
`"Mitigation"` contains the actual action being taken; this is renamed
`"Result"`. The naming is counterintuitive but is preserved exactly as it
appears in the reference output. No post-mitigation scores exist — the
register tracks only a single priority label per risk.

---

## File 3 — Digital Security IT Sample Register

### Source Characteristics

10-column register with only 3 rows. **The central challenge is not
transformation but inference: Project Stage, Risk Category, and Risk Owner
are entirely blank for at least one row. The output fills all of these.
[docx]**

### Column Mapping

| Source Col | Source Header | Transformation | Output Field |
|---|---|---|---|
| col 2 | Number | Pass-through | Number |
| col 3 | Risk Description | `\n\n` → single space | Risk Description |
| col 6 | Probability | `to_int()` | Likelihood |
| col 7 | Severity | `to_int()` | Impact |
| col 8 | Score | Text pass-through | Risk Priority |
| col 10 | Action Plan | `\n\n` → single space; column renamed | Mitigating Action |
| — | (blank in source) | Inferred: always `"Operations"` | Project Stage |
| — | (blank in source) | Inferred from description keywords | Project Category |
| — | (blank in source) | Inferred: always `"Infrastructure Manager"` | Risk Owner |

No fields are dropped — all 10 input columns map to 10 output columns.

### Awareness — Inferred Fields

**Project Stage [docx]:** Blank for all 3 rows. Assigned `"Operations"` for
all, derived from the IT operational context of the risk descriptions.

**Project Category [docx]:** Blank for all 3 rows. Inferred from risk
content: ICT-001 (IAM / external intrusion) → `"Cybersecurity"`;
ICT-002 and ICT-003 (infrastructure / backup) → `"Infrastructure"`.

`"Cybersecurity"` is assigned when the description contains any of:
`identity`, `access`, `iam`, `intrusion`, `cyber`, `authentication`,
`phishing`, `malware`, `web application`.

**Risk Owner [docx]:** Blank for ICT-001. `"Infrastructure Manager"`
provided for ICT-002 and ICT-003. The output assigns `"Infrastructure Manager"`
to all 3 rows — ICT-001 value inferred from the other two rows and the
nature of the risk.

### Assessment — Notable Anomaly

**[docx]:** The qualitative Score field overrides what the numeric
Probability × Severity product would produce, and this is preserved in the
output without correction:

| Row | Probability | Severity | P × S | Score (input) | Priority (output) |
|---|---|---|---|---|---|
| ICT-001 | 3 | 7 | 21 | High | High |
| ICT-002 | 9 | 7 | 63 | Med | Med |
| ICT-003 | 10 | 7 | 70 | High | High |

ICT-002 has the highest numeric product (63) yet the lowest priority label
(Med). The Score column is qualitative and is carried verbatim — the model
does not recompute from the numeric values.

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

PDF file (or Excel / Word if pre-converted). UK local government corporate
risk register with pre- and post-mitigation scores on a 1–5 scale. Claude
extracts and standardises all fields.

### Processing Flow

1. Text extracted from the source file using `pdfplumber` (PDF),
   `python-docx` (Word), or `openpyxl` (Excel)
2. Extracted content (≤ 12,000 characters) sent to Claude in a single API call
3. Claude returns a JSON array with one object per risk, all fields populated
4. Model writes the results as Layout A
5. If Claude does not return a priority value, the L × I formula is applied
   as a fallback

### Scale Conversion

Source scores are on a 1–5 scale. The model multiplies all scores by 2
before writing: `1 → 2`, `2 → 4`, `3 → 6`, `4 → 8`, `5 → 10`.

---

## Priority Logic Summary

| File | Method | Basis |
|---|---|---|
| 1 | L × I score formula | Applied to pre and post values independently |
| 2 | Risk Index thresholds | Index ≠ L×I in all 45 rows — independently pre-scored |
| 3 | Source Score pass-through | Qualitative label, not computed from P×S |
| 4 | Source Priority pass-through | Text value preserved from input |
| 5 | Claude + formula fallback | Formula applied if Claude returns no priority |

**Score formula (Files 1, 5, fallback):**
```
L × I  < 32   →  Low
L × I 32–59   →  Med
L × I ≥ 60    →  High
```

---

## Cross-File Comparison  *(from training data analysis)*

| Dimension | IVC DOE R2 | City of York Council | Digital Security IT |
|---|---|---|---|
| Awareness inference | Stage, Category (RBS+TYP), Owner role extracted | Column renames only | Stage, Category, Owner for ICT-001 — all inferred |
| Score conversion | Input 0–5 × 2 = output 0–10 | No conversion — pass-through | No conversion — pass-through |
| Priority basis | Computed L×I with fixed thresholds | Pre-scored Risk Index (Index ≠ L×I in all rows) | Qualitative Score field — not recomputed |
| Pre/post assessment | Full pre and post. Post from Residual SEV/FRQ | Pre only. No residual scoring | Pre only. No residual scoring |
| Action field mapping | Response Description → Mitigating Action | Impact text → Mitigating Action; Mitigation text → Result | Action Plan → Mitigating Action |
| Key anomaly | ~13 post-mitigation Impact scores deviate upward for high-consequence risks | Risk 35 (Index=6) assigned High against threshold rules; Risk 32 has no priority; one Yellow label | ICT-002 has highest P×S (63) but lowest priority (Med) |

---

## Formatting Specification

| Element | Value |
|---|---|
| Header fill | Solid `#E4AFAF` (dusty rose/pink) |
| Header font | Calibri, 11 pt, bold |
| Header alignment | Centered horizontally and vertically, word-wrapped |
| Data alignment | Centered horizontally and vertically, word-wrapped |
| Cell borders | Thin on all four sides |
| Date number format | `[$-409]d\-mmm\-yy;@` → `21-Feb-17` |
| Integer number format | `0` — no decimal places |
| Hidden reference row | Row 2 in Layout A only |

The header colour is stored as a fixed hex RGB string rather than an Excel
theme index, ensuring consistent rendering across all installations
regardless of the active document theme.
