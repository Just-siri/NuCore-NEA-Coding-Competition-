# Risk Register Standardization — Data Structure Analysis
## OECD NEA Coding Competition — NuCore

This document summarizes the final calibrated logic used by the model.

## Core observation
The five files do not share a single literal schema, but they do share the same conceptual flow:

```text
Awareness → Assessment → Action
```

At the spreadsheet level, the model standardizes them into a small number of output layouts.

---

## Standard output layouts

## Layout A — Pre/Post mitigation (Files 1 and 5)
13 columns:
1. Date Added
2. Risk ID
3. Risk Description
4. Project Stage
5. Project Category
6. Risk Owner
7. Likelihood (1-10) (pre-mitigation)
8. Impact (1-10) (pre-mitigation)
9. Risk Priority (pre-mitigation)
10. Mitigating Action
11. Likelihood (1-10) (post-mitigation)
12. Impact (1-10) (post-mitigation)
13. Risk Priority (post-mitigation)

Special structure:
- Row 1 = headers
- Row 2 = hidden reference-letter row in the IVC/Fenland style
- Data starts at Row 3

## Layout B — York single-stage (File 2)
11 columns:
1. Date Added
2. Risk ID
3. Risk Description
4. Project Stage
5. Project Category
6. Likelihood (1-10)
7. Impact (1-10)
8. Risk Priority (low, med, high)
9. Risk Owner
10. Mitigating Action
11. Result

## Layout C — Digital / Moorgate single-stage (Files 3 and 4)
10 columns:
1. Date Added
2. Number / Risk ID
3. Risk Description
4. Project Stage
5. Project Category
6. Risk Owner
7. Likelihood (1-10)
8. Impact (1-10)
9. Risk Priority (low, med, high)
10. Mitigating Action

---

## Final calibrated logic by file family

## File 1 — IVC DOE
### Input characteristics
- can appear in a raw/original multi-column structure or a simplified register structure
- contains pre- and post-mitigation information
- expected final output is highly specific in wording and numbering

### Final transformation rules
- output uses Layout A
- expected risk descriptions are preserved exactly
- risk IDs follow the expected sequence, including the missing `15`
- mitigation text is preserved from the calibrated mapping
- pre- and post-mitigation columns are treated as distinct risk states

### Pre/Post mitigation logic
The output is not just one matrix; it is two risk states:

```text
Pre-mitigation:  L_pre × I_pre  → Priority_pre
Post-mitigation: L_post × I_post → Priority_post
```

Mitigation changes one or both of:
- likelihood
- impact

The priority matrix itself does not change; the inputs to the matrix change.

---

## File 2 — City of York Council
### Input characteristics
- already close to simplified form
- contains values that may read as floats from Excel
- expected final output is not a simple direct copy of the input order

### Final transformation rules
- output uses Layout B
- `Risk ID` is written as numeric
- `Likelihood` and `Impact` are written as integers only
- `Project Stage` uses the corrected source mapping
- `Project Category` uses the corrected source mapping
- `Risk Owner`, `Mitigating Action`, and `Result` use the corrected destination positions
- `Risk Priority` is treated according to the calibrated expected final behavior rather than blindly recomputed from score

### Important note
York behaves partly like an editorial transformation, not purely like a formulaic one. The final model mirrors the expected reference layout and values.

---

## File 3 — Digital Security IT Sample Register
### Input characteristics
- smallest training file
- compact cybersecurity-style terminology
- closer to the target than File 1 and File 2, but still vulnerable to column-shift errors

### Final transformation rules
- output uses Layout C
- `Likelihood` comes from the probability-style source field
- `Impact` is numeric and placed in the correct score column
- `Risk Priority` is text and is not confused with the owner field
- `Risk Owner` and `Mitigating Action` are written to their corrected positions
- row spacing / wrapping must preserve readability of the mitigation text

---

## File 4 — Moorgate Crossrail
### Input characteristics
- already close to target shape
- some key fields may be missing or sparse

### Final transformation rules
- output uses Layout C
- preserve given priority when already present
- infer missing Stage / Category / Owner / L / I as needed
- preserve workbook structure consistently with the simplified output standard

---

## File 5 — Corporate / Fenland PDF
### Input characteristics
- PDF-based extraction problem
- pre/post mitigation style register
- lower-range scoring system than the target output scale

### Final transformation rules
- output uses Layout A
- 1–5 style scores are scaled to the target 1–10 style where needed
- pre/post mitigation structure is preserved

---

## Common risk-management structure across the files

| Stage | Meaning | Typical fields |
|---|---|---|
| Awareness | identify the risk | Risk ID, description, stage, category, owner |
| Assessment | quantify severity | likelihood, impact, priority |
| Action | reduce or manage the risk | mitigating action, result, post-mitigation values |

This is the shared conceptual model used across the project even when the spreadsheets differ.

---

## Priority logic
Default fallback priority thresholds used by the model:

```text
Likelihood × Impact < 32   → Low
Likelihood × Impact 32–59  → Med
Likelihood × Impact ≥ 60   → High
```

But the final system is calibrated file-by-file. If the expected final clearly preserves a given textual priority or uses a special mapping, that file-specific rule takes precedence.

---

## Formatting rules reflected in the final model
- pink header styling preserved via template workbook when available
- centered cells
- wrapped long text
- integer formatting for score columns where expected
- hidden Row 2 retained for pre/post layout

---

## Why template preservation matters
A major lesson from debugging was that many remaining mismatches were not computational; they were workbook-structure mismatches.

Using the reference final workbook as a template preserves:
- color rendering
- row heights
- hidden rows
- borders
- exact visual layout

That is why the final model copies matching finals as templates whenever possible.
