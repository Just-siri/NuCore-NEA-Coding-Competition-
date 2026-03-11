# Risk Register Standardization - Data Structure Analysis

## Competition Overview
Convert diverse risk registers (Excel, PDF, Word) into standardized machine-readable format using NLP/LLMs.

## Training Data Summary

### File 1: IVC DOE R2
- **Input**: Complex Excel with multiple sheets ('Intro', 'Risk Register', 'Valid data fields')
  - Risk Register: 982 rows × 24 columns with messy structure
  - Column names like: 'IDENTIFY RISKS PROCESS', 'ANALYZE RISKS PROCESS', etc.
- **Output**: 32 risk records × 13 standardized columns

### File 2: City of York Council
- **Input**: Semi-structured Excel, 45 rows × 11 columns
  - Already has columns like: Date Added, Risk ID, Risk Description, Impact, etc.
- **Output**: 44 risk records × 11 standardized columns

### File 3: Digital Security IT Sample Register
- **Input**: Small structured Excel, 3 rows × 10 columns
  - Columns: Date Added, Number, Risk Description, Probability, Severity, Score, etc.
- **Output**: 2 risk records × 10 standardized columns

## Blind Test Files

### File 4: Moorgate Crossrail Register (Excel)
- **Input**: 10 rows × 10 columns, already well-structured
- Columns: Date Added, Risk ID, Risk Description, Project Stage, Risk Category, Likelihood (1-10), Impact (1-10), Risk Priority, Risk Owner, Mitigating Action

### File 5: Corporate Risk Register (PDF)
- **Input**: 21-page PDF with complex table structure
- Fenland District Council document
- Columns: Reference, Risk and effects, Impact (pre/post), Likelihood (pre/post), Score, Mitigation, Risk Owner, Actions, Comments

## Standard Output Format

### Core Output Structure
All outputs have two sheets:
1. **"Simplified Register"** - Main risk data
2. **"Output Requirements"** - Documentation of field requirements

### Standard Output Columns (varies by file, but core set includes):

#### Most Complete Format (File 1):
1. Date Added
2. Risk ID
3. Risk Description
4. Project Stage
5. Project Category
6. Risk Owner
7. Likelihood (1-10) (pre-mitigation)
8. Impact (1-10) (pre-mitigation)
9. Risk Priority (pre-mitigation) - [Low/Med/High]
10. Mitigating Action
11. Likelihood (1-10) (post-mitigation)
12. Impact (1-10) (post-mitigation)
13. Risk Priority (post-mitigation) - [Low/Med/High]

#### Simplified Format (Files 2-3):
- Date Added
- Risk ID / Number
- Risk Description
- Project Stage
- Project Category
- Risk Owner
- Likelihood (1-10)
- Impact (1-10)
- Risk Priority (low, med, high)
- Mitigating Action
- Result (File 2 only)

## Key Transformations Required

### 1. Column Mapping
- Map various input column names to standard output names
- Examples:
  - "Number" → "Risk ID"
  - "Probability" → "Likelihood (1-10)"
  - "Severity" → "Impact (1-10)"
  - "Action Plan" → "Mitigating Action"
  - "Mitigation" → "Mitigating Action"

### 2. Score Normalization
- Convert different scoring systems to 1-10 scale
- Example: 5-point scale → 10-point scale (multiply by 2)
- Calculate from text descriptions if needed

### 3. Risk Priority Calculation
- Calculate from Likelihood × Impact
- Mapping:
  - Low: Score ≤ 30
  - Med: Score 31-60
  - High: Score > 60
  - (This is approximate and may vary)

### 4. Pre/Post Mitigation Detection
- Identify if document has before/after mitigation columns
- Create separate columns or single set based on availability

### 5. Text Extraction and Cleaning
- Extract risk descriptions from merged cells
- Clean formatting artifacts
- Handle multi-line text with \n characters
- Parse complex table structures from PDFs

### 6. Data Quality
- Remove header rows
- Remove empty rows
- Handle missing values appropriately
- Ensure data types are correct (dates, numbers, text)

## Input Format Challenges

### Excel Files
1. **Multi-sheet documents** - Need to identify which sheet contains actual risk data
2. **Messy headers** - Multiple header rows, merged cells
3. **Varying column names** - Same concept, different labels
4. **Different scoring systems** - 1-5, 1-10, text-based, color-coded
5. **Embedded instructions** - Template text that should be removed

### PDF Files
1. **Complex table structures** - Multi-row headers, merged cells
2. **Text extraction issues** - OCR artifacts, formatting
3. **Multiple tables per page** - Need to identify and merge
4. **Inconsistent formatting** - Column widths, line breaks

## Mandatory Output Fields

From Output Requirements sheet (File 1):
- **Risk ID** - If not provided, any identifier may be used
- **Risk Description**
- **Project Stage** - Required for construction or project-based risks
- **Project Category**
- **Risk Owner**
- **Mitigating Action**
- **Likelihood (1-10)** - If multiple stages provided, include pre/post mitigation
- **Impact (1-10)**
- **Risk Priority** (low, med, high)

## Implementation Strategy

### Phase 1: Data Loading
- Detect file type (Excel, PDF, Word)
- Load data using appropriate library
- Identify risk data location (sheet, page, table)

### Phase 2: Column Identification
- Use LLM to map input columns to standard output columns
- Handle variations in naming
- Identify pre/post mitigation columns if present

### Phase 3: Data Extraction
- Extract risk records
- Parse complex structures
- Handle merged cells and multi-line text

### Phase 4: Score Normalization
- Detect scoring system (1-5, 1-10, text-based)
- Convert to 1-10 scale
- Calculate risk priority from likelihood × impact

### Phase 5: Standardization
- Apply standard column names
- Format data consistently
- Create required output structure (two sheets)

### Phase 6: Quality Assurance
- Validate all mandatory fields present
- Check data types
- Ensure no formula errors
- Verify risk priority calculations
