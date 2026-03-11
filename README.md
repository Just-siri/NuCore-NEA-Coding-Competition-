# Risk Register Standardization Model

## Overview

This model converts diverse risk registers (Excel and PDF formats) into standardized, machine-readable formats using Large Language Models (LLMs) and natural language processing.

## Features

- **Multi-format support**: Handles Excel (.xlsx) and PDF files
- **Intelligent column mapping**: Uses Claude LLM to identify and map columns
- **Automatic data transformation**: Scales scores, categorizes risks, generates missing IDs
- **Robust extraction**: Handles complex layouts including multi-row headers and PDF tables
- **Standardized output**: Produces consistent Excel files with mandatory fields

## Mandatory Output Fields

1. Risk ID
2. Risk Description
3. Project Stage
4. Project Category
5. Risk Owner
6. Mitigating Action
7. Likelihood (1-10)
8. Impact (1-10)
9. Risk Priority (low, med, high)

## Installation

```bash
# Install required packages
pip install -r requirements.txt

# Set API key (required)
# Create a .env file in the root directory and add:
# ANTHROPIC_API_KEY=your-anthropic-api-key-here
```

## Usage

### Standard Usage (Competition Format)

```bash
# Create input and output directories
```bash
# Mac/Linux
mkdir -p input output
# Windows
mkdir input output
```

# Place input files in the input/ directory
cp your_risk_register.xlsx input/
cp another_register.pdf input/

# Run the model
python model.py

# Outputs will be in the output/ directory
```

### Testing with Training Data

```bash
# Test against the provided training examples
python test_model.py
```

This will:
- Process the 3 training input files
- Compare outputs with expected results
- Generate a test report

## How It Works

### 1. File Detection & Extraction
- Detects file type (Excel or PDF)
- Extracts data using appropriate method:
  - **Excel**: Uses pandas/openpyxl, handles multi-sheet workbooks and complex headers
  - **PDF**: Uses pdfplumber to extract tables and text

### 2. LLM-Powered Analysis
- Sends extracted data to Claude Sonnet 4.5
- LLM performs:
  - Column identification and mapping
  - Data transformation planning
  - Risk categorization
  - Missing data inference

### 3. Standardization
- Applies column name mappings
- Scales likelihood/impact scores to 1-10
- Calculates risk priority (Low/Medium/High)
- Generates missing Risk IDs
- Handles pre/post mitigation data

### 4. Output Generation
- Creates Excel workbook with:
  - "Simplified Register" sheet (standardized data)
  - "Output Requirements" sheet (documentation)
- Applies professional formatting

## Transformation Rules

### Column Mappings
```
"Risk Category" → "Project Category"
"Mitigation" / "Action Plan" → "Mitigating Action"
"Number" / "ID" → "Risk ID"
"Probability" / "Freq" → "Likelihood (1-10)"
"Severity" / "SEV" → "Impact (1-10)"
```

### Score Scaling
- **1-5 scale**: Multiply by 2 to get 1-10
- **1-3 scale**: Map 1→3, 2→6, 3→9
- **1-10 scale**: Keep as-is

### Risk Priority Calculation
```python
score = Likelihood × Impact

if score ≤ 10:
    priority = "Low"
elif score ≤ 30:
    priority = "Medium"
else:
    priority = "High"
```

## Architecture

```
model.py
├── RiskRegisterStandardizer (main class)
│   ├── process_file()           # Entry point
│   ├── extract_from_excel()     # Excel data extraction
│   ├── extract_from_pdf()       # PDF data extraction
│   ├── standardize_data()       # LLM-powered transformation
│   ├── manual_extraction()      # Fallback extraction
│   └── create_output_file()     # Excel output generation
└── main()                       # CLI entry point
```

## Error Handling

The model includes robust error handling:
- **Fallback extraction**: If LLM fails, uses rule-based extraction
- **Missing data**: Generates reasonable defaults
- **Format detection**: Auto-detects Excel vs PDF
- **Column validation**: Ensures all mandatory fields present

## Limitations

- Requires Anthropic API key
- PDF extraction depends on table structure quality
- Very large files (>1000 rows) are truncated for LLM processing
- Complex PDF layouts may require manual review

## Files

- `model.py` - Main standardization model
- `requirements.txt` - Python dependencies
- `test_model.py` - Testing script for training data
- `data_structure_analysis.md` - Detailed analysis of input/output formats

## Testing

The model has been tested on:
1. IVC DOE R2 register (complex multi-sheet Excel)
2. City of York Council register (simple single-sheet Excel)
3. Digital Security IT Sample register (small dataset)

Run tests with:
```bash
python test_model.py
```

## Performance Considerations

- **Speed**: ~30-60 seconds per file (depends on LLM API latency)
- **Accuracy**: Depends on LLM's interpretation of input format
- **Token usage**: ~2000-8000 tokens per file

## Future Improvements

- Add caching for repeated similar files
- Implement batch processing optimization
- Add confidence scores for extracted data
- Support additional file formats (Word, CSV)
- Fine-tune transformation thresholds based on feedback

## License

This model was created for the OECD NEA Coding Competition.

## Contact

For questions or issues, refer to the competition guidelines.
