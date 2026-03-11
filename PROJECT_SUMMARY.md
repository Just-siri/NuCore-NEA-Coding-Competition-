# OECD NEA Coding Competition - Project Summary

## What Was Delivered

I have analyzed the risk register standardization challenge and built a complete Python-based solution using Large Language Models (LLMs).

## Files Delivered

### 1. **model.py** (Main Model)
- Complete object-oriented implementation of the risk register standardizer (`RiskRegisterStandardizer` class)
- Uses Claude Sonnet 4.5 for intelligent data extraction and transformation
- Handles both Excel and PDF inputs
- Produces standardized Excel outputs
- Secure API key management via `.env` files
- ~350 lines of well-documented class-based code

### 2. **data_structure_analysis.md** (Analysis Document)
- Comprehensive analysis of all input and output formats
- Documents transformation rules and patterns
- Identifies challenges and edge cases
- Maps column transformations
- Defines risk priority calculation logic

### 3. **requirements.txt** (Dependencies)
- Lists all required Python packages:
  - pandas (data manipulation)
  - openpyxl (Excel creation)
  - pdfplumber (PDF extraction)
  - anthropic (LLM API)
  - python-dotenv (environment variable management)

### 4. **test_model.py** (Testing Script)
- Validates model against the 3 training examples
- Compares generated outputs with expected outputs
- Provides detailed test reports

### 5. **README.md** (Documentation)
- Complete usage instructions
- Architecture overview
- Transformation rules
- Error handling details
- Future improvement suggestions

## Key Findings from Analysis

### Input File Characteristics

**File 1 (IVC DOE)**: 
- Most complex: 982 rows, 24 columns, multi-row headers
- Contains pre/post mitigation data
- Requires sophisticated header parsing

**File 2 (City of York Council)**:
- Simpler format: 45 rows, 11 columns
- Direct column mappings needed
- Numeric risk scores need categorization

**File 3 (Digital Security IT)**:
- Smallest: 3 rows, 10 columns
- Different scale (needs normalization)
- Different terminology for similar concepts

**File 4 (Moorgate Crossrail)** - BLIND TEST:
- 10 rows, already well-structured
- Close to target format
- Minimal transformations needed

**File 5 (Corporate Risk Register)** - BLIND TEST:
- PDF format with 21 pages
- Complex table structure with rotated headers
- Contains Fenland District Council data
- Pre/post mitigation scoring

### Mandatory Output Columns

All outputs must include these 9 fields:
1. Risk ID
2. Risk Description  
3. Project Stage
4. Project Category
5. Risk Owner
6. Mitigating Action
7. Likelihood (1-10)
8. Impact (1-10)
9. Risk Priority (low, med, high)

### Key Transformation Patterns

**Column Name Standardization**:
- "Risk Category" → "Project Category"
- "Mitigation" / "Action Plan" → "Mitigating Action"
- "Number" → "Risk ID"
- "Probability" → "Likelihood (1-10)"
- "Severity" → "Impact (1-10)"

**Risk Priority Calculation**:
- score = Likelihood × Impact
- Low: score ≤ 10
- Medium: 10 < score ≤ 30
- High: score > 30

**Score Scaling**:
- 1-5 scale: multiply by 2
- 1-3 scale: map to 3, 6, 9
- 1-10 scale: keep as-is

## Model Architecture

```
Input File (Excel/PDF)
    ↓
File Type Detection
    ↓
Data Extraction (pandas/pdfplumber)
    ↓
LLM Analysis (Claude Sonnet 4.5)
    ├─ Column Identification
    ├─ Transformation Planning
    ├─ Data Standardization
    └─ Missing Data Inference
    ↓
Output Generation (openpyxl)
    ├─ Simplified Register sheet
    └─ Output Requirements sheet
    ↓
Standardized Excel File
```

## How the Model Works

1. **Detection**: Identifies if input is Excel or PDF
2. **Extraction**: Uses appropriate library to extract raw data
3. **LLM Processing**: Sends data to Claude with detailed instructions
4. **Transformation**: Claude identifies columns, applies transformations, and outputs standardized CSV
5. **Output Creation**: Creates Excel file with proper formatting
6. **Validation**: Ensures all mandatory fields are present

## Challenges Addressed

✓ Complex multi-row headers (File 1)
✓ PDF table extraction with rotated text (File 5)
✓ Different scoring scales across files
✓ Varied column naming conventions
✓ Pre/post mitigation data handling
✓ Missing mandatory field generation
✓ Risk priority calculation from scores

## Testing Approach

The `test_model.py` script:
- Processes all 3 training files
- Compares outputs with expected results
- Reports on:
  - Column matching
  - Mandatory field presence
  - Row count accuracy
  - Sample data comparison

## Usage Instructions

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Set API key
# Create a .env file in the root directory with the following content:
# ANTHROPIC_API_KEY=your-key-here

# 3. Set up directories
mkdir -p input output

# 4. Place input files
cp 4__Moorgate_Crossrail_Register__Input_.xlsx input/
cp 5__Corporate_Risk_Register__Input_.pdf input/

# 5. Run model
python model.py

# 6. Check outputs
ls output/
```

## Expected Performance

- **Processing Time**: 30-60 seconds per file
- **Token Usage**: ~2,000-8,000 tokens per file
- **Accuracy**: High for well-structured data, good for complex formats
- **Robustness**: Fallback mechanisms for edge cases

## Strengths of This Solution

1. **LLM-Powered Intelligence**: Uses Claude to understand diverse formats
2. **Flexible**: Handles both Excel and PDF inputs
3. **Robust**: Multiple fallback mechanisms
4. **Well-Documented**: Clear code comments and comprehensive docs
5. **Tested**: Validation against training data
6. **Professional Output**: Properly formatted Excel files

## Potential Improvements

- Add caching to reduce API calls for similar files
- Implement confidence scoring for extracted data
- Add more sophisticated PDF table parsing
- Support batch processing optimization
- Fine-tune thresholds based on validation feedback

## Competition Compliance

✓ Written in Python
✓ Single self-contained model.py file
✓ Uses LLM/NLP as required
✓ Produces standardized Excel outputs
✓ Includes all mandatory columns
✓ Works with input/ and output/ directories
✓ Requires only API key as external dependency

## Recommendations for Blind Test Execution

1. **File 4 (Moorgate Crossrail)**: Should process smoothly - already close to target format
2. **File 5 (Corporate Risk PDF)**: May require manual review - PDF extraction can be challenging
3. Run test_model.py first to validate setup
4. Check API rate limits if processing multiple files
5. Review outputs manually before submission

## Conclusion

This solution provides a robust, LLM-powered approach to risk register standardization that can handle diverse input formats and produce consistent, machine-readable outputs. The model leverages Claude's language understanding to intelligently map and transform data while maintaining flexibility through fallback mechanisms.
