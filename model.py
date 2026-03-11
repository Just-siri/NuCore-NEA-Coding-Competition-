"""
Risk Register Standardization Model
OECD NEA Coding Competition

This model converts diverse risk registers into standardized machine-readable format.
Uses Claude API for intelligent column mapping and data extraction.
"""

import pandas as pd
import pdfplumber
import json
import os
import sys
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any
import re
import difflib
from datetime import datetime
import anthropic
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

class RiskRegisterStandardizer:
    def __init__(self, api_key: Optional[str] = None):
        """Initialize the standardizer with Anthropic API key."""
        self.api_key = api_key or os.getenv('ANTHROPIC_API_KEY')
        if not self.api_key:
            raise ValueError("ANTHROPIC_API_KEY must be set in environment or passed to initialization")
        
        self.client = anthropic.Anthropic(api_key=self.api_key)
        self.model = "claude-sonnet-4-20250514"

    def load_risk_register(self, file_path: str) -> Tuple[Optional[pd.DataFrame], str]:
        """Load risk register from various formats (Excel, PDF)."""
        file_ext = Path(file_path).suffix.lower()
        
        if file_ext in ['.xlsx', '.xls']:
            try:
                excel_file = pd.ExcelFile(file_path)
                sheet_names = excel_file.sheet_names
                
                risk_sheet = None
                for sheet in sheet_names:
                    if 'risk' in sheet.lower() and 'register' in sheet.lower():
                        risk_sheet = sheet
                        break
                
                if risk_sheet is None:
                    risk_sheet = sheet_names[0]
                
                df = pd.read_excel(file_path, sheet_name=risk_sheet)
                return df, 'excel'
            except Exception as e:
                print(f"Error loading Excel: {e}")
                return None, 'excel'
        
        elif file_ext == '.pdf':
            try:
                df = self.extract_pdf_tables(file_path)
                return df, 'pdf'
            except Exception as e:
                print(f"Error loading PDF: {e}")
                return None, 'pdf'
        
        else:
            return None, 'unknown'

    def extract_pdf_tables(self, pdf_path: str) -> pd.DataFrame:
        """Extract tables from PDF and combine them."""
        all_rows = []
        
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    if not table:
                        continue
                    for row in table:
                        if row and any(cell for cell in row if cell and str(cell).strip()):
                            all_rows.append(row)
        
        if not all_rows:
            return pd.DataFrame()
        
        return pd.DataFrame(all_rows)

    def identify_columns_with_llm(self, df: pd.DataFrame, sample_size: int = 10) -> Dict[str, str]:
        """Use Claude API to identify column mappings."""
        # Smart Row Sampling: Find the rows with the most non-null values
        # This skips over title/metadata rows at the top of Excel files
        row_non_null_counts = df.notna().sum(axis=1)
        densest_rows_idx = row_non_null_counts.nlargest(sample_size).index.sort_values()
        
        sample_data = {
            'columns': list(df.columns),
            'sample_rows': df.loc[densest_rows_idx].replace({pd.NA: None}).to_dict('records'),
            'shape': df.shape
        }
        
        prompt = f"""You are analyzing a risk register to map columns to standardized fields.

INPUT DATA:
Columns: {sample_data['columns']}
Sample rows:
{json.dumps(sample_data['sample_rows'], indent=2, default=str)}

Map to these STANDARD FIELDS:
1. Date Added
2. Risk ID (any identifier)
3. Risk Description
4. Project Stage
5. Project Category
6. Risk Owner
7. Likelihood - Pre-mitigation (1-10 scale)
8. Impact - Pre-mitigation (1-10 scale)
9. Risk Priority - Pre-mitigation (Low/Med/High)
10. Mitigating Action
11. Likelihood - Post-mitigation (optional)
12. Impact - Post-mitigation (optional)
13. Risk Priority - Post-mitigation (optional)

OUTPUT JSON (no explanation):
{{
  "Date Added": "column_name or null",
  "Risk ID": "column_name or null",
  "Risk Description": "column_name or null",
  "Project Stage": "column_name or null",
  "Project Category": "column_name or null",
  "Risk Owner": "column_name or null",
  "Likelihood_Pre": "column_name or null",
  "Impact_Pre": "column_name or null",
  "Risk_Priority_Pre": "column_name or null",
  "Mitigating_Action": "column_name or null",
  "Likelihood_Post": "column_name or null",
  "Impact_Post": "column_name or null",
  "Risk_Priority_Post": "column_name or null",
  "scoring_scale": "1-5 or 1-10",
  "has_pre_post_separate": true/false,
  "data_start_row": 0,
  "notes": "observations"
}}"""

        message = self.client.messages.create(
            model=self.model,
            max_tokens=2000,
            messages=[{"role": "user", "content": prompt}]
        )
        
        response_text = message.content[0].text
        json_match = re.search(r'```json\s*(.*?)\s*```', response_text, re.DOTALL)
        if json_match:
            response_text = json_match.group(1)
        
        try:
            return json.loads(response_text)
        except json.JSONDecodeError as e:
            print(f"Error parsing LLM response: {e}")
            return {}

    def extract_risk_data_with_llm(self, df: pd.DataFrame, column_mapping: Dict[str, str]) -> pd.DataFrame:
        """Extract and clean risk data using column mapping."""
        data_start_row = column_mapping.get('data_start_row', 0)
        
        if data_start_row > 0:
            df = df.iloc[data_start_row:].reset_index(drop=True)
        
        df = df.dropna(how='all')
        
        standardized_data = {}
        standard_fields = [
            'Date Added', 'Risk ID', 'Risk Description', 'Project Stage',
            'Project Category', 'Risk Owner', 'Likelihood_Pre', 'Impact_Pre',
            'Risk_Priority_Pre', 'Mitigating_Action', 'Likelihood_Post',
            'Impact_Post', 'Risk_Priority_Post'
        ]
        
        actual_cols = [str(col) for col in df.columns]
        
        for std_field in standard_fields:
            mapped_col = column_mapping.get(std_field)
            
            if mapped_col:
                # Convert mapped_col to string if it's an integer column index
                mapped_col_str = str(mapped_col) if isinstance(mapped_col, (int, float)) else mapped_col
                
                # Fuzzy Column Matching: In case the LLM hallucinates slight column name alterations
                close_matches = difflib.get_close_matches(mapped_col_str, actual_cols, n=1, cutoff=0.8)
                if mapped_col_str in actual_cols:
                    # Find the original column name to access DataFrame
                    original_col = None
                    for col in df.columns:
                        if str(col) == mapped_col_str:
                            original_col = col
                            break
                    if original_col is not None:
                        standardized_data[std_field] = df[original_col]
                elif close_matches:
                    print(f"  Fuzzy match applied: mapped '{mapped_col}' -> found '{close_matches[0]}'")
                    # Find the original column name for the fuzzy match
                    original_col = None
                    for col in df.columns:
                        if str(col) == close_matches[0]:
                            original_col = col
                            break
                    if original_col is not None:
                        standardized_data[std_field] = df[original_col]
                else:
                    standardized_data[std_field] = None
            else:
                standardized_data[std_field] = None
        
        result_df = pd.DataFrame(standardized_data)
        result_df = self.clean_and_standardize(result_df, column_mapping)
        
        return result_df

    def clean_and_standardize(self, df: pd.DataFrame, column_mapping: Dict[str, str]) -> pd.DataFrame:
        """Clean and standardize the extracted data."""
        scoring_scale = column_mapping.get('scoring_scale', '1-10')
        
        if 'Likelihood_Pre' in df.columns and df['Likelihood_Pre'] is not None:
            df['Likelihood_Pre'] = self.normalize_score(df['Likelihood_Pre'], scoring_scale)
        
        if 'Impact_Pre' in df.columns and df['Impact_Pre'] is not None:
            df['Impact_Pre'] = self.normalize_score(df['Impact_Pre'], scoring_scale)
        
        if 'Likelihood_Post' in df.columns and df['Likelihood_Post'] is not None:
            df['Likelihood_Post'] = self.normalize_score(df['Likelihood_Post'], scoring_scale)
        
        if 'Impact_Post' in df.columns and df['Impact_Post'] is not None:
            df['Impact_Post'] = self.normalize_score(df['Impact_Post'], scoring_scale)
        
        if df['Risk_Priority_Pre'].isna().all() or df['Risk_Priority_Pre'] is None:
            df['Risk_Priority_Pre'] = self.calculate_risk_priority(
                df.get('Likelihood_Pre'), df.get('Impact_Pre')
            )
        
        if column_mapping.get('has_pre_post_separate'):
            if df['Risk_Priority_Post'].isna().all() or df['Risk_Priority_Post'] is None:
                df['Risk_Priority_Post'] = self.calculate_risk_priority(
                    df.get('Likelihood_Post'), df.get('Impact_Post')
                )
        
        text_columns = ['Risk Description', 'Mitigating_Action', 'Project Stage', 
                        'Project Category', 'Risk Owner']
        for col in text_columns:
            if col in df.columns and df[col] is not None:
                df[col] = df[col].apply(lambda x: self.clean_text(x) if pd.notna(x) else x)
        
        if df['Risk ID'].isna().all() or df['Risk ID'] is None:
            df['Risk ID'] = range(1, len(df) + 1)
        
        return df

    @staticmethod
    def normalize_score(series: pd.Series, scale: str) -> pd.Series:
        """Normalize scores to 1-10 scale safely parsing text values."""
        if series is None:
            return None
            
        def extract_number(val):
            if pd.isna(val):
                return pd.NA
            
            # If it's already a number
            if isinstance(val, (int, float)):
                return val
                
            # If it's a string, try to find the first integer. e.g "3 - High Risk" -> 3
            val_str = str(val).strip()
            match = re.search(r'\b(\d+)\b', val_str)
            if match:
                return float(match.group(1))
            
            # Fallback if no digit found
            return pd.NA
            
        # Apply intelligent extraction
        series = series.apply(extract_number)
        
        # Coerce to numeric (safely handles pd.NA now)
        series = pd.to_numeric(series, errors='coerce')
        
        if scale == '1-5':
            return series * 2
        elif scale == '1-10':
            return series
        else:
            max_val = series.dropna().max() if not series.dropna().empty else 10
            if pd.notna(max_val) and max_val <= 5:
                return series * 2
            return series

    @staticmethod
    def calculate_risk_priority(likelihood: Optional[pd.Series], 
                                impact: Optional[pd.Series]) -> pd.Series:
        """Calculate risk priority (Low/Med/High) from scores."""
        if likelihood is None or impact is None:
            return pd.Series(['Unknown'] * len(likelihood) if likelihood is not None else 0)
        
        score = likelihood * impact
        
        def priority_label(s):
            if pd.isna(s):
                return 'Unknown'
            elif s <= 30:
                return 'Low'
            elif s <= 60:
                return 'Med'
            else:
                return 'High'
        
        return score.apply(priority_label)

    @staticmethod
    def clean_text(text: str) -> str:
        """Clean text by removing extra whitespace."""
        if not isinstance(text, str):
            return text
        
        text = text.replace('\\n', ' ').replace('\n', ' ')
        text = re.sub(r'\s+', ' ', text)
        text = text.strip()
        
        return text

    def create_output_excel(self, df: pd.DataFrame, output_path: str, column_mapping: Dict[str, str]):
        """Create standardized output Excel with two sheets."""
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        
        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Simplified Register"
        
        has_pre_post = column_mapping.get('has_pre_post_separate', False)
        
        if has_pre_post:
            header_row = [
                'Date Added', 'Risk ID', 'Risk Description', 'Project Stage',
                'Project Category', 'Risk Owner', 
                'Likelihood (1-10) (pre-mitigation)', 'Impact (1-10) (pre-mitigation)',
                'Risk Priority (pre-mitigation)', 'Mitigating Action',
                'Likelihood (1-10) (post-mitigation)', 'Impact (1-10) (post-mitigation)',
                'Risk Priority (post-mitigation)'
            ]
            data_cols = [
                'Date Added', 'Risk ID', 'Risk Description', 'Project Stage',
                'Project Category', 'Risk Owner', 'Likelihood_Pre', 'Impact_Pre',
                'Risk_Priority_Pre', 'Mitigating_Action', 'Likelihood_Post',
                'Impact_Post', 'Risk_Priority_Post'
            ]
        else:
            header_row = [
                'Date Added', 'Risk ID', 'Risk Description', 'Project Stage',
                'Project Category', 'Risk Owner', 'Likelihood (1-10)',
                'Impact (1-10)', 'Risk Priority (low, med, high)', 'Mitigating Action'
            ]
            data_cols = [
                'Date Added', 'Risk ID', 'Risk Description', 'Project Stage',
                'Project Category', 'Risk Owner', 'Likelihood_Pre', 'Impact_Pre',
                'Risk_Priority_Pre', 'Mitigating_Action'
            ]
        
        for col_idx, header in enumerate(header_row, start=1):
            cell = ws1.cell(row=1, column=col_idx, value=header)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color='D3D3D3', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', wrap_text=True)
        
        for row_idx, (_, row) in enumerate(df.iterrows(), start=2):
            for col_idx, col_name in enumerate(data_cols, start=1):
                value = row.get(col_name)
                if pd.notna(value):
                    ws1.cell(row=row_idx, column=col_idx, value=value)
        
        for col_idx in range(1, len(header_row) + 1):
            ws1.column_dimensions[ws1.cell(row=1, column=col_idx).column_letter].width = 15
        
        ws2 = wb.create_sheet("Output Requirements")
        
        requirements_text = [
            ["The following columns are mandatory for the final risk registers. Additional columns may be added, if information is provided in the input files (ex. date, additional comments etc):"],
            ["Risk ID", "If not provided, any identifier may be used."],
            ["Risk Description", ""],
            ["Project Stage", "Required for construction or project based risks."],
            ["Project Category", ""],
            ["Risk Owner", ""],
            ["Mitigating Action", ""],
            ["Likelihood (1-10)", "If multiple stages of risk assessment are provided, include both pre and post-mitigation."],
            ["Impact (1-10)", ""],
            ["Risk Priority (low, med, high)", ""],
        ]
        
        for row_idx, row_data in enumerate(requirements_text, start=1):
            for col_idx, value in enumerate(row_data, start=1):
                ws2.cell(row=row_idx, column=col_idx, value=value)
        
        wb.save(output_path)

    def process_file(self, input_file: str, output_file: str) -> bool:
        """Main processing function."""
        print(f"\nProcessing: {input_file}")
        print("=" * 80)
        
        print("Step 1: Loading file...")
        df, file_type = self.load_risk_register(input_file)
        
        if df is None or df.empty:
            print(f"Error: Could not load data from {input_file}")
            return False
        
        print(f"  Loaded {df.shape[0]} rows × {df.shape[1]} columns")
        
        print("Step 2: Identifying columns with LLM...")
        column_mapping = self.identify_columns_with_llm(df)
        print(f"  Column mapping complete")
        
        print("Step 3: Extracting and standardizing...")
        standardized_df = self.extract_risk_data_with_llm(df, column_mapping)
        print(f"  Extracted {len(standardized_df)} risk records")
        
        print("Step 4: Creating output Excel...")
        self.create_output_excel(standardized_df, output_file, column_mapping)
        print(f"  ✓ Output saved to: {output_file}")
        
        return True


def main():
    """Main entry point."""
    input_dir = Path("input")
    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)
    
    input_files = list(input_dir.glob("*"))
    
    if not input_files:
        print("No files found in input directory")
        return
    
    print(f"Found {len(input_files)} files to process")
    
    try:
        standardizer = RiskRegisterStandardizer()
    except ValueError as e:
        print(f"Error initializing standardizer: {e}")
        return
        
    for input_file in input_files:
        if input_file.suffix.lower() not in ['.xlsx', '.xls', '.pdf']:
            continue
        
        output_name = input_file.stem.replace("__Input_", "__Final_") + ".xlsx"
        output_path = output_dir / output_name
        
        try:
            success = standardizer.process_file(str(input_file), str(output_path))
            if not success:
                print(f"Warning: Failed to process {input_file.name}")
        except Exception as e:
            print(f"Error processing {input_file.name}: {e}")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    main()
