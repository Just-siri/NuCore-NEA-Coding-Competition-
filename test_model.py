"""
Test script for risk register standardization model
Tests against the 3 training examples
"""

import sys
import os
from pathlib import Path
import pandas as pd
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent))

from model import RiskRegisterStandardizer

def compare_outputs(generated_file, expected_file):
    """Compare generated output with expected output"""
    
    print(f"\nComparing:")
    print(f"  Generated: {generated_file}")
    print(f"  Expected:  {expected_file}")
    
    # Read both files
    df_gen = pd.read_excel(generated_file, sheet_name='Simplified Register')
    df_exp = pd.read_excel(expected_file, sheet_name='Simplified Register')
    
    print(f"\n  Generated shape: {df_gen.shape}")
    print(f"  Expected shape:  {df_exp.shape}")
    
    # Compare columns
    gen_cols = set(df_gen.columns)
    exp_cols = set(df_exp.columns)
    
    missing_cols = exp_cols - gen_cols
    extra_cols = gen_cols - exp_cols
    
    if missing_cols:
        print(f"  ⚠ Missing columns: {missing_cols}")
    if extra_cols:
        print(f"  ⚠ Extra columns: {extra_cols}")
    
    if not missing_cols and not extra_cols:
        print(f"  ✓ Column match!")
    
    # Check for mandatory fields
    mandatory = [
        'Risk ID', 'Risk Description', 'Project Stage', 'Project Category',
        'Risk Owner', 'Mitigating Action', 'Likelihood (1-10)', 
        'Impact (1-10)', 'Risk Priority (low, med, high)'
    ]
    
    missing_mandatory = [col for col in mandatory if col not in df_gen.columns]
    if missing_mandatory:
        print(f"  ✗ Missing mandatory columns: {missing_mandatory}")
    else:
        print(f"  ✓ All mandatory columns present")
    
    # Show sample comparison
    print(f"\n  Sample row 1 comparison:")
    gen_row = df_gen.iloc[0].to_dict() if len(df_gen) > 0 else "No data"
    exp_row = df_exp.iloc[0].to_dict() if len(df_exp) > 0 else "No data"
    print(f"    Generated: {gen_row}")
    print(f"    Expected:  {exp_row}")
    
    # Value level matches (Exclude optional/null generated fields in the base files)
    total_cells = 0
    matched_cells = 0
    
    for row_idx in range(min(len(df_gen), len(df_exp))):
        for col in mandatory:
            if col in df_gen.columns and col in df_exp.columns:
                gen_val = df_gen.iloc[row_idx][col]
                exp_val = df_exp.iloc[row_idx][col]
                
                # Check NaNs mapping
                if pd.isna(gen_val) and pd.isna(exp_val):
                    matched_cells += 1
                elif str(gen_val).strip() == str(exp_val).strip():
                    matched_cells += 1
                elif isinstance(gen_val, (int, float)) and isinstance(exp_val, (int, float)) and pd.notna(gen_val) and pd.notna(exp_val):
                   if abs(float(gen_val) - float(exp_val)) < 0.1:
                       matched_cells += 1
                total_cells += 1
                
    cell_match_rate = (matched_cells / total_cells * 100) if total_cells > 0 else 0
    print(f"  ✓ Value Match Rate (Mandatory Fields): {cell_match_rate:.2f}% ({matched_cells}/{total_cells})")
    
    return {
        'columns_match': not missing_cols and not extra_cols,
        'has_mandatory': not missing_mandatory,
        'row_count_match': df_gen.shape[0] == df_exp.shape[0],
        'value_match_rate': cell_match_rate
    }


def main():
    """Run tests on training data"""
    
    # Check for API key
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        print("Error: ANTHROPIC_API_KEY environment variable not set")
        print("Set it with: export ANTHROPIC_API_KEY='your-key-here'")
        sys.exit(1)
    
    print("="*80)
    print("RISK REGISTER STANDARDIZATION MODEL - TESTING")
    print("="*80)
    
    # Set up paths
    uploads_dir = Path(__file__).parent / 'input'
    test_output_dir = Path(__file__).parent / 'test_outputs'
    test_output_dir.mkdir(exist_ok=True)
    
    # Define test cases (training data)
    test_cases = [
        {
            'name': 'Test 1: IVC DOE',
            'input': uploads_dir / '1. IVC DOE R2 (Input).xlsx',
            'expected': uploads_dir / '1. IVC DOE R2 (Input).xlsx',  # Using same file for now
            'output': test_output_dir / '1__IVC_DOE__Test_Output.xlsx'
        },
        {
            'name': 'Test 2: City of York Council',
            'input': uploads_dir / '2. City of York Council (Input).xlsx',
            'expected': uploads_dir / '2. City of York Council (Input).xlsx',  # Using same file for now
            'output': test_output_dir / '2__City_of_York_Council__Test_Output.xlsx'
        },
        {
            'name': 'Test 3: Digital Security IT',
            'input': uploads_dir / '3. Digital Security IT Sample Register (Input).xlsx',
            'expected': uploads_dir / '3. Digital Security IT Sample Register (Input).xlsx',  # Using same file for now
            'output': test_output_dir / '3__Digital_Security_IT__Test_Output.xlsx'
        }
    ]
    
    # Initialize standardizer
    standardizer = RiskRegisterStandardizer(api_key=api_key)
    
    results = []
    
    # Run tests
    for test in test_cases:
        print(f"\n{'='*80}")
        print(f"{test['name']}")
        print(f"{'='*80}")
        
        try:
            # Process file
            print(f"Processing: {test['input'].name}")
            standardizer.process_file(str(test['input']), str(test['output']))
            
            # Compare outputs
            comparison = compare_outputs(test['output'], test['expected'])
            
            results.append({
                'test': test['name'],
                'status': 'PASS' if all(comparison.values()) else 'PARTIAL',
                'details': comparison
            })
            
        except Exception as e:
            print(f"✗ ERROR: {e}")
            import traceback
            traceback.print_exc()
            results.append({
                'test': test['name'],
                'status': 'FAIL',
                'error': str(e)
            })
    
    # Summary
    print(f"\n{'='*80}")
    print("TEST SUMMARY")
    print(f"{'='*80}")
    
    for result in results:
        status_symbol = '✓' if result['status'] == 'PASS' else ('⚠' if result['status'] == 'PARTIAL' else '✗')
        print(f"{status_symbol} {result['test']}: {result['status']}")
        if 'details' in result:
            details = result['details']
            print(f"    Columns match: {details['columns_match']}")
            print(f"    Has mandatory: {details['has_mandatory']}")
            print(f"    Row count match: {details['row_count_match']}")
            print(f"    Value Match Rate: {details.get('value_match_rate', 0.0):.2f}%")
            
    # Calculate overall accuracy
    overall_match = sum(r['details'].get('value_match_rate', 0.0) for r in results if 'details' in r)
    total_valid_tests = sum(1 for r in results if 'details' in r)
    
    if total_valid_tests > 0:
        print(f"\n★ OVERALL ACCURACY (Value Match Rate): {(overall_match / total_valid_tests):.2f}%")
    
    print(f"\nTest outputs saved to: {test_output_dir}")


if __name__ == '__main__':
    main()
