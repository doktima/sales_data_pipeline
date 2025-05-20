# Functions for loading and cleaning Excel files
import pandas as pd
import os
import glob
import sys

from config.constants import EXPECTED_KEYWORDS, COLUMN_MAPPING_DF_CONFIG
from utils.fuzzy_match import find_header_row, clean_column_name, get_single_fuzzy_match, fuzzy_match_columns

def init_column_mapping_df():
    """Create the column mapping DataFrame from configuration."""
    df = pd.DataFrame(COLUMN_MAPPING_DF_CONFIG)
    
    # Clean variations in mapping table
    df['Possible Variations'] = df['Possible Variations'].apply(
        lambda lst: [clean_column_name(x) for x in lst]
    )
    
    return df

def load_and_clean_excel(filepath, expected_keywords=EXPECTED_KEYWORDS, threshold=85):
    """
    Load an Excel file and clean it by finding the header row and standardizing column names.
    
    Args:
        filepath: Path to the Excel file
        expected_keywords: Keywords to detect header row
        threshold: Fuzzy matching threshold
        
    Returns:
        Cleaned DataFrame or None if processing failed
    """
    column_mapping_df = init_column_mapping_df()
    
    try:
        # Try to find the correct sheet
        xl = pd.ExcelFile(filepath)
        expected_sheet_keywords = ['pet form', 'spgm request', 'av spgm']
        matching_sheets = [s for s in xl.sheet_names if any(keyword in s.lower() for keyword in expected_sheet_keywords)]
        
        if not matching_sheets:
            print(f"No matching sheet found in {os.path.basename(filepath)}.")
            return None
            
        sheet_to_use = matching_sheets[0]
        print(f"Reading sheet: {sheet_to_use}")
        raw_df = pd.read_excel(filepath, sheet_name=sheet_to_use, header=None)
    except Exception as e:
        print(f"Error reading {filepath}: {e}")
        return None

    # Find the header row
    main_header_row = find_header_row(raw_df, expected_keywords, max_rows_to_check=13)
    if main_header_row is None:
        print(f"No suitable header found in {os.path.basename(filepath)}.")
        return None
        
    alt_header_row = main_header_row + 1 if main_header_row + 1 < len(raw_df) else None

    # Get header values from main and alternative header rows
    main_header = raw_df.iloc[main_header_row].astype(str).tolist()
    alt_header = raw_df.iloc[alt_header_row].astype(str).tolist() if alt_header_row is not None else [None] * len(main_header)

    # Clean header values
    main_header = [clean_column_name(val) for val in main_header]
    alt_header = [clean_column_name(val) for val in alt_header]

    # Ensure both headers have the same length
    max_len = max(len(main_header), len(alt_header))
    main_header += [None] * (max_len - len(main_header))
    alt_header += [None] * (max_len - len(alt_header))

    # Combine both headers, preferring valid matches
    final_header = []
    alt_used_count = 0
    for i in range(max_len):
        main_val = main_header[i] if main_header[i] else ""
        alt_val = alt_header[i] if alt_header[i] else ""
        
        if get_single_fuzzy_match(main_val, column_mapping_df, threshold):
            final_header.append(main_val)
        elif get_single_fuzzy_match(alt_val, column_mapping_df, threshold):
            final_header.append(alt_val)
            alt_used_count += 1
        else:
            final_header.append(main_val)

    # Determine which rows to drop
    rows_to_drop = alt_header_row if alt_header_row is not None and alt_used_count > 0 else main_header_row
    
    # Create the cleaned DataFrame
    df = raw_df.iloc[rows_to_drop + 1:].reset_index(drop=True)
    df.columns = final_header[:len(df.columns)]
    
    # Standardize column names
    df = fuzzy_match_columns(df, column_mapping_df, threshold=threshold)
    
    return df