import pandas as pd
try:
    from rapidfuzz import process
except ImportError:
    # Fallback to fuzzywuzzy if rapidfuzz is not available
    from fuzzywuzzy import process

def clean_column_name(name):
    """Clean and normalize column names by removing special characters and whitespace."""
    if not isinstance(name, str):
        return ""
    return (
        name.replace('\n', ' ')
            .replace('\r', ' ')
            .replace('\u00A0', ' ')
            .strip()
            .lower()
    )

def find_header_row(df, expected_keywords, max_rows_to_check=13):
    """
    Find the most likely header row in a dataframe based on expected keywords.
    
    Args:
        df: DataFrame to search in
        expected_keywords: List of keywords to look for
        max_rows_to_check: Maximum number of rows to check from the top
        
    Returns:
        Index of the best header row or None if no match found
    """
    best_index = None
    best_score = 0
    rows_to_check = min(max_rows_to_check, len(df))
    
    for i in range(rows_to_check):
        row_vals = df.iloc[i].astype(str).str.lower().tolist()
        score = sum(
            any(keyword in cell for cell in row_vals)
            for keyword in expected_keywords
        )
        if score > best_score:
            best_score = score
            best_index = i
    
    return best_index if best_score > 0 else None

def get_single_fuzzy_match(test_value, column_mapping_df, threshold=85):
    """
    Check if a value matches any of the expected column variations.
    
    Args:
        test_value: Value to test
        column_mapping_df: DataFrame with column mapping configurations
        threshold: Minimum score for a match to be considered valid
        
    Returns:
        Boolean indicating whether there's a match
    """
    if not isinstance(test_value, str):
        return False
        
    test_value_cleaned = clean_column_name(test_value)

    for variations in column_mapping_df['Possible Variations']:
        if test_value_cleaned in variations:
            return True

    all_possible_names = []
    for variations in column_mapping_df['Possible Variations']:
        all_possible_names.extend(variations)

    result = process.extractOne(test_value_cleaned, all_possible_names)
    if result:
        if isinstance(result, tuple):
            # Handle different return formats from different libraries
            if len(result) >= 2:
                score = result[1]
                return score >= threshold
    return False

def fuzzy_match_columns(df, column_mapping_df, threshold=85):
    """
    Rename columns in a DataFrame based on fuzzy matching against expected variations.
    
    Args:
        df: DataFrame with columns to rename
        column_mapping_df: DataFrame with column mapping configurations
        threshold: Minimum score for a match to be considered valid
        
    Returns:
        DataFrame with renamed columns
    """
    current_cols = [clean_column_name(col) for col in df.columns]
    rename_map = {}

    for idx, row in column_mapping_df.iterrows():
        std_col = row['Standard Column']
        variations = row['Possible Variations']
        exclusions = row.get('Exclusion Variations', [])

        match = None
        for c in current_cols:
            if any(ex in c for ex in exclusions):
                continue
            if c in variations:
                match = c
                break

        if not match:
            best_score = 0
            best_match = None
            for variant in variations:
                result = process.extractOne(variant, current_cols)
                if result and isinstance(result, tuple) and len(result) >= 2:
                    matched_col, score = result[0], result[1]
                    # Check exclusion terms
                    if score >= threshold and not any(ex in matched_col for ex in exclusions):
                        if score > best_score:
                            best_score = score
                            best_match = matched_col
            if best_match:
                match = best_match

        if match:
            original_col_index = current_cols.index(match)
            original_col_name = df.columns[original_col_index]
            rename_map[original_col_name] = std_col

    df.rename(columns=rename_map, inplace=True)
    return df