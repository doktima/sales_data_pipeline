import sys
import os
import glob
import pandas as pd
from rapidfuzz import process
from fuzzywuzzy import process

# Add the project root directory to the Python path
script_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(script_dir)  # Go up one level from Bugatti
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# Import team configuration
from team_config import TEAM_MEMBER, PATHS

# Use the paths from configuration
source_folder = PATHS["pet_forms"]
output_folder = PATHS["uploads"]

# Print configuration info
print(f"Running script for team member: {TEAM_MEMBER}")
print(f"Source folder: {source_folder}")
print(f"Output folder: {output_folder}")

# Make sure the output directory exists
os.makedirs(output_folder, exist_ok=True)

# Define your "expected keywords" for detecting the header row
expected_keywords = ["customer", "account", "model", "sell", "soa", "code", "date"]

###############################################################################
# 2) Column Mapping Configuration
###############################################################################
column_mapping_df = pd.DataFrame({
    'Standard Column': [
        'Customer Code', 'Customer Name', 'Model Code',
        'Type of Support', 'Additional SOA', 'Expected Sell-Out',
        'Start Date', 'End Date', 'Expected Cost', 'Name of Promotion'
    ],
    'Possible Variations': [
        ['Customer Code', 'CustomerCode'],
        ['Customer Name', 'Account', 'CustomerName'],
        ['Model Code', 'Model.Suffix', 'Model', 'Product Code', 'SKU', 'Product'],
        ['TypeOfSupport', 'Type', 'Type of SOA'],
        ['SOA/Unit', 'Additional SOA', 'SOA', 'DC/Unit', 'SOA/unit', 'DC'],
        ['Expected Sell-Out', 'Sell-out Estimated QTY','Sell-Out Expected', 'Sell Out', 
         'Projected Sell', 'QTY', 'Quantity', 'Expected'],
        ['StartDate'],
        ['End Date'],
        ['Expected Cost', 'Total Additional Support AMT'],
        ['Name of promotion', 'Details', 'Comments']
    ],
    'Exclusion Variations': [
        [],  # Customer Code
        [],  # Customer Name
        [],  # Model Code
        ['SOA', 'INVOICE BEFORE SOA'],  # Type of Support
        ['Invoice before SOA', 'Current SOA', 'Total SOA'],  # ⚠️ Here we block misleading matches for Additional SOA
        ['Expected Sell-In', 'Sell-In Quantity'],  # Expected Sell-Out
        ['Request Date'],  # Start Date
        ['Request Date'],  # End Date
        ['Total SOA'],  # Expected Cost
        []   # Name of Promotion
    ]
})

###############################################################################
# 3) Helper Functions
###############################################################################
def clean_column_name(name):
    if not isinstance(name, str):
        return ""
    return (
        name.replace('\n', ' ')
            .replace('\r', ' ')
            .replace('\u00A0', ' ')
            .strip()
            .lower()
    )

# Clean variations in mapping table too
column_mapping_df['Possible Variations'] = column_mapping_df['Possible Variations'].apply(
    lambda lst: [clean_column_name(x) for x in lst]
)

def find_header_row(df, expected_keywords, max_rows_to_check=13):
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
    if not isinstance(test_value, str):
        return False
    test_value_cleaned = clean_column_name(test_value)

    for variations in column_mapping_df['Possible Variations']:
        if test_value_cleaned in variations:
            return True

    all_possible_names = []
    for variations in column_mapping_df['Possible Variations']:
        all_possible_names.extend(variations)

    best_match, score = process.extractOne(test_value_cleaned, all_possible_names)
    return score >= threshold

def fuzzy_match_columns(df, column_mapping_df, threshold=85):
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
                if result:
                    matched_col, score = result
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

def load_and_clean_excel(filepath, expected_keywords=expected_keywords, column_mapping_df=column_mapping_df, threshold=85):
    try:
        xl = pd.ExcelFile(filepath)
        expected_sheet_keywords = ['pet form', 'spgm request', 'av spgm']
        matching_sheets = [s for s in xl.sheet_names if any(keyword in s.lower() for keyword in expected_sheet_keywords)]
        if not matching_sheets:
            print(f"⚠️ No matching sheet found in {os.path.basename(filepath)}.")
            return None
        sheet_to_use = matching_sheets[0]
        print(f"📄 Reading sheet: {sheet_to_use}")
        raw_df = pd.read_excel(filepath, sheet_name=sheet_to_use, header=None)
    except Exception as e:
        print(f"❌ Error reading {filepath}: {e}")
        return None

    main_header_row = find_header_row(raw_df, expected_keywords, max_rows_to_check=13)
    if main_header_row is None:
        print(f"❌ No suitable header found in {os.path.basename(filepath)}.")
        return None
    alt_header_row = main_header_row + 1 if main_header_row + 1 < len(raw_df) else None

    main_header = raw_df.iloc[main_header_row].astype(str).tolist()
    alt_header = raw_df.iloc[alt_header_row].astype(str).tolist() if alt_header_row is not None else [None] * len(main_header)

    # Clean header values
    main_header = [clean_column_name(val) for val in main_header]
    alt_header = [clean_column_name(val) for val in alt_header]

    max_len = max(len(main_header), len(alt_header))
    main_header += [None] * (max_len - len(main_header))
    alt_header += [None] * (max_len - len(alt_header))

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

    rows_to_drop = alt_header_row if alt_header_row is not None and alt_used_count > 0 else main_header_row
    df = raw_df.iloc[rows_to_drop + 1:].reset_index(drop=True)
    df.columns = final_header[:len(df.columns)]
    df = fuzzy_match_columns(df, column_mapping_df, threshold=threshold)
    return df

###############################################################################
# 4) Main Execution
###############################################################################
excel_files = glob.glob(os.path.join(source_folder, "*.xlsx"))
if not excel_files:
    print("❌ No files found in source folder.")
    sys.exit()

print(f"📂 Found {len(excel_files)} files. Processing...")

for file_path in excel_files:
    print(f"\n📄 Processing: {os.path.basename(file_path)}")
    cleaned_df = load_and_clean_excel(file_path)
    if cleaned_df is None:
        continue
    print("🔎 Final columns ->", cleaned_df.columns.tolist())
    print(cleaned_df.head(3))
    output_file = os.path.join(output_folder, os.path.basename(file_path))
    try:
        cleaned_df.to_excel(output_file, engine='openpyxl', index=False)
        print(f"✅ Saved: {output_file}")
    except Exception as e:
        print(f"❌ Error saving {output_file}: {e}")

print("\n✅ Processing completed.")
