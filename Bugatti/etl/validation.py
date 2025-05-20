# Functions for validating rows and detecting errors
import pandas as pd

def detect_errors(row):
    """
    Validate a row and detect common errors.
    
    Args:
        row: Dictionary or DataFrame row to validate
        
    Returns:
        String with comma-separated error messages
    """
    errors = []

    if row.get('Customer Type', '') == 'NA':
        errors.append('Missing Customer Type')
        
    if row.get('Requestor', '') == 'NA':
        errors.append('Missing Requestor')
        
    if row.get('Currency', '') == 'NA':
        errors.append('Missing Currency')
        
    if str(row.get('Start Date', '')).startswith('NA'):
        errors.append('Invalid Start Date')
        
    if str(row.get('End Date', '')).startswith('NA'):
        errors.append('Invalid End Date')
        
    if pd.isna(row.get('Model Code')) or str(row.get('Model Code')).strip() in ['NA', '']:
        errors.append('Missing Model Code')
        
    if pd.isna(row.get('Additional SOA')):
        errors.append('Missing Additional SOA')
        
    if pd.isna(row.get('Expected Sell-Out')):
        errors.append('Missing Expected Sell-Out')

    return ', '.join(errors) if errors else ''

def safe_get(row, colname):
    """
    Safely get a value from a row, returning 'NA' for missing values.
    
    Args:
        row: Dictionary or DataFrame row
        colname: Column name to retrieve
        
    Returns:
        Value or 'NA' if missing
    """
    value = row.get(colname, 'NA')
    if pd.isna(value) or str(value).strip() == '':
        return 'NA'
    return value