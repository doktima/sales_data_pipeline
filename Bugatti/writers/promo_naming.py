# Functions for building and formatting promotion names
import re
from config.constants import ABBREVIATIONS, REMOVE_WORDS, HS_CODES
from etl.validation import safe_get

def format_title_case(text):
    """
    Format text in title case, preserving abbreviations.
    
    Args:
        text: Text to format
        
    Returns:
        Text in title case with abbreviations preserved
    """
    words = re.split(r'(\s+)', text)
    formatted_words = []
    
    for word in words:
        clean_word = word.strip().upper()
        
        if clean_word in REMOVE_WORDS:
            continue
            
        if clean_word in ABBREVIATIONS:
            formatted_words.append(word.upper())
        else:
            formatted_words.append(word.capitalize())
            
    result = ''.join(formatted_words)
    result = re.sub(r'\s+', ' ', result).strip()
    
    return result

def build_name_of_promotion(row):
    """
    Build a standardized promotion name based on row data.
    
    Args:
        row: DataFrame row containing promotion data
        
    Returns:
        Formatted promotion name
    """
    customer = safe_get(row, "Customer Name")
    segment = safe_get(row, "Segment")
    promo = safe_get(row, "Name of Promotion")
    start = safe_get(row, "Start Date")
    end = safe_get(row, "End Date")
    budget_alloc = safe_get(row, "Budget Allocation")
    support_type = safe_get(row, "Type of Support")

    # Clean promo name by removing unwanted words
    promo_clean = " ".join(word for word in str(promo).split() if word.upper() not in REMOVE_WORDS)

    # Format differently based on budget allocation
    if budget_alloc in HS_CODES:
        if "PRM" not in str(promo).upper():
            full_promo = f"HS - PET - {budget_alloc} - {promo_clean} - {support_type} - {customer} - {start} TO {end}"
        else:
            full_promo = f"HS - PRM - {budget_alloc} - {promo_clean} - {support_type} - {customer} - {start} TO {end}"
    else:
        base = f"{customer} ({segment})"
        full_promo = f"{base} {promo_clean} PET {support_type} {start} TO {end}"

    # Normalize whitespace and convert to uppercase
    full_promo = re.sub(r'\s+', ' ', full_promo).strip().upper()
    
    return full_promo