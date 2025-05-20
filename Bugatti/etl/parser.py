import re
import pandas as pd
import calendar
from datetime import datetime
from dateutil import parser
from decimal import Decimal, ROUND_HALF_UP

def parse_and_correct_date(val, is_start=True, start_reference=None):
    """
    Parse and standardize date values in various formats with improved regex handling.
    
    Args:
        val: Date value to parse
        is_start: Whether this is a start date (affects validation)
        start_reference: Reference start date for validating end dates
        
    Returns:
        Standardized date string in YYYYMMDD format or "NA" if invalid
    """
    if not val or str(val).strip().upper().startswith("NA"):
        return "NA"

    val_str = str(val).strip()
    today = datetime.today()
    
    # Common date formats with regex patterns - DD/Month/YYYY has highest priority
    formats = [
        # DD/Month/YYYY (Unambiguous format with text month - highest priority)
        (r'^(\d{1,2})[/\-\s]+(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[/\-\s]+(\d{4})$', '%d %B %Y'),
        
        # DD/MM/YYYY (Preferred format)
        (r'^(\d{1,2})/(\d{1,2})/(\d{4})$', '%d/%m/%Y'),
        # DD-MM-YYYY
        (r'^(\d{1,2})-(\d{1,2})-(\d{4})$', '%d-%m-%Y'),
        # YYYYMMDD (8 digits)
        (r'^(\d{4})(\d{2})(\d{2})$', '%Y%m%d'),
        # YYYY-MM-DD
        (r'^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$', '%Y-%m-%d'),
        # MM-DD-YYYY or MM/DD/YYYY
        (r'^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$', '%m-%d-%Y'),
        # DD.MM.YYYY
        (r'^(\d{1,2})\.(\d{1,2})\.(\d{4})$', '%d.%m.%Y'),
        # YYYY.MM.DD
        (r'^(\d{4})\.(\d{1,2})\.(\d{1,2})$', '%Y.%m.%d'),
        # Month name formats
        (r'^(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{4})$', '%d %b %Y'),
        (r'^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{1,2}),?\s+(\d{4})$', '%b %d %Y'),
    ]
    
    # Try regex patterns first
    parsed_dates = []
    
    # First check for the unambiguous format with named month
    for pattern, fmt in formats[:1]:  # Just the first format (DD/Month/YYYY)
        match = re.match(pattern, val_str, re.IGNORECASE)
        if match:
            try:
                groups = match.groups()
                day = int(groups[0])
                month_name = groups[1].lower()
                year = int(groups[2])
                
                # Convert month name to number
                month_map = {
                    'january': 1, 'jan': 1,
                    'february': 2, 'feb': 2,
                    'march': 3, 'mar': 3,
                    'april': 4, 'apr': 4,
                    'may': 5, 
                    'june': 6, 'jun': 6,
                    'july': 7, 'jul': 7,
                    'august': 8, 'aug': 8,
                    'september': 9, 'sep': 9,
                    'october': 10, 'oct': 10,
                    'november': 11, 'nov': 11,
                    'december': 12, 'dec': 12
                }
                
                month = month_map.get(month_name.lower(), 0)
                
                if 1 <= day <= 31 and 1 <= month <= 12 and 1900 <= year <= 2100:
                    if day <= calendar.monthrange(year, month)[1]:
                        parsed = datetime(year, month, day)
                        # This is unambiguous, so return immediately
                        return parsed.strftime("%Y%m%d")
            except:
                pass
    
    # First try the strict YYYYMMDD format
    numeric_val = re.sub(r"\D", "", val_str)
    if len(numeric_val) == 8 and 1900 <= int(numeric_val[:4]) <= 2100:
        try:
            year = int(numeric_val[:4])
            month = int(numeric_val[4:6])
            day = int(numeric_val[6:])
            
            # Validate month and day
            if 1 <= month <= 12 and 1 <= day <= 31:
                # Check if day is valid for the month
                if day <= calendar.monthrange(year, month)[1]:
                    parsed = datetime(year, month, day)
                    parsed_dates.append(parsed)
        except:
            pass
    
    # Try other regex patterns
    for pattern, fmt in formats[1:]:  # Skip the first format (already checked)
        match = re.match(pattern, val_str, re.IGNORECASE)
        if match:
            try:
                if fmt == '%d/%m/%Y' or fmt == '%d-%m-%Y' or fmt == '%d.%m.%Y' or fmt == '%d %b %Y':
                    # DD/MM/YYYY format
                    groups = match.groups()
                    day = int(groups[0])
                    month = groups[1] if len(groups) == 3 and isinstance(groups[1], str) and not groups[1].isdigit() else int(groups[1])
                    year = int(groups[2])
                    if isinstance(month, str):
                        month_map = {'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6, 
                                     'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12}
                        month = month_map.get(month.lower()[:3], 1)
                    if 1 <= day <= 31 and 1 <= month <= 12 and 1900 <= year <= 2100:
                        if day <= calendar.monthrange(year, month)[1]:
                            parsed = datetime(year, month, day)
                            parsed_dates.append(parsed)
                
                elif fmt == '%m-%d-%Y':
                    # MM-DD-YYYY format - evaluate context with today's date
                    groups = match.groups()
                    potential_month = int(groups[0])
                    potential_day = int(groups[1])
                    year = int(groups[2])
                    
                    # Check if both interpretations are valid dates
                    dd_mm_valid = (potential_month <= 31 and 1 <= potential_day <= 12 and 
                                   potential_day <= calendar.monthrange(year, potential_day)[1])
                    
                    mm_dd_valid = (1 <= potential_month <= 12 and 1 <= potential_day <= 31 and 
                                   potential_day <= calendar.monthrange(year, potential_month)[1])
                    
                    # Create both date interpretations if valid
                    dates_to_consider = []
                    
                    if dd_mm_valid:
                        try:
                            dd_mm_date = datetime(year, potential_day, potential_month)
                            # Calculate how far this date is from today
                            dd_mm_days_diff = abs((dd_mm_date - today).days)
                            dates_to_consider.append((dd_mm_date, dd_mm_days_diff, "DD/MM"))
                        except:
                            pass
                    
                    if mm_dd_valid:
                        try:
                            mm_dd_date = datetime(year, potential_month, potential_day)
                            # Calculate how far this date is from today
                            mm_dd_days_diff = abs((mm_dd_date - today).days)
                            dates_to_consider.append((mm_dd_date, mm_dd_days_diff, "MM/DD"))
                        except:
                            pass
                    
                    # Choose interpretation based on context
                    if dates_to_consider:
                        # First check if one interpretation is in the past and one is in the future
                        past_dates = [(d, diff, fmt_type) for d, diff, fmt_type in dates_to_consider if d < today]
                        future_dates = [(d, diff, fmt_type) for d, diff, fmt_type in dates_to_consider if d >= today]
                        
                        # For start dates, future dates are preferred
                        if is_start and future_dates:
                            # If we have future dates for start_date, prefer closest future date
                            closest_future = min(future_dates, key=lambda x: x[1])
                            parsed_dates.append(closest_future[0])
                            
                            # If we have both interpretations and DD/MM is a future date, add it with priority
                            dd_mm_future = [d for d, diff, fmt_type in future_dates if fmt_type == "DD/MM"]
                            if dd_mm_future and len(future_dates) > 1:
                                # If DD/MM format is in future dates, prioritize it
                                parsed_dates.insert(0, dd_mm_future[0])
                        
                        # If no future dates or not start date, use other heuristics
                        elif len(dates_to_consider) > 1:
                            # If both interpretations are valid, generally prefer DD/MM format
                            # But use temporal proximity as a deciding factor
                            
                            # Get the date difference for each interpretation
                            date_diffs = [(d, diff, fmt_type) for d, diff, fmt_type in dates_to_consider]
                            
                            # Sort by closeness to today's date
                            date_diffs.sort(key=lambda x: x[1])
                            
                            # If the difference between the closest date and next closest is significant
                            # (more than 30 days), choose the closest date
                            if len(date_diffs) >= 2 and date_diffs[0][1] + 30 < date_diffs[1][1]:
                                parsed_dates.append(date_diffs[0][0])
                            else:
                                # Otherwise default to DD/MM format with higher priority
                                dd_mm_dates = [d for d, diff, fmt_type in dates_to_consider if fmt_type == "DD/MM"]
                                if dd_mm_dates:
                                    parsed_dates.insert(0, dd_mm_dates[0])
                                
                                # Also add MM/DD interpretation as a fallback
                                mm_dd_dates = [d for d, diff, fmt_type in dates_to_consider if fmt_type == "MM/DD"]
                                if mm_dd_dates:
                                    parsed_dates.append(mm_dd_dates[0])
                        else:
                            # If only one interpretation is valid, add it
                            parsed_dates.append(dates_to_consider[0][0])
                    
                    # If we couldn't interpret either way, continue to next format
                    continue
                
                elif fmt == '%b %d %Y':
                    # Month name formats
                    groups = match.groups()
                    month_map = {'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6, 
                                 'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12}
                    month = month_map.get(groups[0].lower()[:3], 1)
                    day = int(groups[1])
                    year = int(groups[2])
                    if 1 <= day <= 31 and 1900 <= year <= 2100:
                        if day <= calendar.monthrange(year, month)[1]:
                            parsed = datetime(year, month, day)
                            parsed_dates.append(parsed)
                
                else:
                    # YYYY-MM-DD and YYYY.MM.DD formats
                    groups = match.groups()
                    year = int(groups[0])
                    month = int(groups[1])
                    day = int(groups[2])
                    if 1 <= month <= 12 and 1 <= day <= 31 and 1900 <= year <= 2100:
                        if day <= calendar.monthrange(year, month)[1]:
                            parsed = datetime(year, month, day)
                            parsed_dates.append(parsed)
            except:
                pass
    
    # If regex failed, fall back to dateutil parser
    if not parsed_dates:
        try:
            # Try dayfirst and monthfirst parsing
            try:
                parsed_dayfirst = parser.parse(val_str, dayfirst=True)
                if 1900 <= parsed_dayfirst.year <= 2100:
                    parsed_dates.append(parsed_dayfirst)
            except Exception:
                pass

            try:
                parsed_monthfirst = parser.parse(val_str, dayfirst=False)
                if 1900 <= parsed_monthfirst.year <= 2100 and parsed_monthfirst not in parsed_dates:
                    parsed_dates.append(parsed_monthfirst)
            except Exception:
                pass
        except:
            return "NA"
    
    # If we have no valid parsed dates, return NA
    if not parsed_dates:
        return "NA"
    
    # If we have multiple candidates, choose the most likely one
    if len(parsed_dates) > 1:
        if is_start:
            # For start dates, prefer dates that are today or in the future
            future_dates = [d for d in parsed_dates if d >= today]
            if future_dates:
                # Choose the closest future date
                parsed = min(future_dates, key=lambda d: abs((d - today).days))
            else:
                # If no future dates, choose today or the most recent past date
                parsed = max(parsed_dates)
        else:
            # For end dates, use start_reference if available
            if start_reference:
                try:
                    start_dt = datetime.strptime(start_reference, "%Y%m%d")
                    # Choose the closest date after start_reference
                    valid_dates = [d for d in parsed_dates if d >= start_dt]
                    if valid_dates:
                        # Prefer end dates within a reasonable timeframe (4 months)
                        reasonable_dates = [d for d in valid_dates if (d - start_dt).days <= 120]
                        if reasonable_dates:
                            parsed = min(reasonable_dates, key=lambda d: (d - start_dt).days)
                        else:
                            parsed = min(valid_dates, key=lambda d: (d - start_dt).days)
                    else:
                        # If no valid dates after start, choose the closest one
                        parsed = min(parsed_dates, key=lambda d: abs((d - start_dt).days))
                except:
                    parsed = max(parsed_dates)
            else:
                # Without start reference, choose the latest date
                parsed = max(parsed_dates)
    else:
        parsed = parsed_dates[0]
    
    # Ensure start date is not in the past
    if is_start and parsed < today:
        parsed = today
    
    # Return in standard format
    return parsed.strftime("%Y%m%d")

def get_apply_months_and_days(start, end):
    """
    Calculate the months and days covered by a date range.
    
    Args:
        start: Start date in YYYYMMDD format
        end: End date in YYYYMMDD format
        
    Returns:
        List of tuples containing (month_string, days_in_month)
    """
    from datetime import datetime
    import calendar
    
    months = []
    try:
        # Input validation
        if not isinstance(start, str) or not isinstance(end, str):
            return months
            
        if len(start) != 8 or len(end) != 8:
            return months
            
        # Parse dates
        start_dt = datetime.strptime(start, '%Y%m%d')
        end_dt = datetime.strptime(end, '%Y%m%d')
        
        # Validate date order
        if start_dt > end_dt:
            return months
            
        # Begin processing months
        current = start_dt.replace(day=1)
        
        while current <= end_dt:
            month_str = current.strftime('%Y%m')
            first_day = current
            last_day = datetime(current.year, current.month, 
                            calendar.monthrange(current.year, current.month)[1])
            
            range_start = max(start_dt, first_day)
            range_end = min(end_dt, last_day)
            days_in_month = (range_end - range_start).days + 1
            
            if days_in_month > 0:
                months.append((month_str, days_in_month))
            
            # Move to next month
            if current.month == 12:
                current = current.replace(year=current.year + 1, month=1)
            else:
                current = current.replace(month=current.month + 1)
                
    except Exception as e:
        print(f"Error calculating months: {e}")
        
    return months

def standardize_customer_code(val):
    """
    Standardize a customer code to the correct format (always uppercase).
    Special handling for specific codes like OBSIDIAN and HEKEYINDY.
    
    Args:
        val: Customer code to standardize
        
    Returns:
        Standardized customer code or original value if not a customer code
    """
    if not isinstance(val, str):
        return val
        
    # Map of special codes that need specific formatting
    special_codes = {
        'obsidian': 'OBSIDIAN',
        'Obsidian': 'OBSIDIAN',
        'ObSiDiAn': 'OBSIDIAN',
        'hekeyindy': 'HEKEYINDY',
        'he key indy': 'HEKEYINDY',
        'he-key-indy': 'HEKEYINDY',
        'HE KEY INDY': 'HEKEYINDY',
        'HE-KEY-INDY': 'HEKEYINDY',
        'He Key Indy': 'HEKEYINDY'
    }
    
    # Check if it's one of the special codes
    if val.lower() in special_codes:
        return special_codes[val.lower()]
    
    # Otherwise return it in uppercase
    return val.upper()

def is_likely_customer_code(val):
    """
    Check if a value is likely to be a customer code based on patterns.
    Always converts to uppercase for consistency.
    
    Args:
        val: Value to check
        
    Returns:
        Boolean indicating whether the value matches customer code patterns
    """
    valid_code_prefixes = ('IE', 'GB', '50', 'OB', 'JE', 'GG', '55', 'OB', '11', 'JE', 
                          '18', 'SE', 'HE', 'RA', '74', '13', '10')
    exceptions = {'HETIER1', 'HETIER2', 'HEBNO', 'SEVENOAKS_AWE', 'HEKEYINDY', 
                 'RADIUS_CIH', 'OBSIDIAN', '50380042-S'}
    
    if not isinstance(val, str):
        return False
    
    # Create a standardized version of the code    
    val_upper = val.upper()
    
    # Check if it's a valid customer code
    is_valid = val_upper.startswith(valid_code_prefixes) or val_upper in exceptions
    
    return is_valid

def is_likely_customer_name(val):
    """
    Check if a value is likely to be a customer name.
    
    Args:
        val: Value to check
        
    Returns:
        Boolean indicating whether the value is likely a customer name
    """
    if not isinstance(val, str):
        return False
        
    return not is_likely_customer_code(val)