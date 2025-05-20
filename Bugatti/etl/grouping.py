import pandas as pd
from datetime import datetime
import calendar

def get_apply_months_and_days(start, end):
    """
    Calculate the months and days covered by a date range.
    
    Args:
        start: Start date in YYYYMMDD format
        end: End date in YYYYMMDD format
        
    Returns:
        List of tuples containing (month_string, days_in_month)
    """
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

def group_similar_rows(extracted_df):
    """
    Group similar rows using the original logic.
    
    Args:
        extracted_df: DataFrame with extracted data
        
    Returns:
        DataFrame with grouped rows
    """
    # Make a copy to avoid modifying the original
    extracted_df = extracted_df.copy()
    
    # Define grouping columns
    group_cols = [
        'Customer Name', 'Customer Code', 'Model Code', 
        'Start Date', 'End Date', 'Additional SOA', 
        'Source File', 'Name of Promotion'
    ]
    
    # Add 'Is WBW' to grouping if it exists
    if 'Is WBW' in extracted_df.columns:
        group_cols.append('Is WBW')
    
    # Make sure all group_cols exist
    for col in group_cols:
        if col not in extracted_df.columns:
            print(f"⚠️ Missing column for grouping: {col}")
            extracted_df[col] = 'NA'  # Add placeholder
    
    # Define aggregation functions
    agg_dict = {col: 'first' for col in extracted_df.columns if col not in group_cols}
    agg_dict['Expected Sell-Out'] = 'sum'  # Sum quantities
    
    # Print before grouping stats
    print(f"➡️ Before grouping: {len(extracted_df)} rows, {extracted_df['Expected Sell-Out'].sum()} units")
    
    # Perform the grouping
    grouped_df = extracted_df.groupby(group_cols, as_index=False).agg(agg_dict)
    
    # Print after grouping stats
    print(f"➡️ After grouping: {len(grouped_df)} rows, {grouped_df['Expected Sell-Out'].sum()} units")
    
    return grouped_df

def distribute_quantities_by_month(grouped_df):
    """
    Expands a grouped DataFrame by apply month and distributes quantities.
    Handles special cases for 1–5 units and uses proportional logic for larger quantities.
    
    Args:
        grouped_df: Grouped DataFrame with Start Date and End Date columns
        
    Returns:
        Expanded DataFrame with apply month and distributed quantities
    """
    expanded_rows = []
    
    for _, row in grouped_df.iterrows():
        start, end = row['Start Date'], row['End Date']
        try:
            months = get_apply_months_and_days(start, end)
            if not months:
                row_copy = row.copy()
                row_copy['Apply Month'] = 'NA'
                row_copy['Errors in Combined Extract'] = 'Could not calculate apply months'
                expanded_rows.append(row_copy)
                continue

            total_days = sum(d for _, d in months)
            qty = int(row['Expected Sell-Out'])
            month_count = len(months)

            # Distribution logic
            dist_qty = []

            if qty == 1:
                dist_qty = [1] * month_count
            elif qty == 2:
                if month_count == 1:
                    dist_qty = [2]
                elif month_count == 2:
                    dist_qty = [1, 1]
                else:
                    dist_qty = [1] * min(qty, month_count) + [0] * (month_count - qty)
            elif qty == 3:
                if month_count == 1:
                    dist_qty = [3]
                elif month_count == 2:
                    dist_qty = [2, 1]
                else:
                    dist_qty = [1] * 3 + [0] * (month_count - 3)
            elif qty == 4:
                if month_count >= 4:
                    dist_qty = [1] * 4 + [0] * (month_count - 4)
                else:
                    dist_qty = [0] * month_count
                    month_indices = sorted(range(month_count), key=lambda i: months[i][1], reverse=True)
                    for i in range(qty):
                        dist_qty[month_indices[i % month_count]] += 1
            elif qty == 5:
                if month_count >= 5:
                    dist_qty = [1] * 5 + [0] * (month_count - 5)
                else:
                    dist_qty = [0] * month_count
                    month_indices = sorted(range(month_count), key=lambda i: months[i][1], reverse=True)
                    for i in range(qty):
                        dist_qty[month_indices[i % month_count]] += 1
            else:
                # For larger quantities, distribute proportionally based on days
                dist_qty = []
                remaining = qty
                for i, (_, days) in enumerate(months[:-1]):
                    part = round(qty * days / total_days)
                    dist_qty.append(part)
                    remaining -= part
                dist_qty.append(max(0, remaining))  # ensure total matches qty

            # Create new rows
            for (month, _), qty_month in zip(months, dist_qty):
                new_row = row.copy()
                new_row['Apply Month'] = month
                new_row['Expected Sell-Out'] = qty_month
                new_row['Errors in Combined Extract'] = ''
                expanded_rows.append(new_row)

        except Exception as e:
            row_copy = row.copy()
            row_copy['Apply Month'] = 'NA'
            row_copy['Errors in Combined Extract'] = f"Expansion error: {e}"
            expanded_rows.append(row_copy)

    if expanded_rows:
        expanded_df = pd.DataFrame(expanded_rows)
        return expanded_df.sort_values(by='Original Row Index').reset_index(drop=True)
    else:
        return pd.DataFrame()
