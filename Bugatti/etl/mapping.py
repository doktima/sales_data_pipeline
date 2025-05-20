# Functions for mapping metadata and handling suffix rules
from config.constants import (
    SALES_PGM_REASON_VARIATIONS,
    DIVISION_PREFIX_MAP,
    DIVISION_BUDGET_ALLOCATIONS,
    AV_SUFFIXES,
    TV_SUFFIXES
)

def map_all_promo_metadata(model_code, support_input):
    """
    Map model codes and support types to standardized metadata.
    
    Args:
        model_code: Product model code
        support_input: Type of support
        
    Returns:
        Tuple of (budget_allocation, product_type, reason_code, pgm_type)
    """
    try:
        model_code = str(model_code).strip().upper()
        support_input = str(support_input).strip().upper()

        # Step 1: Identify Budget Allocation
        if any(model_code.endswith(suffix) for suffix in AV_SUFFIXES):
            budget_allocation = "PNT"
        elif any(model_code.endswith(suffix) for suffix in TV_SUFFIXES):
            budget_allocation = "GLT"
        elif model_code in DIVISION_BUDGET_ALLOCATIONS:
            budget_allocation = model_code
        else:
            budget_allocation = "NA"
            for division, prefixes in DIVISION_PREFIX_MAP.items():
                if any(model_code.startswith(prefix) for prefix in prefixes):
                    budget_allocation = division
                    break

        # Step 2: Determine Product Type directly from Product Code content
        product_type = "Division" if any(div in model_code for div in DIVISION_BUDGET_ALLOCATIONS) else "Model"

        # Step 3: Map Reason Code and Sales PGM Type
        reason_code = "NA"
        pgm_type = "NA"
        
        # Special case for A SOA and similar values
        if support_input in ["A SOA", "SOA", "A-SOA"] or support_input.upper() == "SOA":
            reason_code = "TM_Z02"
            pgm_type = "Lumpsum"
        else:
            # Normal mapping logic
            for item in SALES_PGM_REASON_VARIATIONS:
                if support_input in item["variations"]:
                    reason_code = item["reason_code"]
                    pgm_type = item["pgm_type"]
                    break

        return budget_allocation, product_type, reason_code, pgm_type

    except Exception as e:
        # Print error for debugging
        print(f"Error in mapping: {e} - model_code: {model_code}, support_input: {support_input}")
        # fallback safe default tuple
        return "NA", "Model"

def classify_model_code(model_code):
    """
    Classify a model code as AV, TV, or UNKNOWN.
    
    Args:
        model_code: Product model code
        
    Returns:
        String classification ("AV", "TV", or "UNKNOWN")
    """
    if not isinstance(model_code, str):
        return "UNKNOWN"
        
    if any(model_code.endswith(sfx) for sfx in AV_SUFFIXES):
        return "AV"
        
    if any(model_code.endswith(sfx) for sfx in TV_SUFFIXES):
        return "TV"
        
    return "NA"