
# Constants used across scripts

# Sales PGM Reason Code mapping with all possible input variations
SALES_PGM_REASON_VARIATIONS = [
    {"reason_code": "TM_R03", "variations": ["TM_R03", "DISPLAY SUPPORT REBATE"], "pgm_type": "Lumpsum", "product_type": "Division"},
    {"reason_code": "TM_R12", "variations": ["TM_R12", "NTSI", "ADDITIONAL SELL IN REBATE"], "pgm_type": "Lumpsum", "product_type": "Division"},
    {"reason_code": "TM_C01", "variations": ["TM_C01", "CO-OP", "COOP", "CO-OP AD", "CO-OP AD.", "CO-OP ADVERTISING", "COP"], "pgm_type": "Lumpsum", "product_type": "Division"},
    {"reason_code": "TM_P01", "variations": ["TM_P01", "PRICE PROTECTION"], "pgm_type": "Lumpsum", "product_type": "Model"},
    {"reason_code": "TM_Z02", "variations": ["TM_Z02", "SOA", "SELL OUT SUPPORT REBATE", "A SOA"], "pgm_type": "Lumpsum", "product_type": "Model"},
]

# Division Prefix map expect GLT and PNT
DIVISION_PREFIX_MAP = {
    "CDT": ["DB", "DF"],
    "CNT": ["GB", "GM", "GS"],
    "DFT": ["F4", "FW", "FD", "F2", "FH", "LS", "WT", "S3", "W4"]
}

# Valid Budget Allocations = Divisions
DIVISION_BUDGET_ALLOCATIONS = {"GNT", "GJT", "GLT", "DFT", "CNT", "GTT", "PNT", "CDT", "PCT", "GKT"}

# AV and TV suffix logic
AV_SUFFIXES = {
    "PNT", "DGBRLLK", "CEUSCL2", ".ABEUBK", "AGBRLLK",
    ".ABEUWH", ".ABSWBK", ".ABSWWH", "AGBRLLX", "BGBRLLK",
    "EGBRLLK", "BGBRJJK", "ABEUWHF", "CGBRLLK", "AGBRLLZ",
    "CGBRLBI", "CGBRLBK", "CEUSLLK", "AEUSLLA", "AEUSLLB"
}

TV_SUFFIXES = {
    "GLT", ".AEK", "AEKQ", "AEKW", "AEKM", "AEKD"
}

HS_CODES = {"CDT", "CNT", "DFT"}

# Formatting constants
ABBREVIATIONS = {"AV", "TV", "SOA", "UK", "EE", "QNED", "OLED", "HDR", "UHD", "AI", "LG", "FOC", "WBW"}
REMOVE_WORDS = {"CIH", "EXRTIS"}

# Column mapping configuration
COLUMN_MAPPING_DF_CONFIG = {
    'Standard Column': [
        'Customer Code', 'Customer Name', 'Model Code',
        'Type of Support', 'Additional SOA', 'Expected Sell-Out',
        'Start Date', 'End Date', 'Expected Cost', 'Name of Promotion'
    ],
    'Possible Variations': [
        ['Customer Code', 'CustomerCode'],
        ['Customer Name', 'Account', 'CustomerName'],
        ['Model Code', 'Model.Suffix', 'Model', 'Product Code', 'SKU', 'Product'],
        ['Type Of Support'],
        ['SOA/Unit','SOA / unit','Additional SOA', 'DC/Unit',  'DC'],
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
        ['SOA', 'INVOICE BEFORE SOA', '① SOA'],  # Type of Support
        ['Invoice before SOA', 'Current SOA', 'Total SOA''① SOA'],  # Additional SOA
        ['Expected Sell-In', 'Sell-In Quantity'],  # Expected Sell-Out
        ['Request Date'],  # Start Date
        ['Request Date'],  # End Date
        ['Total SOA'],  # Expected Cost
        []   # Name of Promotion
    ]
}

# Expected keywords for detecting the header row
EXPECTED_KEYWORDS = ["customer", "account", "model", "sell", "soa", "code", "date"]

# Placeholder counter initialization
PLACEHOLDER_COUNTERS = {
    'Customer Code': 0,
    'Customer Name': 0,
    'Model Code': 0,
    'Name of Promotion': 0,
    'Start Date': 0,
    'End Date': 0,
    'Expected Sell-Out': 0,
    'Additional SOA': 0
}