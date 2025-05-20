SPMS_Registration_Structured/
│
├── config/                 # Configuration files
│   ├── paths.py           # Path definitions and environment setup
│   └── constants.py       # Constant definitions and mapping tables
│
├── etl/                    # Extract, Transform, Load functionality
│   ├── loader.py          # Functions for loading and cleaning Excel files
│   ├── parser.py          # Functions for parsing and correcting dates and values
│   ├── mapping.py         # Functions for mapping metadata and handling suffix rules
│   ├── grouping.py        # Functions for grouping and distributing quantities
│   └── validation.py      # Functions for validating rows and detecting errors
│
├── writers/                # Output generation functionality
│   ├── excel_writer.py    # Functions for saving Excel files and formatting
│   └── promo_naming.py    # Functions for building and formatting promotion names
│
├── utils/                  # Utility functions
│   └── fuzzy_match.py     # Helper functions for fuzzy matching and column cleaning
│
├── tests/                  # Unit tests
│   ├── test_loader.py     # Tests for loader functions
│   ├── test_parser.py     # Tests for parser functions
│   └── test_mapping.py    # Tests for mapping functions
│
├── data/                   # Sample data directory
│
└── main.py                 # Main entry point script