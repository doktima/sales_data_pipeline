# PET Form Processor (Finished)

This project provides a streamlined solution for automating the validation, transformation, and metadata enrichment of PET Form Excel files submitted by Key Account Managers (KAMs). The system prepares the data for structured mass upload and ensures consistency in promotional inputs across teams.
**Disclaimer:** This project is a simplified and sanitized version of an internal automation tool developed for demonstration and educational purposes only.

---

## Project Structure

```
SPMS_Registration_Structured/
├── Bugatti/
│   ├── config/
│   ├── data/
│   ├── etl/
│   ├── tests/
│   ├── utils/
│   ├── writers/
│   ├── main.py
│   ├── requirements.txt
│   └── README.md
├── .gitignore
├── CustomerMapping.xlsx
└── run_pet_processor_portable.bat
```

---

## Key Features

- Automated ingestion and cleansing of PET Form Excel files  
- Detection and correction of mismatched or swapped `Customer Name` and `Customer Code` fields  
- Standardized date parsing, validation, and logical correction  
- Mapping of missing promotional metadata, including `Requestor`, `Currency`, and `Customer Type`  
- Identification and handling of WBW (Weekly Bonus Window) TV models  
- Grouping and time-based distribution of expected sell-out quantities  
- Outputs include:
  - `CombinedExtractedColumns.xlsx` containing structured, enriched data
  - `MassUpload.xlsx` file ready for pricing system import

---

## Getting Started

1. Clone the repository

```bash
git clone https://github.com/doktima/sales_data_pipeline.git
cd sales_data_pipeline/Bugatti
```

2. Set up the environment (optional if using the provided portable Python)

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

Alternatively, use:

```bash
run_pet_processor_portable.bat
```

---

## Required Files

- `CustomerMapping.xlsx` should be present in the root directory  
- Source PET Form `.xlsx` files should be placed in the path defined by `PATHS['pet_forms']` in `config/paths.py`

---

## How It Works

Run:

```bash
python main.py
```

The process will:

- Load and clean Excel input files  
- Extract and validate required fields  
- Apply metadata mapping and logic-based corrections  
- Distribute quantities across relevant apply months  
- Export outputs to the `uploads/` folder

---

## Output Files

- `CombinedExtractedColumns.xlsx` – contains enriched, validated promotional data  
- `MassUpload.xlsx` – structured file ready for direct system input

---

## Core Modules

| Module                 | Responsibility                                |
|------------------------|-----------------------------------------------|
| `etl.loader`           | Ingests and cleans Excel source files         |
| `etl.parser`           | Date parsing and structural validation        |
| `etl.mapping`          | Metadata enrichment and classification        |
| `etl.validation`       | Field-level validation and consistency checks |
| `etl.grouping`         | Grouping and quantity distribution logic      |
| `writers.excel_writer` | Output writing with formatting/highlighting  |
| `writers.promo_naming` | Generates structured promotion names          |

---

## WBW TV Models

WBW (Weekly Bonus Window) models are detected based on specific columns (e.g., `WBW TV MODEL`).  
Where applicable, these models are used to augment promotion titles.

---

## Notes

- Default values such as `"A SOA"` are applied when type of support is missing  
- Missing mapping fields like `Requestor`, `Customer Type`, or `Currency` are defaulted to `"NA"`  
- All customer codes are normalized using internal standards

---

## Roadmap and BEBS Integration

The following developments are planned or currently under design as part of the BEBS (Business Efficiency & Backend Systems) platform:

- Drag-and-drop upload interface for SFM users  
- Automatic background generation of `MassUpload.xlsx` from input forms  
- Background processing with progress indicators and status tracking  
- Role-based dashboards (e.g., AR, SFM, APM, SFM-MGR)  
- Full audit logging and file traceability  
- Future connection to internal systems for real-time feedback and validation

---

## Implementation Constraints

Some features have not been developed due to the following constraints:

- Use of internal APIs requires formal approval for integration  
- The BEBS platform is under review and may be replaced or restructured  
- As a result, integrated backend workflows (e.g., drag-and-drop auto-processing) are on hold

These limitations are organizational and not related to technical feasibility. Feature development will resume when platform decisions are finalized.

---

## License

MIT License – open for use, modification, and distribution.

---

## Maintainer

Developed by [@doktima](https://github.com/doktima) to support automation, accuracy, and scalability in sales finance workflows.
