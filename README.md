# SEQ_TLS_AF_work

This repository contains scripts for analyzing TLS tag data and generating AVEVA Asset Framework (AF) templates and import files.

## Project Overview

The project processes Excel tag list data to:
1. Build hierarchical AF element trees for asset structure
2. Extract and analyze asset attributes for template creation
3. Generate AF template specifications and import-ready files

## Scripts Overview

### Core Analysis Scripts

- **010_TreeTagList.py** (formerly 001_TreeTagList.py)
  - Purpose: Builds a hierarchical Asset Framework (AF) element tree from the TLS tag workbook and a small config workbook.
  - Inputs: `data/TLS - Tags for AF rev 1.xlsx` and `data/Book2.xlsx`.
  - Output: `data/TLS_AF_Import.xlsx` with rows ready for AF import (columns like Parent, Name, ObjectType, Description, SecurityString). The level structure is: Level 1 root (site), Level 2, Level 3, then leaf elements (P&ID Asset). Description on leaf elements includes the Asset Name for readability.

- **020_TemplateExtraction.py** (formerly 002_Template01.py)
  - Purpose: Analyzes the tag list to derive attribute templates for common asset types and exports helpful artifacts.
  - Key steps:
    - Computes attribute coverage per asset type and identifies high-coverage attributes (>70%).
    - Exports per-asset-type attribute presence matrices into the `Attribute Matrix/` folder.
    - Compares attribute sets between asset types for similarity.
    - Summarizes template stats and coverage.
    - Generates `AF_Templates_Specification.json` describing element templates: template name/description, attributes (with data type, engineering units), and a standard substitution pattern `"<%AssetName%><%@Attribute>"` for PI Point auto-assignment in AF.
  - Inputs: `data/TLS - Tags for AF rev 1.xlsx`.
  - Outputs: CSV matrices under `Attribute Matrix/` and `AF_Templates_Specification.json`.

### Template Generation Scripts

- **025_AFtemplateGeneration.py**
  - Purpose: Imports the JSON template specification and creates a CSV file formatted for AVEVA Asset Framework import via PI Builder Excel Add-on.
  - Key features:
    - Reads `AF_Templates_Specification.json` created by 020_TemplateExtraction.py
    - Generates CSV file matching PI Builder import format structure
    - Creates ElementTemplate and AttributeTemplate rows with proper security strings and data references
    - Template naming convention: `TLS.{TemplateName}.001`
    - Data reference format: `\\%@\zzz.GlobalConfiguration|Data Archive Name%\TLS_%Element%{AttributeSuffix}`
  - Input: `AF_Templates_Specification.json`
  - Output: `data/AF_Templates_PIBuilder_YYYYMMDD_HHMMSS.csv`

### Data Processing Scripts

- **030_AssetsAttributesExtraction.py** (formerly 003_AssetsAttributesExtraction.py)
  - Purpose: Processes tag names to extract Asset and Attribute information using tokenization rules.
  - Key features:
    - Parses tag names by removing first 4 characters and tokenizing by underscores
    - Classifies middle tokens as Asset (if digits-only) or Attribute
    - Preserves exact underscore placement during parsing
    - Creates backup before modifying Excel file
    - Hardcoded input: `data/TLS - Tags for AF rev 1.xlsx`, sheet "PI System - Import Tags - Final"
  - Output: Adds Asset and Attribute columns to the Excel file

## File Structure

```
/
├── data/                          # Input and output data files
│   ├── TLS - Tags for AF rev 1.xlsx     # Main tag list input
│   ├── Book2.xlsx                       # Configuration input
│   ├── Template_Import.xlsx             # Template import file
│   └── AF_Templates_PIBuilder_*.csv     # Generated PI Builder files
├── Attribute Matrix/              # Generated attribute matrices
│   └── {AssetType}_attributes_matrix.csv
├── AF_Templates_Specification.json      # Template specifications
├── requirements.txt               # Python dependencies
└── README.md                     # This file
```

## Dependencies

The scripts require the following Python packages (see `requirements.txt`):
- pandas
- openpyxl
- pathlib
- json

## Usage Workflow

1. **Prepare data**: Place tag list in `data/TLS - Tags for AF rev 1.xlsx`
2. **Extract templates**: Run `020_TemplateExtraction.py` to analyze and create template specifications
3. **Generate AF files**: Run `025_AFtemplateGeneration.py` to create PI Builder import file
4. **Optional**: Run `010_TreeTagList.py` for hierarchical element structure
5. **Optional**: Run `030_AssetsAttributesExtraction.py` for asset/attribute parsing

## Output Files

- `AF_Templates_Specification.json`: Detailed template specifications with metadata
- `data/AF_Templates_PIBuilder_*.csv`: Ready-to-import file for PI Builder Excel Add-on
- `Attribute Matrix/*.csv`: Per-asset-type attribute presence matrices
- `data/TLS_AF_Import.xlsx`: Hierarchical element structure for AF import

## Notes

- All scripts use hardcoded input paths for consistency
- Backup files are created before modifying existing Excel files
- Template naming follows TLS convention with version numbers
- Substitution patterns are standardized for PI Point integration