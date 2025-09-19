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
  - **Template Functionality**: Contains hardcoded switch `TemplatesExtractProvided` (0 = original behavior, 1 = use AF templates)
  - **Template Mode Features**:
    - Imports reference AF templates from `Ref/RefAFTemplates.xlsx`
    - Matches assets to templates based on attribute availability
    - Only includes assets that match at least one template (filters output)
    - Adds "Template" column to output with matched template names
    - Processes templates in descending order by attribute count (most restrictive first)
  - **Template File Format**: Excel file with columns: Name, Parent, ObjectType, BaseTemplate, AttributeConfigString
  - Inputs: `data/TLS - Tags for AF rev 1.xlsx`, `data/Book2.xlsx`, and optionally `Ref/RefAFTemplates.xlsx`
  - Output: `data/TLS_AF_Import.xlsx` with hierarchical structure and optional Template column

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
├── Ref/                           # Reference template files
│   └── RefAFTemplates.xlsx              # AF template definitions for 010 script
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

### Standard Template Analysis Workflow
1. **Prepare data**: Place tag list in `data/TLS - Tags for AF rev 1.xlsx`
2. **Extract templates**: Run `020_TemplateExtraction.py` to analyze and create template specifications
3. **Generate AF files**: Run `025_AFtemplateGeneration.py` to create PI Builder import file

### Template-Based Asset Filtering Workflow
1. **Prepare data**: Place tag list in `data/TLS - Tags for AF rev 1.xlsx`
2. **Create reference templates**: Create `Ref/RefAFTemplates.xlsx` with desired ElementTemplate and AttributeTemplate definitions
3. **Set template mode**: In `010_TreeTagList.py`, set `TemplatesExtractProvided = 1`
4. **Generate filtered hierarchy**: Run `010_TreeTagList.py` to create asset hierarchy with only template-matching assets

### Optional Scripts
- **Asset/Attribute parsing**: Run `030_AssetsAttributesExtraction.py` for asset/attribute parsing
- **Standard hierarchy**: Run `010_TreeTagList.py` with `TemplatesExtractProvided = 0` for complete asset hierarchy

## Output Files

- `AF_Templates_Specification.json`: Detailed template specifications with metadata
- `data/AF_Templates_PIBuilder_*.csv`: Ready-to-import file for PI Builder Excel Add-on
- `Attribute Matrix/*.csv`: Per-asset-type attribute presence matrices
- `data/TLS_AF_Import.xlsx`: Hierarchical element structure for AF import (with optional Template column)

## Template Matching Details

When using template mode in `010_TreeTagList.py`:

### Reference Template File Format (`Ref/RefAFTemplates.xlsx`)
- **Name**: Element template name or attribute name
- **Parent**: Empty for ElementTemplate, contains template name for AttributeTemplate
- **ObjectType**: "ElementTemplate" or "AttributeTemplate"
- **BaseTemplate**: Inheritance base (e.g., "TLS" for base assets)
- **AttributeConfigString**: Contains mapping like `\\%@\zzz.GlobalConfiguration|Data Archive Name%\TLS_%Element%{attribute}`

### Template Matching Algorithm
1. Extracts tag attributes from "Attribute Optimised" column in TLS data
2. Groups attributes by P&ID Asset
3. Processes templates in descending order by attribute count (most restrictive first)
4. Matches assets that have ALL required attributes for a template
5. Filters output to include only successfully matched assets
6. Adds Template column with matched template name

### Example Template Matching Results
- Original mode: All 767 assets included, Template column empty
- Template mode: 87 assets matched to templates (out of 682 processed), filtered output

## Notes

- All scripts use hardcoded input paths for consistency
- Backup files are created before modifying existing Excel files
- Template naming follows TLS convention with version numbers
- Template matching uses exact attribute name matching (case-insensitive)
- In template mode, only assets with complete attribute sets are included
- Substitution patterns are standardized for PI Point integration