# SEQ_TLS_AF_work

## Scripts overview

- 001_TreeTagList.py
  - Purpose: Builds a hierarchical Asset Framework (AF) element tree from the TLS tag workbook and a small config workbook.
  - Inputs: `data/TLS - Tags for AF rev 1.xlsx` and `data/Book2.xlsx`.
  - Output: `data/TLS_AF_Import.xlsx` with rows ready for AF import (columns like Parent, Name, ObjectType, Description, SecurityString). The level structure is: Level 1 root (site), Level 2, Level 3, then leaf elements (P&ID Asset). Description on leaf elements includes the Asset Name for readability.

- 002_Template01.py
  - Purpose: Analyzes the tag list to derive attribute templates for common asset types and exports helpful artifacts.
  - Key steps:
    - Computes attribute coverage per asset type and identifies high-coverage attributes (>70%).
    - Exports per-asset-type attribute presence matrices into the `Attribute Matrix/` folder.
    - Compares attribute sets between asset types for similarity.
    - Summarizes template stats and coverage.
    - Generates `AF_Templates_Specification.json` describing element templates: template name/description, attributes (with data type, engineering units), and a standard substitution pattern `"<%AssetName%><%@Attribute>"` for PI Point auto-assignment in AF.
  - Inputs: `data/TLS - Tags for AF rev 1.xlsx`.
  - Outputs: CSV matrices under `Attribute Matrix/` and `AF_Templates_Specification.json`.