"""
003_AssetsAttributesExtraction.py

Purpose
-------
Read the Excel tag list, ensure columns 'Asset' and 'Attribute' exist, and fill them
based on the Name column using the specified parsing rules.

Parsing rules
------------
1) Take Name, remove the first 4 characters (e.g., 'TNP_').
2) Split the remainder by underscores. When reconstructing, keep underscores between parts.
3) Assignment:
   - Leftmost token starts Asset.
   - Rightmost token is part of Attribute.
   - For any middle tokens:
	   * If token is all digits -> belongs to Asset
	   * Otherwise (has any letters) -> belongs to Attribute
4) Output back to the same Excel file, creating/overwriting 'Asset' and 'Attribute' columns.

Safety
------
Before writing, the script creates a timestamped backup copy next to the input file.

Usage (PowerShell)
------------------
This script uses hardcoded input file and sheet name. Just run:

./venv1/Scripts/python.exe 003_AssetsAttributesExtraction.py
"""

from __future__ import annotations
from datetime import datetime
from pathlib import Path
from typing import List, Tuple

import pandas as pd
import re

# Hardcoded input Excel and sheet name
INPUT_PATH = Path("TNP/TNP - Tags - PI Builder Fromat - 20250624.xlsx")
SHEET_NAME = "PI System - Import Tags - Final"


# No defaults: input file and sheet name are required via CLI


def is_all_digits(token: str) -> bool:
	"""Return True if token contains only digits (ignoring surrounding whitespace)."""
	return bool(token) and token.isdigit()


def parse_name_to_asset_attribute(name: str) -> Tuple[str, str]:
	"""Parse a tag Name into (Asset, Attribute) according to the rules.

	Steps:
	- Remove first 4 characters from name.
	- Split by underscores.
	- Leftmost token -> start of Asset, rightmost -> part of Attribute.
	- Middle tokens: digits -> Asset; otherwise -> Attribute.
	"""
	if not isinstance(name, str):
		name = "" if pd.isna(name) else str(name)

	if len(name) < 4:
		return "", ""

	core = name[4:]  # drop first 4 characters (e.g., TNP_)

	# Split while attaching underscores to the right token to preserve exact separators
	# This keeps single or multiple underscores (e.g., '__') with the token that follows them.
	tokens = re.findall(r'_+[^_]+|[^_]+', core) if core else []
	if not tokens:
		return "", ""

	leftmost = tokens[0]
	rightmost = tokens[-1]
	middle = tokens[1:-1]

	asset_parts: List[str] = [leftmost]
	attr_parts: List[str] = []

	for t in middle:
		t_clean = t.lstrip('_')
		if is_all_digits(t_clean):
			asset_parts.append(t)
		else:
			attr_parts.append(t)

	# Rightmost always belongs to attribute
	attr_parts.append(rightmost)

	# Join without inserting new separators; parts already contain original underscores
	asset = "".join([p for p in asset_parts if p])
	attribute = "".join([p for p in attr_parts if p])
	return asset, attribute


def process_file(input_path: Path, sheet_name: str) -> Path:
	"""Load Excel sheet, ensure Asset/Attribute columns, compute values, and write back.

	Returns the written file path. Requires explicit sheet name.
	"""
	if not input_path.exists():
		raise FileNotFoundError(f"Input Excel not found: {input_path}")

	# Read all sheets into memory and require the provided sheet
	xls = pd.ExcelFile(input_path)
	if sheet_name not in xls.sheet_names:
		raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {xls.sheet_names}")

	# Load all sheets as original snapshot
	sheets_original = {sn: pd.read_excel(input_path, sheet_name=sn) for sn in xls.sheet_names}
	# Work on a deep copy of the target sheet to avoid mutating the snapshot
	df = sheets_original[sheet_name].copy(deep=True)
	if "Name" not in df.columns:
		raise ValueError(f"Sheet '{sheet_name}' does not contain a 'Name' column.")

	# --- Write backup BEFORE any modifications ---
	timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
	backup_path = input_path.with_suffix("")
	backup_path = backup_path.parent / f"{backup_path.name}.backup-{timestamp}.xlsx"
	with pd.ExcelWriter(backup_path, engine="openpyxl") as writer:
		for sn, sdf in sheets_original.items():
			sdf.to_excel(writer, sheet_name=sn, index=False)
	print(f"Backup written: {backup_path}")

	# Ensure columns exist
	if "Asset" not in df.columns:
		df["Asset"] = ""
	if "Attribute" not in df.columns:
		df["Attribute"] = ""

	# Compute for each row
	parsed = df["Name"].apply(parse_name_to_asset_attribute)
	df["Asset"] = parsed.apply(lambda t: t[0])
	df["Attribute"] = parsed.apply(lambda t: t[1])

	# Write updated sheet back to the original file (preserving other sheets)
	out_path = input_path
	# Build updated sheets mapping
	sheets_updated = dict(sheets_original)
	sheets_updated[sheet_name] = df
	with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
		for sn, sdf in sheets_updated.items():
			sdf.to_excel(writer, sheet_name=sn, index=False)

	print(f"Updated file written: {out_path}")
	return out_path


    
def main():
	input_path = INPUT_PATH.resolve()
	process_file(input_path, sheet_name=SHEET_NAME)


if __name__ == "__main__":
	main()

