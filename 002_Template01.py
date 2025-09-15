import pandas as pd
from pathlib import Path

def build_template_import(
    tls_path: str,
    book2_path: str,
    output_path: str,
    asset_type_filter: list = None,  # Placeholder - fill this list before running
    template_name: str = "YourTemplateName"  # Placeholder - fill this before running
) -> None:
    """
    Create template import file by filtering assets by type and finding common attributes.
    
    Args:
        tls_path: Path to TLS Excel file
        book2_path: Path to Book2 Excel file  
        output_path: Path for output Excel file
        asset_type_filter: List of asset types to filter by (e.g., ['Pump', 'Tank'])
        template_name: Name of the template to use in the Template column
    """
    
    # --- Load inputs (same as 001_TreeTagList) ---
    try:
        tls_sheets = pd.read_excel(tls_path, sheet_name=None)
    except Exception as e:
        print(f"ERROR loading TLS file: {e}")
        return
    
    if "PI System - Import Tags - Final" in tls_sheets:
        tls_df = tls_sheets["PI System - Import Tags - Final"]
    else:
        tls_df = None
        for name, df in tls_sheets.items():
            if {"P&ID Asset", "Asset Name", "Level 2", "Level 3"} <= set(df.columns):
                tls_df = df
                break
        if tls_df is None:
            print("ERROR: Could not find a sheet with required columns.")
            return

    try:
        book_df = pd.read_excel(book2_path)
    except Exception as e:
        print(f"ERROR loading Book2 file: {e}")
        return

    # Use SecurityString from Book2 if available; otherwise blank
    sec_str = ""
    if "SecurityString" in book_df.columns and not book_df["SecurityString"].dropna().empty:
        sec_str = str(book_df["SecurityString"].dropna().iloc[0])

    # --- Sanity / cleanup ---
    needed_cols = ["P&ID Asset", "Asset Name", "Level 2", "Level 3"]
    
    # Check if Asset Type column exists
    if "Asset Type" not in tls_df.columns:
        print("ERROR: 'Asset Type' column not found!")
        return
    
    missing = [c for c in needed_cols if c not in tls_df.columns]
    if missing:
        print(f"ERROR: Missing required columns in TLS sheet: {missing}")
        return

    # Keep only rows with a P&ID Asset
    df = tls_df.copy()
    df = df[df["P&ID Asset"].notna()].copy()

    # Normalise types/whitespace
    for col in ["P&ID Asset", "Asset Name", "Level 2", "Level 3"]:
        df[col] = df[col].astype(str).str.strip()
    
    if "Asset Type" in df.columns:
        df["Asset Type"] = df["Asset Type"].astype(str).str.strip()

    # --- 1. Filter by Asset Type ---
    if asset_type_filter is None:
        asset_type_filter = ["PLACEHOLDER_ASSET_TYPE_1", "PLACEHOLDER_ASSET_TYPE_2"]
        print(f"WARNING: Using placeholder asset types: {asset_type_filter}")
        return
    
    if "Asset Type" in df.columns:
        filtered_df = df[df["Asset Type"].isin(asset_type_filter)].copy()
        if len(filtered_df) == 0:
            print("WARNING: No assets found matching the filter criteria!")
            return
    else:
        print("Asset Type column not found, using all assets")
        filtered_df = df.copy()

    # Show filtered assets in one compact line
    unique_assets = sorted(filtered_df["P&ID Asset"].unique())
    print(f"Assets included ({len(unique_assets)}): {', '.join(unique_assets)}")

    # --- 2. Find common attributes for the filtered assets ---
    # Check if 'Attribute' column exists
    if "Attribute" not in filtered_df.columns:
        print("ERROR: 'Attribute' column not found!")
        return
    
    # Group by asset to see what attributes each asset has
    asset_attributes = {}
    for _, row in filtered_df.iterrows():
        asset = row["P&ID Asset"]
        attribute = row["Attribute"]
        if pd.notna(attribute):
            if asset not in asset_attributes:
                asset_attributes[asset] = set()
            asset_attributes[asset].add(str(attribute).strip())
    
    # Show what each asset has for debugging
    total_assets = len(asset_attributes)
    
    # Calculate attribute coverage percentages
    from collections import Counter
    all_attributes = []
    for attrs in asset_attributes.values():
        all_attributes.extend(attrs)
    
    attr_counts = Counter(all_attributes)
    
    print(f"\nAttribute coverage (>50%):")
    # Sort by percentage descending, only show >50% coverage
    for attr, count in attr_counts.most_common():
        percentage = (count / total_assets) * 100
        if percentage > 50:
            print(f"  {attr}: {percentage:.1f}% ({count}/{total_assets} assets)")
    
    # Find attributes that appear in ALL assets (truly common)
    if total_assets == 0:
        print("No assets found!")
        return
    
    # Start with attributes from first asset, then find intersection with others
    common_attributes = None
    for asset, attrs in asset_attributes.items():
        if common_attributes is None:
            common_attributes = attrs.copy()
        else:
            common_attributes = common_attributes.intersection(attrs)
    
    if common_attributes is None:
        common_attributes = set()
    
    print(f"\nCommon attributes ({len(common_attributes)}): {', '.join(sorted(common_attributes))}")

    # --- Build output rows ---
    # Start with the original filtered data
    output_df = filtered_df.copy()
    
    # Add Template column
    output_df["Template"] = ""
    
    # Fill Template column for rows that have common attributes
    template_rows_count = 0
    for idx, row in output_df.iterrows():
        attribute = row["Attribute"]
        if pd.notna(attribute) and str(attribute).strip() in common_attributes:
            output_df.at[idx, "Template"] = template_name
            template_rows_count += 1

    # Prepare final output with required columns
    final_columns = ["Selected(x)", "Parent", "Name", "ObjectType", "Error", 
                    "Description", "Template", "SecurityString"]
    
    # Add missing columns with default values
    if "Selected(x)" not in output_df.columns:
        output_df["Selected(x)"] = "x"
    if "Parent" not in output_df.columns:
        output_df["Parent"] = ""
    if "Name" not in output_df.columns:
        output_df["Name"] = output_df["P&ID Asset"]
    if "ObjectType" not in output_df.columns:
        output_df["ObjectType"] = "Attribute"
    if "Error" not in output_df.columns:
        output_df["Error"] = ""
    if "Description" not in output_df.columns:
        output_df["Description"] = output_df["Asset Name"]
    if "SecurityString" not in output_df.columns:
        output_df["SecurityString"] = sec_str
    
    # Select only the required columns for output
    out_df = output_df[final_columns].copy()
    
    print(f"Output DataFrame shape: {out_df.shape}")
    print(f"Output DataFrame preview:")
    print(out_df.head())

    # Save to Excel
    print(f"Saving to Excel: {output_path}")
    try:
        out_df.to_excel(output_path, index=False)
        template_count = len(out_df[out_df['Template'] != ''])
        print(f"SUCCESS: Created {output_path} with {len(out_df)} rows ({template_count} with templates)")
    except Exception as e:
        print(f"ERROR saving to Excel: {e}")
        return

if __name__ == "__main__":
    # Fill these before running:
    asset_type_filter = ["Flowmeter","Transmitter","Level Transmitter","Pressure Transmitter",
                         "Temperature Transmitter", "Vibration Transmitter","Weight Transmitter",
                         "Analyser - DO"]  # e.g., ["Pump", "Tank", "Valve"]
    template_name = "Analog Value"  # e.g., "StandardPumpTemplate"
    
    tls_path = "data/TLS - Tags for AF rev 1.xlsx"
    book2_path = "data/Book2.xlsx" 
    output_path = "data/Template_Import.xlsx"
    
    build_template_import(tls_path, book2_path, output_path, asset_type_filter, template_name)
    print(f"Written: {output_path}")
