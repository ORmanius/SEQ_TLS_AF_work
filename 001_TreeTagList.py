import pandas as pd
from pathlib import Path

def build_af_import(
    tls_path: str,
    book2_path: str,
    output_path: str,
    level1_name: str = "TLS - Landers Shute WTP"
) -> None:
    # --- Load inputs ---
    tls_sheets = pd.read_excel(tls_path, sheet_name=None)
    if "PI System - Import Tags - Final" in tls_sheets:
        tls_df = tls_sheets["PI System - Import Tags - Final"]
    else:
        tls_df = None
        for name, df in pd.read_excel(tls_path, sheet_name=None).items():
            if {"P&ID Asset", "Asset Name", "Level 2", "Level 3"} <= set(df.columns):
                tls_df = df
                break
        if tls_df is None:
            raise ValueError("Could not find a sheet with required columns.")

    book_df = pd.read_excel(book2_path)

    # Use SecurityString from Book2 if available; otherwise blank
    sec_str = ""
    if "SecurityString" in book_df.columns and not book_df["SecurityString"].dropna().empty:
        sec_str = str(book_df["SecurityString"].dropna().iloc[0])

    # --- Sanity / cleanup ---
    needed_cols = ["P&ID Asset", "Asset Name", "Level 2", "Level 3"]
    missing = [c for c in needed_cols if c not in tls_df.columns]
    if missing:
        raise ValueError(f"Missing required columns in TLS sheet: {missing}")

    # Keep only rows with a P&ID Asset (that's the AF Element name)
    df = tls_df.copy()
    df = df[df["P&ID Asset"].notna()].copy()

    # Normalise types/whitespace
    for col in ["P&ID Asset", "Asset Name", "Level 2", "Level 3"]:
        df[col] = df[col].astype(str).str.strip()

    # Deduplicate by Element name (P&ID Asset) within its hierarchy
    df.sort_values(by=["Level 2", "Level 3", "P&ID Asset"], inplace=True)
    dedup_cols = ["Level 2", "Level 3", "P&ID Asset"]
    df = df.groupby(dedup_cols, as_index=False).agg({
        "Asset Name": "first"
    })

    rows = []

    def add_row(parent, name, description=""):
        # Add description to name with a ' - ' separator if description exists
        display_name = name
        if description:
            display_name = f"{name} - {description}"
            
        rows.append({
            "Selected(x)": "x",
            "Parent": parent,
            "Name": display_name,
            "ObjectType": "Element",
            "Error": "",
            "Description": description,
            "SecurityString": sec_str
        })

    # Level 1 (root)
    add_row(parent=pd.NA, name=level1_name, description="")

    # Level 2 elements
    for l2 in df["Level 2"].dropna().unique():
        if not str(l2).strip() or str(l2).lower() == "nan":
            continue
        add_row(parent=level1_name, name=str(l2), description="")

    # Level 3 under each Level 2
    for _, grp in df.groupby("Level 2", dropna=True):
        l2_val = str(grp["Level 2"].iloc[0])
        for l3_val in grp["Level 3"].dropna().unique():
            if not str(l3_val).strip() or str(l3_val).lower() == "nan":
                continue
            add_row(parent=f"{level1_name}\\{l2_val}", name=str(l3_val), description="")

    # Leaf elements: P&ID Asset under Level 3
    for _, row in df.iterrows():
        l2 = str(row["Level 2"])
        l3 = str(row["Level 3"])
        pid_asset = str(row["P&ID Asset"])
        asset_name = str(row["Asset Name"]) if pd.notna(row["Asset Name"]) else ""
        parent_path = f"{level1_name}\\{l2}" if l2 and l2.lower() != "nan" else level1_name
        if l3 and l3.lower() != "nan" and str(l3).strip():
            parent_path = f"{parent_path}\\{l3}"
        add_row(parent=parent_path, name=pid_asset, description=asset_name or "")

    out_df = pd.DataFrame(rows, columns=[
        "Selected(x)","Parent","Name","ObjectType","Error","Description","SecurityString"
    ])

    # Save to Excel
    out_df.to_excel(output_path, index=False)

if __name__ == "__main__":
    tls_path = "data/TLS - Tags for AF.xlsx"
    book2_path = "data/Book2.xlsx"
    output_path = "data/TLS_AF_Import.xlsx"
    build_af_import(tls_path, book2_path, output_path)
    print(f"Written: {output_path}")
