import pandas as pd
from pathlib import Path

# Hardcoded list of asset types to process in this script
asset_types_to_process = ["Motor", "Motor VSD", "Valve", "Analog Sensor", "PID Controller",
                          "Control Valve", "Flowmeter Totaliser", "Filter"]  # Edit this list as needed

# Import Excel file as DataFrame
excel_path = Path('data/TLS - Tags for AF rev 1.xlsx')
df = pd.read_excel(excel_path)

# Convert all attributes to lower case
df['Attribute'] = df['Attribute'].astype(str).str.lower()

# Filter out rows where 'Level 3' is empty or NaN and only include asset types in asset_types_to_process
df_filtered = df[
    df['Level 3'].notna() &
    (df['Level 3'] != '') &
    (df['Asset Type Optimised'].isin(asset_types_to_process))
]
# Filter out rows where 'Level 3' is empty
df_wohierarchy = df[(df['Level 3'] == '') | (df['Level 3'].isna())]

print(df_filtered.head(5))
print(list(df_filtered.columns))

print(df_wohierarchy.head(5))


grouped = df_filtered.groupby('Asset Type Optimised')['P&ID Asset'].nunique()
grouped = grouped.sort_values(ascending=False)
print(grouped)

# Section 1.
print("\n=== Section 1: Attribute Templates and Coverage ===")
print("This section finds template attributes for each asset type and shows attribute coverage for asset types with more than 2 assets.\n")

# Find template attributes for each asset group with more than 2 assets
templates = {}
for asset_type, group in df_filtered.groupby('Asset Type Optimised'):
    asset_attrs = group.groupby('P&ID Asset')['Attribute'].apply(set)
    if len(asset_attrs) > 2:
        template = set.intersection(*asset_attrs)
        templates[asset_type] = template

# Print templates only for asset types with more than 2 assets
for asset_type, template in templates.items():
    asset_count = grouped.get(asset_type, 0)
    if asset_count > 2:
        print(f"Template for {asset_type}: {sorted(template)}")

# For each asset type with more than 2 assets, show attribute coverage > 80% and asset count, sorted by asset count descending
print("\n--- Attribute Coverage by Asset Type ---")
print("Shows, for each asset type, the percentage of assets that have each attribute (only attributes present in >70% of assets are shown).")
for asset_type in grouped[grouped > 2].index:
    asset_count = grouped[asset_type]
    group = df_filtered[df_filtered['Asset Type Optimised'] == asset_type]
    attr_counts = group.groupby('Attribute')['P&ID Asset'].nunique()
    attr_percent = (attr_counts / asset_count * 100).sort_values(ascending=False)
    filtered = attr_percent[attr_percent > 70]
    selected_attrs = list(filtered.index)
    if selected_attrs:
        print(f"\n{asset_type} (assets: {asset_count}):")
        for attr, percent in filtered.items():
            print(f"  {attr}: {percent:.1f}%")
        # Count assets that have all selected attributes
        assets_with_all = group.groupby('P&ID Asset')['Attribute'].apply(set)
        count_all = (assets_with_all.apply(lambda attrs: set(selected_attrs).issubset(attrs))).sum()
        percent_all = count_all / asset_count * 100
        print(f"\n  Assets with all above attributes: {count_all} ({percent_all:.1f}%)")

        # Also check in df_wohierarchy for assets with all selected attributes
        assets_wohierarchy = df_wohierarchy[df_wohierarchy['Asset Type Optimised'] == asset_type]
        assets_wohierarchy_grouped = assets_wohierarchy.groupby('SCADA Asset')['Attribute'].apply(set)
        count_wohierarchy = (assets_wohierarchy_grouped.apply(lambda attrs: set(selected_attrs).issubset(attrs))).sum()
        print(f"  (In df_wohierarchy) Assets with all above attributes: {count_wohierarchy}")


# Section 1.1. Analytics: which 'SCADA Asset' have different 'P&ID Asset'
print("\n=== Section 1.1: SCADA Asset to P&ID Asset Mapping ===")
print("This section lists SCADA Assets that are linked to more than one unique P&ID Asset.\n")
scada_pid_counts = df_filtered.groupby('SCADA Asset')['P&ID Asset'].nunique()
scada_with_multiple_pid = scada_pid_counts[scada_pid_counts > 1]
if not scada_with_multiple_pid.empty:
    print("\nSection 1.1. SCADA Assets with multiple P&ID Assets:")
    for scada_asset, count in scada_with_multiple_pid.items():
        pids = df_filtered[df_filtered['SCADA Asset'] == scada_asset]['P&ID Asset'].unique()
        print(f"  {scada_asset}: {count} P&ID Assets -> {list(pids)}")
else:
    print("\nSection 1.1. All SCADA Assets map to a single P&ID Asset.")

# Section 2.
print("\n=== Section 2: Asset-Attribute Matrix CSV Export ===")
print("This section exports a CSV matrix for each selected asset type, showing which assets have which attributes. Filtered attributes are shown first, followed by all others.\n")

# List of asset types for CSV export (edit as needed)
asset_types_for_csv = ["Motor", "Motor VSD", "Valve", "Analog Sensor", "PID Controller",
                          "Control Valve", "Flowmeter Totaliser", "Filter"]  # Change as needed

# Ensure output folder exists
output_folder = Path("Attribute Matrix")
output_folder.mkdir(exist_ok=True)

for asset_type_for_csv in asset_types_for_csv:
    group = df_filtered[df_filtered['Asset Type Optimised'] == asset_type_for_csv]
    asset_count = grouped.get(asset_type_for_csv, 0)
    if asset_count and asset_count > 2:
        # Get attribute percentages
        attr_counts = group.groupby('Attribute')['P&ID Asset'].nunique()
        attr_percent = (attr_counts / asset_count * 100).sort_values(ascending=False)
        filtered_attrs = list(attr_percent[attr_percent > 70].index)
        # All attributes: filtered first, then remaining by descending percentage
        all_attrs = list(filtered_attrs) + [attr for attr in attr_percent.index if attr not in filtered_attrs]
        # Get unique asset names from df_filtered
        assets = sorted([str(a) for a in group['P&ID Asset'].unique()])
        # Build DataFrame: rows=assets, columns=all_attrs
        data = []
        for asset in assets:
            asset_attrs = set(group[group['P&ID Asset'] == asset]['Attribute'])
            row = ['yes' if attr in asset_attrs else '' for attr in all_attrs]
            data.append(row)
        csv_df = pd.DataFrame(data, index=assets, columns=all_attrs)

        # --- Add assets from df_wohierarchy ---
        assets_wohierarchy = df_wohierarchy[df_wohierarchy['Asset Type Optimised'] == asset_type_for_csv]
        scada_assets = sorted([str(a) for a in assets_wohierarchy['SCADA Asset'].unique() if pd.notna(a) and a != ''])
        data_wohierarchy = []
        for scada_asset in scada_assets:
            asset_attrs = set(assets_wohierarchy[assets_wohierarchy['SCADA Asset'] == scada_asset]['Attribute'])
            row = ['yes' if attr in asset_attrs else '' for attr in all_attrs]
            data_wohierarchy.append(row)
        if data_wohierarchy:
            df_wohierarchy_matrix = pd.DataFrame(data_wohierarchy, index=scada_assets, columns=all_attrs)
            # Append to the main DataFrame
            csv_df = pd.concat([csv_df, df_wohierarchy_matrix], axis=0)

        # Write to CSV in output folder
        csv_path = output_folder / f"{asset_type_for_csv}_attributes_matrix.csv"
        csv_df.to_csv(csv_path)
        print(f"\nCSV saved to: {csv_path.resolve()}")

# Section 3: Attribute Set Similarity Between Asset Types
print("\n=== Attribute Set Similarity Between Asset Types ===")
print("This section compares the filtered attribute sets (from previous section) between all asset types and outputs the percentage similarity for pairs with more than 70% overlap.\n")

# Collect filtered attribute sets for each asset type
filtered_attr_sets = {}
for asset_type in grouped[grouped > 2].index:
    asset_count = grouped[asset_type]
    group = df_filtered[df_filtered['Asset Type Optimised'] == asset_type]
    attr_counts = group.groupby('Attribute')['P&ID Asset'].nunique()
    attr_percent = (attr_counts / asset_count * 100).sort_values(ascending=False)
    filtered = set(attr_percent[attr_percent > 70].index)
    if filtered:
        filtered_attr_sets[asset_type] = filtered

# Compare sets and output similarity > 70%
asset_types = list(filtered_attr_sets.keys())
for i in range(len(asset_types)):
    for j in range(i + 1, len(asset_types)):
        at1, at2 = asset_types[i], asset_types[j]
        set1, set2 = filtered_attr_sets[at1], filtered_attr_sets[at2]
        if set1 and set2:
            intersection = set1 & set2
            union = set1 | set2
            if union:
                similarity = len(intersection) / len(union) * 100
                if similarity > 70:
                    print(f"{at1} <-> {at2}: {similarity:.1f}% similarity ({len(intersection)} shared of {len(union)} total attributes)")

# Section 4: Total Statistics and Coverage
print("\n=== Section 4: Template Attribute Statistics and File Coverage ===")
print("For each template, shows number of filtered attributes (from coverage section), number of assets with ALL attributes, and their product (total attributes for that template).")
print("Sums all template-asset attributes and compares to total number of rows in the original file.\n")

# Prepare filtered attribute counts per asset type (from coverage section)
filtered_attrs_per_type = {}
assets_with_all_per_type = {}
for asset_type in grouped[grouped > 2].index:
    asset_count = grouped[asset_type]
    group = df_filtered[df_filtered['Asset Type Optimised'] == asset_type]
    attr_counts = group.groupby('Attribute')['P&ID Asset'].nunique()
    attr_percent = (attr_counts / asset_count * 100).sort_values(ascending=False)
    filtered = attr_percent[attr_percent > 70]
    selected_attrs = list(filtered.index)
    filtered_attrs_per_type[asset_type] = selected_attrs
    # Count assets that have all selected attributes
    assets_with_all = group.groupby('P&ID Asset')['Attribute'].apply(set)
    count_all = (assets_with_all.apply(lambda attrs: set(selected_attrs).issubset(attrs))).sum()
    # Also add from df_wohierarchy
    assets_wohierarchy = df_wohierarchy[df_wohierarchy['Asset Type Optimised'] == asset_type]
    assets_wohierarchy_grouped = assets_wohierarchy.groupby('SCADA Asset')['Attribute'].apply(set)
    count_wohierarchy = (assets_wohierarchy_grouped.apply(lambda attrs: set(selected_attrs).issubset(attrs))).sum()
    assets_with_all_per_type[asset_type] = count_all + count_wohierarchy

total_template_asset_attributes = 0
print(f"{'Asset Type':30} {'#Attrs':>8} {'#Assets':>8} {'Total Attrs':>14}")
print("-" * 60)
for asset_type in grouped.index:
    num_attrs = len(filtered_attrs_per_type.get(asset_type, []))
    num_assets = assets_with_all_per_type.get(asset_type, 0)
    total_attrs = num_attrs * num_assets
    total_template_asset_attributes += total_attrs
    print(f"{asset_type:30} {num_attrs:8} {num_assets:8} {total_attrs:14}")

total_rows_in_file = len(df)
print("-" * 60)
print(f"{'SUM':30} {'':8} {'':8} {total_template_asset_attributes:14}")
print(f"\nTotal number of rows in original file: {total_rows_in_file}")
if total_rows_in_file > 0:
    percent_covered = total_template_asset_attributes / total_rows_in_file * 100
    print(f"Template-asset attributes as % of file rows: {percent_covered:.1f}%")
else:
    print("No rows in original file.")



