import pandas as pd
from pathlib import Path

# Import Excel file as DataFrame
excel_path = Path('data/TLS - Tags for AF rev 1.xlsx')
df = pd.read_excel(excel_path)

# Filter out rows where 'Level 3' is empty or NaN
df = df[df['Level 3'].notna() & (df['Level 3'] != '')]

print(df.head(5))
print(list(df.columns))

grouped = df.groupby('Asset Type Optimised')['P&ID Asset'].nunique()
grouped = grouped.sort_values(ascending=False)
print(grouped)

# Find template attributes for each asset group with more than 2 assets
templates = {}
for asset_type, group in df.groupby('Asset Type Optimised'):
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
for asset_type in grouped[grouped > 2].index:
    asset_count = grouped[asset_type]
    group = df[df['Asset Type Optimised'] == asset_type]
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


