import pandas as pd
from pathlib import Path
import json
import re

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

# Section 5: Generate JSON Template Specifications
print("\n=== Section 5: JSON Template Specifications ===")
print("This section creates a JSON file with template specifications including name, description, attributes with descriptions, data types, and substitution patterns.\n")

def extract_common_description_patterns(descriptions):
    """Extract common rightmost patterns from descriptions for template description"""
    if descriptions is None or len(descriptions) == 0:
        return "Asset template"
    
    # Clean and collect all valid descriptions
    valid_descriptions = []
    for desc in descriptions:
        if pd.notna(desc) and str(desc).strip():
            valid_descriptions.append(str(desc).strip())
    
    if len(valid_descriptions) == 0:
        return "Asset template"
    
    threshold = max(1, int(len(valid_descriptions) * 0.8))  # 80% threshold, minimum 1
    
    # First, remove asset identifiers (like NV2611, NV41107, etc.) from descriptions
    cleaned_descriptions = []
    for desc in valid_descriptions:
        # Remove asset identifiers at the beginning (e.g., NV2611, NV41107, etc.)
        cleaned_desc = re.sub(r'^[A-Z]+\d+[A-Z]*\s+', '', desc)
        cleaned_descriptions.append(cleaned_desc)
    
    # Try different rightmost substring lengths (starting from longer patterns)
    best_common = "Asset template"
    best_count = 0
    
    # Check rightmost substrings of different lengths (start from longer ones)
    for length in range(min(50, max(len(d) for d in cleaned_descriptions)), 2, -1):  
        rightmost_patterns = {}
        
        for desc in cleaned_descriptions:
            if len(desc) >= length:
                # Get rightmost substring
                rightmost = desc[-length:].strip()
                # Clean up spacing
                rightmost = re.sub(r'\s+', ' ', rightmost)
                
                if rightmost and len(rightmost) > 3:  # Only consider meaningful patterns
                    rightmost_patterns[rightmost] = rightmost_patterns.get(rightmost, 0) + 1
        
        # Find pattern that appears most frequently and meets threshold
        for pattern, count in sorted(rightmost_patterns.items(), key=lambda x: x[1], reverse=True):
            if count >= threshold and count > best_count:
                best_common = pattern
                best_count = count
                break  # Take the most common pattern that meets threshold
    
    # If no rightmost pattern found, try finding common suffix words
    if best_count < threshold:
        # Split descriptions into words and find common endings
        word_pattern_counts = {}
        for desc in cleaned_descriptions:
            words = desc.split()
            # Try different combinations of ending words
            for i in range(1, min(5, len(words) + 1)):  # Check 1-4 ending words
                ending = ' '.join(words[-i:])
                if len(ending) > 3:  # Only meaningful endings
                    word_pattern_counts[ending] = word_pattern_counts.get(ending, 0) + 1
        
        # Find most common ending that meets threshold
        for pattern, count in sorted(word_pattern_counts.items(), key=lambda x: x[1], reverse=True):
            if count >= threshold and count > best_count:
                best_common = pattern
                best_count = count
                break
    
    # Final fallback
    if best_common == "Asset template" or best_count == 0:
        return "Asset control and monitoring"
    
    return best_common

def map_to_aveva_datatype(pointtype, engunits=None):
    """Map Excel pointtype to AVEVA Asset Framework acceptable data types"""
    if pd.isna(pointtype):
        return "Float64"
    
    pointtype_str = str(pointtype).lower()
    
    # AVEVA AF Data Types mapping
    if pointtype_str in ['digital', 'bool', 'boolean']:
        return "Boolean"
    elif pointtype_str in ['int16', 'integer', 'int']:
        return "Int32"
    elif pointtype_str in ['int32']:
        return "Int32"
    elif pointtype_str in ['float', 'real', 'double', 'single']:
        return "Float64"
    elif pointtype_str in ['string', 'text']:
        return "String"
    elif pointtype_str in ['datetime', 'timestamp']:
        return "DateTime"
    else:
        # Default based on engineering units
        if pd.notna(engunits):
            return "Float64"  # Most engineering values are float
        return "Float64"  # Default

def create_substitution_pattern(asset_type, attribute):
    """Create substitution pattern for PI Point assignment"""
    # Use the literal pattern format for all attributes
    return "<%AssetName%><%@Attribute>"

# Create templates dictionary
templates_json = {
    "metadata": {
        "created_date": "2025-09-16",
        "description": "AVEVA Asset Framework templates generated from tag analysis",
        "source_file": "TLS - Tags for AF rev 1.xlsx",
        "version": "1.0"
    },
    "templates": []
}

# Process each asset type with filtered attributes
for asset_type in grouped[grouped > 2].index:
    asset_count = grouped[asset_type]
    group = df_filtered[df_filtered['Asset Type Optimised'] == asset_type]
    attr_counts = group.groupby('Attribute')['P&ID Asset'].nunique()
    attr_percent = (attr_counts / asset_count * 100).sort_values(ascending=False)
    filtered_attrs = attr_percent[attr_percent > 70]
    
    if len(filtered_attrs) > 0:
        print(f"Processing template for {asset_type}...")
        
        # Get template description by analyzing asset descriptions
        asset_descriptions = group['Description'].unique()
        template_description = extract_common_description_patterns(asset_descriptions)
        
        # Create attributes list
        attributes_list = []
        for attr in filtered_attrs.index:
            attr_group = group[group['Attribute'] == attr]
            
            # Get most common description for this attribute
            attr_descriptions = attr_group['Description'].dropna()
            if len(attr_descriptions) > 0:
                attr_desc = attr_descriptions.iloc[0]  # Take first non-null description
            else:
                attr_desc = f"{attr} attribute for {asset_type}"
            
            # Get most common point type and eng units
            pointtypes = attr_group['poInttype'].dropna()
            engunits = attr_group['engunits'].dropna()
            
            pointtype = pointtypes.iloc[0] if len(pointtypes) > 0 else None
            engunit = engunits.iloc[0] if len(engunits) > 0 else None
            
            # Map to AVEVA data type
            aveva_datatype = map_to_aveva_datatype(pointtype, engunit)
            
            # Create substitution pattern
            substitution = create_substitution_pattern(asset_type, attr)
            
            attribute_spec = {
                "name": attr,
                "description": str(attr_desc) if pd.notna(attr_desc) else f"{attr} attribute",
                "data_type": aveva_datatype,
                "engineering_units": str(engunit) if pd.notna(engunit) else "",
                "point_type": str(pointtype) if pd.notna(pointtype) else "",
                "substitution_pattern": substitution,
                "coverage_percentage": round(filtered_attrs[attr], 1),
                "pi_point_config": {
                    "point_source": "L",
                    "point_class": "classic",
                    "auto_create": True
                }
            }
            
            attributes_list.append(attribute_spec)
        
        # Create template specification
        template_spec = {
            "name": asset_type,
            "description": template_description,
            "category": "Equipment",
            "asset_count_with_template": int(assets_with_all_per_type.get(asset_type, 0)),
            "total_asset_count": int(asset_count),
            "coverage_percentage": round((int(assets_with_all_per_type.get(asset_type, 0)) / int(asset_count) * 100), 1),
            "attributes": attributes_list,
            "element_template_config": {
                "allow_element_to_extend": True,
                "security": "AF_SECURITY",
                "categories": ["Equipment", asset_type]
            }
        }
        
        templates_json["templates"].append(template_spec)

# Save JSON file
json_output_path = Path("AF_Templates_Specification.json")
with open(json_output_path, 'w', encoding='utf-8') as f:
    json.dump(templates_json, f, indent=2, ensure_ascii=False)

print(f"\nJSON template specification saved to: {json_output_path.resolve()}")
print(f"Total templates created: {len(templates_json['templates'])}")

# Print summary
print("\n--- Template Summary ---")
for template in templates_json["templates"]:
    print(f"{template['name']}: {len(template['attributes'])} attributes, {template['asset_count_with_template']} assets ({template['coverage_percentage']}% coverage)")

print(f"\nDetailed specifications are available in: {json_output_path}")



