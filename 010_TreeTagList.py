import pandas as pd
from pathlib import Path

# Hardcoded switch for template extraction
TemplatesExtractProvided = 1  # 0 = original behavior, 1 = use AF templates

def load_af_templates(template_file_path: str):
    """Load AF templates from reference Excel file"""
    if not Path(template_file_path).exists():
        print(f"Warning: Template file {template_file_path} not found. Using original behavior.")
        return None, None
    
    template_df = pd.read_excel(template_file_path)
    
    # Validate required columns
    required_cols = ["Name", "Parent", "ObjectType"]
    missing_cols = [col for col in required_cols if col not in template_df.columns]
    if missing_cols:
        print(f"Warning: Missing columns in template file: {missing_cols}. Using original behavior.")
        return None, None
    
    # Separate element templates and attribute templates
    element_templates = template_df[template_df["ObjectType"] == "ElementTemplate"].copy()
    attribute_templates = template_df[template_df["ObjectType"] == "AttributeTemplate"].copy()
    
    # Process templates to create a structured dictionary
    templates = {}
    
    # First pass: collect all direct attributes for each template
    for _, elem_row in element_templates.iterrows():
        template_name = elem_row["Name"]
        base_template = elem_row.get("BaseTemplate", "")
        
        # Get attributes for this template
        template_attrs = attribute_templates[attribute_templates["Parent"] == template_name]
        attributes = []
        
        for _, attr_row in template_attrs.iterrows():
            attr_config = attr_row.get("AttributeConfigString", "")
            # Extract TAGATTRIBUTE from the config string
            tag_attribute = ""
            
            # Check if attr_config is not NaN and is a string
            if pd.notna(attr_config) and isinstance(attr_config, str):
                if "%@|Site Code%_%@|SCADA Asset Name%" in attr_config:
                    # Extract the part after %@|Site Code%_%@|SCADA Asset Name%
                    parts = attr_config.split("%@|Site Code%_%@|SCADA Asset Name%")
                    if len(parts) > 1:
                        tag_attribute = parts[1].lower().strip()  # Convert to lowercase and strip whitespace
            
            attributes.append({
                "name": attr_row["Name"],
                "tag_attribute": tag_attribute,
                "config_string": attr_config
            })
        
        templates[template_name] = {
            "base_template": base_template,
            "direct_attributes": attributes,
            "all_attributes": [],  # Will be populated with inheritance
            "attribute_count": 0   # Will be updated after inheritance
        }
    
    # Second pass: resolve inheritance and populate all_attributes
    def get_all_attributes(template_name, visited=None):
        """Recursively get all attributes including inherited ones"""
        if visited is None:
            visited = set()
        
        if template_name in visited:
            return []  # Avoid circular dependencies
        
        visited.add(template_name)
        
        if template_name not in templates:
            return []
        
        template = templates[template_name]
        all_attrs = template["direct_attributes"].copy()
        
        # Add inherited attributes from base template
        base_template = template["base_template"]
        if base_template and base_template in templates:
            inherited_attrs = get_all_attributes(base_template, visited.copy())
            
            # Merge attributes (direct attributes override inherited ones with same name)
            existing_names = {attr["name"] for attr in all_attrs}
            for inherited_attr in inherited_attrs:
                if inherited_attr["name"] not in existing_names:
                    all_attrs.append(inherited_attr)
        
        return all_attrs
    
    # Populate all_attributes for each template
    for template_name in templates:
        all_attrs = get_all_attributes(template_name)
        templates[template_name]["all_attributes"] = all_attrs
        templates[template_name]["attribute_count"] = len(all_attrs)
    
    # Sort templates by attribute count (descending)
    sorted_templates = dict(sorted(templates.items(), key=lambda x: x[1]["attribute_count"], reverse=True))
    
    return sorted_templates, attribute_templates

def match_assets_to_templates(df, templates):
    """Match assets to templates based on attribute availability"""
    if not templates:
        return df
    
    # Load the main TLS data to get attribute information
    tls_path = "data/TLS - Tags for AF rev 1.xlsx"
    try:
        tls_full_df = pd.read_excel(tls_path, sheet_name="PI System - Import Tags - Final")
    except:
        print("Warning: Could not load full TLS data for template matching")
        return df
    
    # Ensure we have the required column
    if "Attribute Optimised" not in tls_full_df.columns:
        print("Warning: 'Attribute Optimised' column not found in TLS data")
        return df
    
    # Create a mapping of P&ID Asset to their attributes
    asset_attributes = {}
    for _, row in tls_full_df.iterrows():
        if pd.notna(row.get("P&ID Asset")) and pd.notna(row.get("Attribute Optimised")):
            asset_name = str(row["P&ID Asset"]).strip()
            attribute = str(row["Attribute Optimised"]).lower().strip()
            
            if asset_name not in asset_attributes:
                asset_attributes[asset_name] = set()
            asset_attributes[asset_name].add(attribute)
    
    # Match assets to templates
    asset_template_mapping = {}
    
    # Try templates in order of most attributes first
    for template_name, template_info in templates.items():
        required_attributes = set(attr["tag_attribute"] for attr in template_info["all_attributes"] if attr["tag_attribute"])
        
        # Skip templates with no valid attributes
        if not required_attributes:
            continue
        
        matched_count = 0
        # Find assets that have all required attributes and are not already matched
        for asset_name, asset_attrs in asset_attributes.items():
            if asset_name not in asset_template_mapping:  # Not already matched
                if required_attributes.issubset(asset_attrs):
                    asset_template_mapping[asset_name] = template_name
                    matched_count += 1
        
        print(f"Template '{template_name}': matched {matched_count} assets with {len(required_attributes)} required attributes")
    
    # Add Template column to the dataframe
    df["Template"] = df["P&ID Asset"].map(asset_template_mapping).fillna("")
    
    return df

def build_af_import(
    tls_path: str,
    book2_path: str,
    output_path: str,
    level1_name: str = "TLS - Landers Shute WTP rev2"
) -> None:
    # --- Load AF Templates if enabled ---
    templates = None
    attribute_templates = None
    
    if TemplatesExtractProvided == 1:
        template_file_path = "Ref/RefAFTemplates.xlsx"
        print(f"Loading AF templates from: {template_file_path}")
        templates, attribute_templates = load_af_templates(template_file_path)
        if templates:
            print(f"Loaded {len(templates)} templates:")
            for name, info in templates.items():
                print(f"  - {name}: {info['attribute_count']} attributes")
        else:
            print("No templates loaded, using original behavior")
    
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

    # Create SCADA Asset mapping before deduplication
    scada_asset_mapping = {}
    for _, row in df.iterrows():
        pid_asset = str(row["P&ID Asset"]).strip()
        if pd.notna(row.get("SCADA Asset")):
            scada_asset_value = str(row["SCADA Asset"]).strip()
            scada_asset_mapping[pid_asset] = scada_asset_value

    # Normalise types/whitespace
    for col in ["P&ID Asset", "Asset Name", "Level 2", "Level 3"]:
        df[col] = df[col].astype(str).str.strip()

    # Deduplicate by Element name (P&ID Asset) within its hierarchy
    df.sort_values(by=["Level 2", "Level 3", "P&ID Asset"], inplace=True)
    dedup_cols = ["Level 2", "Level 3", "P&ID Asset"]
    df = df.groupby(dedup_cols, as_index=False).agg({
        "Asset Name": "first"
    })

    # --- Apply template matching if enabled ---
    if TemplatesExtractProvided == 1 and templates:
        print("Matching assets to templates...")
        df = match_assets_to_templates(df, templates)
        # Filter to only include assets that have been matched to templates
        original_count = len(df)
        df = df[df["Template"] != ""].copy()
        matched_count = len(df)
        print(f"Filtered assets: {original_count} -> {matched_count} (only assets matching templates)")
    else:
        # Add empty Template column for consistency
        df["Template"] = ""

    rows = []

    def add_row(parent, name, description="", template=""):
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
            "SecurityString": sec_str,
            "Template": template,
            "Value": ""
        })

    def add_attribute_row(parent_element, attribute_name, attribute_value=""):
        """Add an attribute row for setting attribute values"""
        rows.append({
            "Selected(x)": "x",
            "Parent": parent_element,
            "Name": attribute_name,
            "ObjectType": "Attribute",
            "Error": "",
            "Description": "",
            "SecurityString": "",
            "Template": "",
            "Value": attribute_value
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
    # First, create a mapping of sensor assets for PID controller parent matching
    sensor_assets = {}
    controller_assets = []
    regular_assets = []
    
    if TemplatesExtractProvided == 1 and templates:
        # Separate sensors, controllers, and regular assets for proper ordering
        for _, row in df.iterrows():
            template = str(row["Template"]) if pd.notna(row["Template"]) else ""
            pid_asset = str(row["P&ID Asset"])
            
            if template == "TLS.Analog.Sensor.001":
                sensor_assets[pid_asset] = row
            elif template == "TLS.PID.Controller.001":
                controller_assets.append(row)
            else:
                regular_assets.append(row)
    else:
        # If no templates, treat all as regular assets
        regular_assets = df.to_dict('records')
    
    def find_corresponding_sensor(controller_name):
        """Find corresponding sensor for a PID controller based on name pattern"""
        # Extract the base pattern by replacing 'C' with 'T'
        # Examples: LIC001 -> LIT001, PIC931 -> PIT931
        if len(controller_name) >= 3:
            # Look for pattern where controller has 'C' and sensor has 'T'
            for i, char in enumerate(controller_name):
                if char == 'C':
                    # Try replacing C with T
                    potential_sensor = controller_name[:i] + 'T' + controller_name[i+1:]
                    if potential_sensor in sensor_assets:
                        return potential_sensor
        return None
    
    def get_asset_display_name(pid_asset, asset_name):
        """Get the full display name as it appears in the output"""
        if asset_name:
            return f"{pid_asset} - {asset_name}"
        return pid_asset
    
    # Add assets in proper order: sensors first, then regular assets, then controllers
    assets_to_process = []
    
    # 1. Add sensors first (they can be parents)
    for sensor_name, row in sensor_assets.items():
        assets_to_process.append(('sensor', row))
    
    # 2. Add regular assets 
    for row in regular_assets:
        assets_to_process.append(('regular', row))
    
    # 3. Add controllers last (they may be children of sensors)
    for row in controller_assets:
        assets_to_process.append(('controller', row))
    
    # Process assets in the determined order
    for asset_type, row in assets_to_process:
        l2 = str(row["Level 2"])
        l3 = str(row["Level 3"])
        pid_asset = str(row["P&ID Asset"])
        asset_name = str(row["Asset Name"]) if pd.notna(row["Asset Name"]) else ""
        template = str(row["Template"]) if pd.notna(row["Template"]) else ""
        
        # Default parent path (hierarchy-based)
        parent_path = f"{level1_name}\\{l2}" if l2 and l2.lower() != "nan" else level1_name
        if l3 and l3.lower() != "nan" and str(l3).strip():
            parent_path = f"{parent_path}\\{l3}"
        
        # Special handling for PID Controllers - try to find corresponding sensor as parent
        if asset_type == 'controller' and TemplatesExtractProvided == 1:
            corresponding_sensor = find_corresponding_sensor(pid_asset)
            if corresponding_sensor:
                # Get the sensor's full display name and use it as parent
                sensor_row = sensor_assets[corresponding_sensor]
                sensor_asset_name = str(sensor_row["Asset Name"]) if pd.notna(sensor_row["Asset Name"]) else ""
                sensor_display_name = get_asset_display_name(corresponding_sensor, sensor_asset_name)
                parent_path = f"{parent_path}\\{sensor_display_name}"
                print(f"Controller '{pid_asset}' will be child of sensor '{sensor_display_name}'")
        
        add_row(parent=parent_path, name=pid_asset, description=asset_name or "", template=template)
        
        # Add SCADA Asset Name attribute if this asset has a template
        if template and template.strip() and TemplatesExtractProvided == 1:
            # Get SCADA Asset value from the mapping we created earlier
            scada_asset_value = scada_asset_mapping.get(pid_asset, "")
            
            # Create the element display name (same as what add_row creates)
            element_display_name = pid_asset
            if asset_name:
                element_display_name = f"{pid_asset} - {asset_name}"
                
            # The parent for the attribute should be the full path to the element
            full_element_path = f"{parent_path}\\{element_display_name}"
                
            # Add the SCADA Asset Name attribute row
            add_attribute_row(full_element_path, "SCADA Asset Name", scada_asset_value)

    out_df = pd.DataFrame(rows, columns=[
        "Selected(x)","Parent","Name","ObjectType","Error","Description","SecurityString","Template","Value"
    ])

    # Save to Excel
    out_df.to_excel(output_path, index=False)

if __name__ == "__main__":
    tls_path = "data/TLS - Tags for AF rev 1.xlsx"
    book2_path = "data/Book2.xlsx"
    output_path = "data/TLS_AF_Import.xlsx"
    build_af_import(tls_path, book2_path, output_path)
    print(f"Written: {output_path}")
