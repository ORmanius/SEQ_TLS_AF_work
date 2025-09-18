"""
025_AFtemplateGeneration.py

This script imports the JSON template specification file created by 020_TemplateExtraction.py
and creates a CSV file formatted for importing templates into AVEVA Asset Framework
via the PI Builder Excel Add-on.

The CSV file format matches the structure in TemplateExcelFileExample.csv.
"""

import pandas as pd
import json
from pathlib import Path
from datetime import datetime

def load_template_json(json_path):
    """Load the JSON template specification file"""
    with open(json_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def create_pi_builder_excel(templates_data, output_path):
    """
    Create a CSV file formatted for PI Builder Excel Add-on import
    
    The format matches the structure in TemplateExcelFileExample.csv:
    - First row: headers
    - Each template starts with an ElementTemplate row
    - Followed by AttributeTemplate rows for that template's attributes
    """
    
    # Create the data rows following the example structure
    data_rows = []
    
    # Add header row
    headers = [
        "Selected(x)", "Parent", "Name", "ObjectType", "Error", "Description", "SecurityString",
        "Type", "AllowElementToExtend", "BaseTemplateOnly", "Categories", "AttributeIsHidden",
        "AttributeIsManualDataEntry", "AttributeIsConfigurationItem", "AttributeIsExcluded",
        "AttributeIsIndexed", "AttributeDefaultUOM", "AttributeType", "AttributeDefaultValue",
        "AttributeDataReference", "AttributeConfigString", "AttributeDisplayDigits", "", ""
    ]
    data_rows.append(headers)
    
    # Process each template
    for template in templates_data["templates"]:
        # Add ElementTemplate row
        template_row = create_element_template_row(template)
        data_rows.append(template_row)
        
        # Add AttributeTemplate rows for this template
        for attr in template["attributes"]:
            attr_row = create_attribute_template_row(template, attr)
            data_rows.append(attr_row)
        
        # Add empty separator rows
        data_rows.append([""] * len(headers))
        data_rows.append([""] * len(headers))
    
    # Create DataFrame and save as CSV
    df = pd.DataFrame(data_rows)
    df.to_csv(output_path, index=False, header=False)
    print(f"PI Builder CSV file saved to: {output_path}")

def create_element_template_row(template):
    """Create a row for ElementTemplate definition"""
    
    # Create security string similar to example
    security_string = "Administrators:A(r,w,rd,wd,d,x,a,s,so,an)|Engineers:A(r,w,rd,wd,d,x,s,so,an)|World:A(r,rd)|Asset Analytics:A(r,w,rd,wd,x,an)|Asset Analytics Recalculation:A(x)|RTQP Engine:A(r,rd)"
    
    # Create categories string
    categories = f"Basic Asset;TLS;{template['name']};"
    
    # Template name with TLS prefix
    template_name = f"TLS.{template['name'].replace(' ', '.')}.001"
    
    row = [
        "x",  # Selected(x)
        "",   # Parent (empty for ElementTemplate)
        template_name,  # Name
        "ElementTemplate",  # ObjectType
        "",   # Error
        f"{template['name']} Template",  # Description
        security_string,  # SecurityString
        "None",  # Type
        "FALSE",  # AllowElementToExtend
        "FALSE",  # BaseTemplateOnly
        categories,  # Categories
        "",   # AttributeIsHidden
        "",   # AttributeIsManualDataEntry
        "",   # AttributeIsConfigurationItem
        "",   # AttributeIsExcluded
        "",   # AttributeIsIndexed
        "",   # AttributeDefaultUOM
        "",   # AttributeType
        "",   # AttributeDefaultValue
        "",   # AttributeDataReference
        "",   # AttributeConfigString
        "",   # AttributeDisplayDigits
        "",   # Empty column
        ""    # Empty column
    ]
    
    return row

def create_attribute_template_row(template, attr):
    """Create a row for AttributeTemplate definition"""
    
    # Template name with TLS prefix
    template_name = f"TLS.{template['name'].replace(' ', '.')}.001"
    
    # Map data types
    attr_type_map = {
        "Boolean": "Boolean",
        "Float64": "Double", 
        "Int32": "Int32",
        "String": "String",
        "DateTime": "DateTime"
    }
    
    attr_type = attr_type_map.get(attr["data_type"], "Double")
    
    # Default value based on type
    default_value = "FALSE" if attr_type == "Boolean" else "0"
    
    # Create data reference similar to example format
    # Convert substitution pattern to match example: \\%@\zzz.GlobalConfiguration|Data Archive Name%\TLS_%Element%<attribute_suffix>
    attr_suffix = attr["name"].title()  # Use full attribute name and capitalize
    data_reference = f"\\\\%@\\zzz.GlobalConfiguration|Data Archive Name%\\TLS_%Element%{attr_suffix}"
    
    row = [
        "x",  # Selected(x)
        template_name,  # Parent (template name)
        attr["name"].title(),  # Name (capitalize attribute name)
        "AttributeTemplate",  # ObjectType
        "",   # Error
        attr["description"],  # Description
        "",   # SecurityString (empty for attributes)
        "",   # Type (empty for attributes)
        "",   # AllowElementToExtend (empty for attributes)
        "",   # BaseTemplateOnly (empty for attributes)
        "",   # Categories (empty for attributes)
        "FALSE",  # AttributeIsHidden
        "FALSE",  # AttributeIsManualDataEntry
        "FALSE",  # AttributeIsConfigurationItem
        "FALSE",  # AttributeIsExcluded
        "FALSE",  # AttributeIsIndexed
        attr["engineering_units"],  # AttributeDefaultUOM
        attr_type,  # AttributeType
        default_value,  # AttributeDefaultValue
        "PI Point",  # AttributeDataReference
        data_reference,  # AttributeConfigString
        "-5",  # AttributeDisplayDigits
        "",   # Empty column
        ""    # Empty column
    ]
    
    return row

def main():
    """Main function to process JSON and create PI Builder CSV file"""
    
    # Input JSON file path (created by 020_TemplateExtraction.py)
    json_path = Path("AF_Templates_Specification.json")
    
    # Output CSV file path - save to data folder
    output_folder = Path("data")
    output_folder.mkdir(exist_ok=True)
    output_path = output_folder / f"AF_Templates_PIBuilder_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    
    try:
        # Load JSON template data
        print(f"Loading template data from: {json_path}")
        templates_data = load_template_json(json_path)
        
        print(f"Found {len(templates_data['templates'])} templates")
        
        # Create PI Builder CSV file
        print("Creating PI Builder CSV file...")
        create_pi_builder_excel(templates_data, output_path)
        
        # Print summary
        print("\n=== Template Import File Created ===")
        print(f"Output file: {output_path}")
        print(f"Templates included: {len(templates_data['templates'])}")
        
        total_attributes = sum(len(template['attributes']) for template in templates_data['templates'])
        print(f"Total attributes: {total_attributes}")
        
        print("\nTemplate Summary:")
        for template in templates_data['templates']:
            template_name = f"TLS.{template['name'].replace(' ', '.')}.001"
            print(f"  {template_name}: {len(template['attributes'])} attributes")
        
        print(f"\nCSV file ready for import via PI Builder Excel Add-on!")
        print("Format matches the structure in TemplateExcelFileExample.csv")
        
    except FileNotFoundError:
        print(f"Error: JSON file not found at {json_path}")
        print("Please run 020_TemplateExtraction.py first to generate the JSON file.")
    except Exception as e:
        print(f"Error processing templates: {e}")

if __name__ == "__main__":
    main()