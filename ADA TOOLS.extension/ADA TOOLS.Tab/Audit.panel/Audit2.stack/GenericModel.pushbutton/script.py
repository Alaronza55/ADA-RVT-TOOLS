# -*- coding: utf-8 -*-
"""List All Worksets - Simple Table
Lists all worksets in the current project in a simple table format.
"""

__title__ = "Generic Models Export"
__author__ = "Almog Davidson"

from Autodesk.Revit.DB import *
from pyrevit import revit, DB, forms, script
# Get current document
doc = revit.doc

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
import csv
import os

# Get the current Revit document
doc = __revit__.ActiveUIDocument.Document

folder_name = doc.Title

try:
    # Collect all generic model elements
    collector = FilteredElementCollector(doc)
    generic_models = collector.OfCategory(BuiltInCategory.OST_GenericModel).WhereElementIsNotElementType().ToElements()
    
    # Prepare data for CSV
    csv_data = []
    csv_data.append(['Family Name', 'Type Name', 'Element ID'])  # Header
    
    for element in generic_models:
        try:
            # Get element type
            element_type = doc.GetElement(element.GetTypeId())
            
            # Get family name - use different approach
            family_name = 'N/A'
            if element_type:
                try:
                    family_name = element_type.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString()
                except:
                    try:
                        family_name = element_type.Family.Name
                    except:
                        family_name = 'N/A'
            
            # Get type name - use parameter instead of direct property
            type_name = 'N/A'
            if element_type:
                try:
                    type_name = element_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                except:
                    try:
                        type_name = element_type.Name
                    except:
                        type_name = 'N/A'
            
            # Get element ID
            element_id = element.Id.IntegerValue
            
            # Add to data
            csv_data.append([family_name, type_name, element_id])
            
        except Exception as e:
            print("Error processing element {}: {}".format(element.Id, str(e)))
            # Add the element with N/A values instead of skipping
            csv_data.append(['N/A', 'N/A', element.Id.IntegerValue])
            continue
    
    # Create CSV file path - use your specific path
    output_folder = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker\{}".format(folder_name)
    
    # Check if directory exists, if not create it
    if not os.path.exists(output_folder):
        try:
            os.makedirs(output_folder)
            print("Created directory: {}".format(output_folder))
        except Exception as e:
            print("Could not create directory: {}".format(str(e)))
            # Fallback to temp directory
            import tempfile
            csv_folder = tempfile.gettempdir()
            print("Using temp directory instead: {}".format(csv_folder))
    
    csv_file_path = os.path.join(output_folder, 'Generic_Models_Export.csv')
    
    # Write to CSV file
    with open(csv_file_path, 'w') as csvfile:
        writer = csv.writer(csvfile)
        for row in csv_data:
            # Convert all items to strings to avoid encoding issues
            string_row = [str(item) for item in row]
            writer.writerow(string_row)
    
    print("CSV export completed successfully!")
    print("File saved to: {}".format(csv_file_path))
    print("Total generic models exported: {}".format(len(csv_data) - 1))  # -1 for header
    
except Exception as e:
    print("Error writing CSV file: {}".format(str(e)))
    # Try alternative path
    try:
        import tempfile
        csv_file_path = os.path.join(tempfile.gettempdir(), 'Generic_Models_Export.csv')
        
        with open(csv_file_path, 'w') as csvfile:
            writer = csv.writer(csvfile)
            for row in csv_data:
                string_row = [str(item) for item in row]
                writer.writerow(string_row)
        
        print("CSV export completed successfully (saved to temp folder)!")
        print("File saved to: {}".format(csv_file_path))
        print("Total generic models exported: {}".format(len(csv_data) - 1))
        
    except Exception as e2:
        print("Failed to save file: {}".format(str(e2)))