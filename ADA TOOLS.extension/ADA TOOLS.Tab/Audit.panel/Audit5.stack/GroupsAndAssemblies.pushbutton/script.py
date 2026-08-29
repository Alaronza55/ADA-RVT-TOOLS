__doc__ = """Count Groups and Assemblies
Displays the total number of Groups and Assemblies in the current project and exports to CSV."""

__title__ = "Count Groups\nand Assemblies"
__author__ = "Almog Davidson"

from Autodesk.Revit.DB import FilteredElementCollector, Group, AssemblyInstance, BuiltInCategory
from pyrevit import forms, script
import csv
import os

# Shared ADA-Tools dark/gold themed report (see lib/GUI/ReportTheme.py)
from GUI.ReportTheme import ADAReport

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

# Collect all Groups
groups = FilteredElementCollector(doc)\
    .OfClass(Group)\
    .WhereElementIsNotElementType()\
    .ToElements()

# Collect all Assemblies
assemblies = FilteredElementCollector(doc)\
    .OfClass(AssemblyInstance)\
    .ToElements()

# Prepare data for CSV
csv_data = []

# Process Groups
for group in groups:
    group_name = group.Name
    group_id = group.Id.IntegerValue
    
    # Get the group type to check its category
    group_type_id = group.GetTypeId()
    group_type_elem = doc.GetElement(group_type_id)
    
    # Check the category of the group type
    if group_type_elem and group_type_elem.Category:
        category_id = group_type_elem.Category.Id.IntegerValue
        
        # Model groups have category OST_IOSModelGroups (-2000095)
        # Detail groups have category OST_IOSDetailGroups (-2000095)
        if category_id == int(BuiltInCategory.OST_IOSDetailGroups):
            group_type = "Detail Group"
        elif category_id == int(BuiltInCategory.OST_IOSModelGroups):
            group_type = "Model Group"
        else:
            group_type = "Unknown Group (Category: {})".format(category_id)
    else:
        group_type = "Unknown Group (No Category)"
    
    csv_data.append([group_name, group_type, group_id])

# Process Assemblies
for assembly in assemblies:
    assembly_name = assembly.AssemblyTypeName
    assembly_id = assembly.Id.IntegerValue
    csv_data.append([assembly_name, "Assembly", assembly_id])

# Display Results
model_groups = [d for d in csv_data if d[1] == "Model Group"]
detail_groups = [d for d in csv_data if d[1] == "Detail Group"]
assemblies_list = [d for d in csv_data if d[1] == "Assembly"]

report = ADAReport(__title__.replace(chr(10), " "))

report.subheader("Groups")
report.line("Total Groups: <b>{}</b>".format(len(groups)))
report.line("Model Groups: <b>{}</b>".format(len(model_groups)))
report.line("Detail Groups: <b>{}</b>".format(len(detail_groups)))

report.subheader("Assemblies")
report.line("Total Assemblies: <b>{}</b>".format(len(assemblies_list)))

report.subheader("Grand Total")
report.line("<b>{}</b>".format(len(csv_data)))

# Export to CSV - ALWAYS export, even if empty
folder_name = doc.Title
output_folder = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker\{}".format(folder_name)

# Create the folder if it doesn't exist
if not os.path.exists(output_folder):
    try:
        os.makedirs(output_folder)
        report.line("Created folder: <b>{}</b>".format(output_folder))
    except Exception as e:
        report.warn("Error creating folder: {}".format(str(e)))
        report.warn("Attempting to save to default location...")
        output_folder = os.path.expanduser("~\\Desktop")

# Create filename with project name
project_name = doc.Title
csv_filename = "GroupsAndAssemblies.csv"
csv_filepath = os.path.join(output_folder, csv_filename)

# Write CSV file
try:
    with open(csv_filepath, 'wb') as csvfile:
        writer = csv.writer(csvfile)
        # Write header
        writer.writerow(['Name', 'Type', 'ID'])
        # Write data (or empty if no data)
        if csv_data:
            writer.writerows(csv_data)
        else:
            # Write a note that no groups/assemblies were found
            writer.writerow(['No groups or assemblies found', '', ''])

    report.subheader("CSV Export")
    report.success("CSV file exported successfully!")
    report.line("Location: <b>{}</b>".format(csv_filepath))
    if not csv_data:
        report.warn("No groups or assemblies found in the project.")
    report.flush()

except Exception as e:
    report.subheader("CSV Export - ERROR")
    report.error("Failed to save CSV file: {}".format(str(e)))
    report.line("Attempted location: {}".format(csv_filepath))
    report.flush()