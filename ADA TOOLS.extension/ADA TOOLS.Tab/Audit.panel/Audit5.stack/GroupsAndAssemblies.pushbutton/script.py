"""Count Groups and Assemblies
Displays the total number of Groups and Assemblies in the current project and exports to CSV."""

__title__ = "Count Groups\nand Assemblies"
__author__ = "Almog Davidson"

from Autodesk.Revit.DB import FilteredElementCollector, Group, AssemblyInstance
from pyrevit import forms, script
import csv
import os

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
    # Determine if Model or Detail Group
    if group.GroupType.GetType().Name == "GroupType":
        group_type = "Model Group"
    else:
        group_type = "Detail Group"
    
    csv_data.append([group_name, group_type])

# Process Assemblies
for assembly in assemblies:
    assembly_name = assembly.AssemblyTypeName
    csv_data.append([assembly_name, "Assembly"])

# Display Results
model_groups = [d for d in csv_data if d[1] == "Model Group"]
detail_groups = [d for d in csv_data if d[1] == "Detail Group"]
assemblies_list = [d for d in csv_data if d[1] == "Assembly"]

output.print_md("## Groups and Assemblies Count")
output.print_md("---")

output.print_md("### Groups")
print("Total Groups: {}".format(len(groups)))
print("  - Model Groups: {}".format(len(model_groups)))
print("  - Detail Groups: {}".format(len(detail_groups)))

output.print_md("\n### Assemblies")
print("Total Assemblies: {}".format(len(assemblies_list)))

output.print_md("\n---")
output.print_md("**Grand Total: {}**".format(len(csv_data)))

# Export to CSV
folder_name = doc.Title
output_folder = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker\{}".format(folder_name)

if csv_data:
    # Create the folder if it doesn't exist
    if not os.path.exists(output_folder):
        try:
            os.makedirs(output_folder)
            print("**Created folder:** `{}`".format(output_folder))
        except Exception as e:
            print("**Error creating folder:** {}".format(str(e)))
            print("**Attempting to save to default location...**")
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
            writer.writerow(['Name', 'Type'])
            # Write data
            writer.writerows(csv_data)

        output.print_md("\n---")
        output.print_md("### CSV Export")
        print("CSV file exported successfully!")
        print("Location: {}".format(csv_filepath))
        
    except Exception as e:
        output.print_md("\n---")
        output.print_md("### CSV Export - ERROR")
        print("Failed to save CSV file!")
        print("Error: {}".format(str(e)))
        print("Attempted location: {}".format(csv_filepath))

else:
    print("\nNo groups or assemblies found in the project.")
