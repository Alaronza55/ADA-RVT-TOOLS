__title__ = "Count Loadable\nFamilies"

from pyrevit import revit, DB
import os
import csv

doc = revit.doc
folder_name = doc.Title

# Get all family symbols (loadable family types)
collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilySymbol)
family_symbols = collector.ToElements()

# Get all family instances placed in the project
instances_collector = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance)
family_instances = instances_collector.ToElements()

# Get loadable families and in-place families
loadable_families = {}  # Changed to dictionary to store category info
in_place_families = {}  # Changed to dictionary to store category info
placed_families = {}    # Changed to dictionary to store category info

for symbol in family_symbols:
    if symbol.Family:
        family = symbol.Family
        category_name = family.FamilyCategory.Name if family.FamilyCategory else "Unknown"

        if family.IsInPlace:
            in_place_families[family.Name] = category_name
        else:
            loadable_families[family.Name] = category_name

# Check which families have instances placed in the project
for instance in family_instances:
    if instance.Symbol and instance.Symbol.Family:
        family = instance.Symbol.Family
        if not family.IsInPlace:  # Only count loadable families
            category_name = family.FamilyCategory.Name if family.FamilyCategory else "Unknown"
            placed_families[family.Name] = category_name

# Count totals
total_symbols = len(family_symbols)
total_loadable = len(loadable_families)
total_in_place = len(in_place_families)
total_placed = len(placed_families)
total_unplaced = total_loadable - total_placed

# Display results
print("FAMILIES COUNT")
print("  - Loadable Families: {}".format(total_loadable))
print("  - In-Place Families: {}".format(total_in_place))
print("  - Placed Loadable Families: {}".format(total_placed))
print("  - Unplaced Loadable Families: {}".format(total_unplaced))
print("")

# Export to CSV
def export_to_csv():
    try:
        # Get project name
        project_name = doc.Title if doc.Title else "Untitled"

        # Create output folder path
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
                output_folder = tempfile.gettempdir()
                print("Using temp directory instead: {}".format(output_folder))

        # Export Summary CSV
        summary_filename = "Family_Count_Summary.csv"
        summary_filepath = os.path.join(output_folder, summary_filename)
        
        with open(summary_filepath, 'wb') as csvfile:
            writer = csv.writer(csvfile)
            
            # Write summary data
            writer.writerow(["Metric", "Count"])
            writer.writerow(["Loadable Families", total_loadable])
            writer.writerow(["In-Place Families", total_in_place])
            writer.writerow(["Placed Loadable Families", total_placed])
            writer.writerow(["Unplaced Loadable Families", total_unplaced])

        print("Summary CSV exported successfully to: {}".format(summary_filepath))

        # Export Details CSV
        details_filename = "Family_Count_Details.csv"
        details_filepath = os.path.join(output_folder, details_filename)
        
        # Create unplaced families dictionary
        unplaced_families = {}
        for family_name, category in loadable_families.items():
            if family_name not in placed_families:
                unplaced_families[family_name] = category

        # Sort families by name
        sorted_placed = sorted(placed_families.items())
        sorted_unplaced = sorted(unplaced_families.items())
        sorted_inplace = sorted(in_place_families.items())

        with open(details_filepath, 'wb') as csvfile:
            writer = csv.writer(csvfile)

            # Write column headers
            writer.writerow(["PLACED FAMILIES", "CATEGORY PLACED", "UNPLACED FAMILIES", "CATEGORY UNPLACED", "IN-PLACE FAMILIES", "CATEGORY IN-PLACE"])

            # Find the maximum length among all lists
            max_length = max(len(sorted_placed), len(sorted_unplaced), len(sorted_inplace))

            # Write all family data in columns
            for i in range(max_length):
                row = []

                # Placed families column
                if i < len(sorted_placed):
                    row.extend([sorted_placed[i][0], sorted_placed[i][1]])  # Name, Category
                else:
                    row.extend(["", ""])  # Empty cells

                # Unplaced families column
                if i < len(sorted_unplaced):
                    row.extend([sorted_unplaced[i][0], sorted_unplaced[i][1]])  # Name, Category
                else:
                    row.extend(["", ""])  # Empty cells

                # In-place families column
                if i < len(sorted_inplace):
                    row.extend([sorted_inplace[i][0], sorted_inplace[i][1]])  # Name, Category
                else:
                    row.extend(["", ""])  # Empty cells

                writer.writerow(row)

        print("Details CSV exported successfully to: {}".format(details_filepath))

    except Exception as e:
        print("Error exporting CSV files: {}".format(str(e)))

# Call the export function
export_to_csv()
