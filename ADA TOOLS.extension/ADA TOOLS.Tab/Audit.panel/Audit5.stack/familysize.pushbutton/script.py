from Autodesk.Revit.DB import FilteredElementCollector, Family, SaveAsOptions, FamilyInstance
import os
import tempfile
import csv

__title__ = "Family Sizes"
__author__ = "Almog Davidson"

doc = __revit__.ActiveUIDocument.Document

families = FilteredElementCollector(doc).OfClass(Family).ToElements()

# Get all placed family instances to check which families are used
placed_instances = FilteredElementCollector(doc).OfClass(FamilyInstance).ToElements()
placed_family_ids = set(instance.Symbol.Family.Id for instance in placed_instances)

temp_dir = tempfile.gettempdir()
family_data = []

print("{:<40} {:<20} {:>15} {:>10} {:>10}".format("Family Name", "Category", "Size (KB)", "Types", "Placed"))
print("-" * 100)

for family in families:
    family_name = family.Name
    num_types = len(family.GetFamilySymbolIds())
    is_placed = "Yes" if family.Id in placed_family_ids else "No"
    
    # Get family category
    family_category = family.FamilyCategory.Name if family.FamilyCategory else "Unknown"

    if family.IsEditable and not family.IsInPlace:
        try:
            # Open family document
            family_doc = doc.EditFamily(family)

            # Create temp file path
            temp_path = os.path.join(temp_dir, "{}.rfa".format(family_name.replace(" ", "_")))

            # Save family
            save_opts = SaveAsOptions()
            save_opts.OverwriteExistingFile = True
            save_opts.Compact = True

            family_doc.SaveAs(temp_path, save_opts)

            # Get size in KB and MB (decimal)
            size_bytes = os.path.getsize(temp_path)
            size_kb = size_bytes / 1000.0  # Bytes to KB (decimal)
            size_mb = size_kb / 1000.0     # KB to MB (decimal)

            family_data.append({
                'Family Name': family_name,
                'Category': family_category,
                'Size (KB)': size_kb,
                'Size (MB)': size_mb,
                'Number of Types': num_types,
                'Placed in Project': is_placed
            })

            print("{:<40} {:<20} {:>15.2f} {:>10} {:>10}".format(
                family_name[:40], family_category[:20], size_kb, num_types, is_placed))

            # Close family document
            family_doc.Close(False)

            # Clean up temp file
            if os.path.exists(temp_path):
                os.remove(temp_path)

        except Exception as e:
            print("{:<40} {:<20} {:>15} {:>10} {:>10}".format(
                family_name[:40], family_category[:20], "Error", num_types, is_placed))

# Sort by size (largest first)
family_data.sort(key=lambda x: x['Size (KB)'], reverse=True)

print("\n" + "=" * 100)
print("SUMMARY - Top 10 Largest Families:")
print("=" * 100)

for i, fam in enumerate(family_data[:10], 1):
    print("{}. {:<35} {:<20} {:>10.2f} MB ({} types) [{}]".format(
        i, fam['Family Name'][:35], fam['Category'][:20], 
        fam['Size (MB)'], fam['Number of Types'], fam['Placed in Project']
    ))

total_mb = sum([f['Size (MB)'] for f in family_data])
placed_count = sum([1 for f in family_data if f['Placed in Project'] == 'Yes'])
unplaced_count = len(family_data) - placed_count

print("\nTotal: {} families | {:.2f} MB".format(len(family_data), total_mb))
print("Placed: {} | Not Placed: {}".format(placed_count, unplaced_count))

# Export to CSV
folder_name = doc.Title
output_folder = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker\{}".format(folder_name)

# Create the folder if it doesn't exist
if not os.path.exists(output_folder):
    try:
        os.makedirs(output_folder)
        print("**Created folder:** `{}`".format(output_folder))
    except Exception as e:
        print("**Error creating folder:** {}".format(str(e)))
        print("**Attempting to save to default location...**")
        output_folder = os.path.expanduser("~\\Desktop")

# Save detailed element breakdown
filename_detailed = "Family_Sizes.csv"
filepath_detailed = os.path.join(output_folder, filename_detailed)

try:
    with open(filepath_detailed, 'wb') as csvfile:
        fieldnames = ['Family Name', 'Category', 'Size (KB)', 'Size (MB)', 'Number of Types', 'Placed in Project']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

        writer.writeheader()
        
        # Write rows with rounded values for display
        for fam in family_data:
            writer.writerow({
                'Family Name': fam['Family Name'],
                'Category': fam['Category'],
                'Size (KB)': round(fam['Size (KB)'], 2),
                'Size (MB)': round(fam['Size (MB)'], 2),
                'Number of Types': fam['Number of Types'],
                'Placed in Project': fam['Placed in Project']
            })

    print("\n" + "=" * 100)
    print("CSV exported successfully to:")
    print(filepath_detailed)
    print("=" * 100)

except Exception as e:
    print("\nError exporting CSV: {}".format(str(e)))