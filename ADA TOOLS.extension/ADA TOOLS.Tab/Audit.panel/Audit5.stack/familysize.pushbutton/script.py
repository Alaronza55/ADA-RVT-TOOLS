from Autodesk.Revit.DB import FilteredElementCollector, Family, SaveAsOptions
import os
import tempfile
import csv

doc = __revit__.ActiveUIDocument.Document

families = FilteredElementCollector(doc).OfClass(Family).ToElements()

temp_dir = tempfile.gettempdir()
family_data = []

print("{:<50} {:>15} {:>10}".format("Family Name", "Size (KB)", "Types"))
print("-" * 80)

for family in families:
    family_name = family.Name
    num_types = len(family.GetFamilySymbolIds())

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

            # Get size in bytes first - using decimal (1000) throughout
            size_bytes = os.path.getsize(temp_path)
            size_kb = size_bytes / 1000.0  # Bytes to KB (decimal)
            size_mb = size_kb / 1000.0     # KB to MB (decimal)

            family_data.append({
                'Family Name': family_name,
                'Size (Bytes)': size_bytes,
                'Size (KB)': size_kb,
                'Size (MB)': size_mb,
                'Number of Types': num_types
            })

            print("{:<50} {:>15.2f} {:>10}".format(family_name, size_kb, num_types))

            # Close family document
            family_doc.Close(False)

            # Clean up temp file
            if os.path.exists(temp_path):
                os.remove(temp_path)

        except Exception as e:
            print("{:<50} {:>15} {:>10}".format(family_name[:50], "Error", num_types))

# Sort by size (largest first)
family_data.sort(key=lambda x: x['Size (KB)'], reverse=True)

print("\n" + "=" * 80)
print("SUMMARY - Top 10 Largest Families:")
print("=" * 80)

for i, fam in enumerate(family_data[:10], 1):
    print("{}. {:<45} {:>10.2f} MB ({} types)".format(
        i, fam['Family Name'][:45], fam['Size (MB)'], fam['Number of Types']
    ))

total_mb = sum([f['Size (MB)'] for f in family_data])
print("\nTotal: {} families | {:.2f} MB".format(len(family_data), total_mb))

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
        fieldnames = ['Family Name', 'Size (KB)', 'Size (MB)', 'Number of Types']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

        writer.writeheader()
        
        # Write rows with rounded values for display
        for fam in family_data:
            writer.writerow({
                'Family Name': fam['Family Name'],
                'Size (KB)': round(fam['Size (KB)'], 2),
                'Size (MB)': round(fam['Size (MB)'], 4),
                'Number of Types': fam['Number of Types']
            })

    print("\n" + "=" * 80)
    print("CSV exported successfully to:")
    print(filepath_detailed)
    print("=" * 80)

except Exception as e:
    print("\nError exporting CSV: {}".format(str(e)))
