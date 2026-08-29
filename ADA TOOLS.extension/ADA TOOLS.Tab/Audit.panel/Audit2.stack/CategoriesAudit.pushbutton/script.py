# -*- coding: utf-8 -*-
__doc__ = """Count every model element (excluding element types) in the project,
grouped by category, sorted highest count first.

Results are printed to the pyRevit console and exported to a CSV
file - a quick way to see what dominates the model (e.g. how many
walls vs. generic models vs. pipes) without opening a schedule."""

__title__ = "Categories Audit"
__author__ = "Almog Davidson"

from pyrevit import revit, DB, forms, script
import datetime
import os
import csv

# Shared ADA-Tools dark/gold themed report (see lib/GUI/ReportTheme.py)
from GUI.ReportTheme import ADAReport

# Get current document
doc = revit.doc

folder_name = doc.Title

# Prepare output
output = script.get_output()

def Revit_Categories():
    results_categories = []

    # Get all model elements (not element types)
    collector = DB.FilteredElementCollector(doc).WhereElementIsNotElementType()
    category_counts = {}
    total_count = 0

    # Count elements by category
    for element in collector:
        if element.Category and element.Category.CategoryType == DB.CategoryType.Model:
            cat_name = element.Category.Name
            if cat_name in category_counts:
                category_counts[cat_name] += 1
            else:
                category_counts[cat_name] = 1
            total_count += 1

    # Sort by count (highest first)
    sorted_categories = sorted(category_counts.items(), key=lambda x: x[1], reverse=True)

    for category, count in sorted_categories:
        results_categories.append([category, count])

    report = ADAReport(__title__)
    report.line("Total model elements: <b>{}</b> across <b>{}</b> categories".format(
        total_count, len(category_counts)))
    report.table(["Category", "Element Count"],
                 [[category, str(count)] for category, count in sorted_categories])
    report.flush()

    # Empty row for spacing, kept for CSV export layout
    results_categories.insert(0, ["", ""])

    return results_categories

def save_categories_to_csv(results_categories):
    """Save the category audit results to a CSV file"""

    # Define the output folder - change this path as needed
    output_folder = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker\{}".format(folder_name)

    # Create the folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    filename = "Revit_Categories_Audit.csv"
    filepath = os.path.join(output_folder, filename)

    try:
        # Open file for writing
        with open(filepath, 'w') as csvfile:
            writer = csv.writer(csvfile)
            
            # Write timestamp header
            writer.writerow(["Revit Categories Audit Report"])
            writer.writerow(["Generated on: {}".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))])
            writer.writerow(["Document: {}".format(doc.Title or "Unknown")])
            writer.writerow([])  # Empty row for spacing
            
            # Write column headers
            writer.writerow(["Category Name", "Element Count"])
            
            # Write data rows
            for result in results_categories:
                writer.writerow(result)

        ADAReport(__title__).success("CSV report saved to: <b>{}</b>".format(filepath)).flush()
        return filepath

    except Exception as e:
        ADAReport(__title__).error("Error saving CSV file: {}".format(str(e))).flush()
        return None

if __name__ == '__main__':
    # Run the category audit and get results
    results_categories = Revit_Categories()

    # Save results to CSV
    save_categories_to_csv(results_categories)
