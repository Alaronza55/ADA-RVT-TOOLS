# -*- coding: utf-8 -*-
"""
Lists all categories in the current project in a simple table format.
"""

__title__ = "Categories Audit"
__author__ = "Almog Davidson"

from pyrevit import revit, DB, forms, script
import datetime
import os
import csv

# Get current document
doc = revit.doc

folder_name = doc.Title

# Prepare output
output = script.get_output()

def Revit_Categories():
    results_categories = []
    print("=== Revit Categories in Project ===")
    
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

    # Display results
    print("MODEL ELEMENTS COUNT")
    # print("Total Model Elements: {}".format(total_count))
    # print("Categories Found: {}".format(len(category_counts)))
    print("-" * 50)

    # Add summary info to results
    # results_categories.append(["Total Model Elements", total_count])
    # results_categories.append(["Categories Found", len(category_counts)])
    results_categories.append(["", ""])  # Empty row for spacing

    for category, count in sorted_categories:
        print("{}: {}".format(category, count))
        results_categories.append([category, count])
    
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

        print("CSV report saved to: {}".format(filepath))
        return filepath
        
    except Exception as e:
        print("Error saving CSV file: {}".format(str(e)))
        return None

if __name__ == '__main__':
    # Run the category audit and get results
    results_categories = Revit_Categories()

    # Save results to CSV
    save_categories_to_csv(results_categories)
