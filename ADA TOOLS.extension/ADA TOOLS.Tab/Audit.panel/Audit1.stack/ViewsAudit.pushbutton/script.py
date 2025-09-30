# -*- coding: utf-8 -*-
"""Count Views and Placed Views
This script counts total views in the project and views placed on sheets.
"""

__title__ = "Views"
__author__ = "Almog Davidson"
__doc__ = """Views Audit Information about the current Revit project."""

from pyrevit import revit, DB, script, forms
import os
import datetime
import re
import csv

# Get the current document
doc = revit.doc

folder_name = doc.Title

def Views():
        results_views = []
        print("=== Views Audit of Project ===")
        # Get all views in the project (excluding view templates)
        all_views = DB.FilteredElementCollector(doc)\
                      .OfClass(DB.View)\
                      .WhereElementIsNotElementType()\
                      .ToElements()

        # Filter out view templates
        views = [view for view in all_views if not view.IsTemplate]

        # Get all sheets in the project
        all_sheets = DB.FilteredElementCollector(doc)\
                       .OfClass(DB.ViewSheet)\
                       .WhereElementIsNotElementType()\
                       .ToElements()

        # Get all viewports (views placed on sheets)
        all_viewports = DB.FilteredElementCollector(doc)\
                          .OfClass(DB.Viewport)\
                          .ToElements()

        # Get unique view IDs that are placed on sheets
        placed_view_ids = set()
        for viewport in all_viewports:
            view_id = viewport.ViewId
            placed_view_ids.add(view_id)

        # Count totals
        total_views = len(views)
        placed_views_count = len(placed_view_ids)
        unplaced_views_count = total_views - placed_views_count

        # Display results
        output = script.get_output()

        # print(" Total Views : **{0}** |".format(total_views))
        print(" Views on Sheets :  **{0}** |".format(placed_views_count))
        print(" Views NOT on Sheets :  **{0}** |".format(unplaced_views_count))
        # print(" Total Sheets :  **{0}** |".format(len(all_sheets)))
        print("")

        # Store results for CSV export
        results_views.append(["Views", "Count"])
        # results_views.append(["Total Views", total_views])
        results_views.append(["Views on Sheets", placed_views_count])
        results_views.append(["Views NOT on Sheets", unplaced_views_count])
        # results_views.append(["Total Sheets", len(all_sheets)])

        return results_views

def save_to_csv(results_views):
    """Save the audit results to a CSV file"""

    # Define the output folder - change this path as needed
    output_folder = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker\{}".format(folder_name)
    
    # Create the folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Clean document title for filename (remove invalid characters)
    doc_title = doc.Title or "Unknown"
    clean_title = re.sub(r'[<>:"/\\|?*]', '_', doc_title)

    filename = "Audit_Views.csv"
    filepath = os.path.join(output_folder, filename)

    try:
        # Open file for writing
        with open(filepath, 'w') as csvfile:
            writer = csv.writer(csvfile)
            
            # Write header information
            writer.writerow(["Views Audit Report"])
            writer.writerow(["Generated on: {}".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))])
            writer.writerow(["Document: {}".format(doc.Title or "Unknown")])
            writer.writerow([])  # Empty row for spacing
            
            # Write data rows
            for result in results_views:
                writer.writerow(result)

        print("CSV report saved to: {}".format(filepath))
        return filepath
        
    except Exception as e:
        print("Error saving CSV file: {}".format(str(e)))
        return None

if __name__ == '__main__':
    # Run the audit and get results
    results_views = Views()

    # Save results to CSV
    save_to_csv(results_views)
