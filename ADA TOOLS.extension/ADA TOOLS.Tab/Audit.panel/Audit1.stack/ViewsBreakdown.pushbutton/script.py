# -*- coding: utf-8 -*-
"""Views Breakdown by Type
This script provides a detailed breakdown of views by type, showing which are placed on sheets.
"""

__title__ = "Views Breakdown"
__author__ = "Almog Davidson"

from pyrevit import revit, DB, script
import os
import datetime
import re
import csv

# Get the current document
doc = revit.doc

folder_name = doc.Title

def Views_Breakdown():
    # Get all views in the project (excluding view templates)
    all_views = DB.FilteredElementCollector(doc)\
                  .OfClass(DB.View)\
                  .WhereElementIsNotElementType()\
                  .ToElements()

    # Filter out view templates
    views = [view for view in all_views if not view.IsTemplate]

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

    # Prepare detailed breakdown by view type
    view_types_breakdown = {}
    placed_view_types_breakdown = {}

    for view in views:
        view_type_name = view.ViewType.ToString()

        # Count all views by type
        if view_type_name not in view_types_breakdown:
            view_types_breakdown[view_type_name] = 0
        view_types_breakdown[view_type_name] += 1

        # Count placed views by type
        if view.Id in placed_view_ids:
            if view_type_name not in placed_view_types_breakdown:
                placed_view_types_breakdown[view_type_name] = 0
            placed_view_types_breakdown[view_type_name] += 1

    # Main summary
    print("=== View Count Summary ===")
    print("| View Type | On Sheets | Not on Sheets |")

    results_views = []
    for view_type in sorted(view_types_breakdown.keys()):
        total_count = view_types_breakdown[view_type]
        placed_count = placed_view_types_breakdown.get(view_type, 0)
        unplaced_count = total_count - placed_count

        print("| {0} | {1} | {2} |".format(
            view_type, 
            placed_count, 
            unplaced_count
        ))
        
        # Add to results for CSV export
        results_views.append([view_type, placed_count, unplaced_count, total_count])

    return results_views

def save_to_csv(results_views):
    """Save the views breakdown results to a CSV file"""
    
    # Define the output folder
    output_folder = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker\{}".format(folder_name)

    # Create the folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Clean document title for filename
    doc_title = doc.Title or "Unknown"
    clean_title = re.sub(r'[<>:"/\\|?*]', '_', doc_title)

    filename = "Views_Breakdown.csv"
    filepath = os.path.join(output_folder, filename)

    try:
        with open(filepath, 'w') as csvfile:
            writer = csv.writer(csvfile)
            
            # Write timestamp header
            writer.writerow(["Views Breakdown Report"])
            writer.writerow(["Generated on: {}".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))])
            writer.writerow(["Document: {}".format(doc.Title or "Unknown")])
            writer.writerow([])  # Empty row for spacing
            
            # Write column headers
            writer.writerow(["View Type", "On Sheets", "Not on Sheets", "Total"])
            
            # Write data rows
            for result in results_views:
                writer.writerow(result)

        print("CSV report saved to: {}".format(filepath))
        return filepath
        
    except Exception as e:
        print("Error saving CSV file: {}".format(str(e)))
        return None

if __name__ == '__main__':
    # Run the breakdown and get results
    results_views = Views_Breakdown()

    # Save results to CSV
    save_to_csv(results_views)
