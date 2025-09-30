# -*- coding: utf-8 -*-
"""Count Views and Placed Views
This script counts total views in the project and views placed on sheets.
"""

__title__ = "AUDIT"
__author__ = "Almog Davidson"
__doc__ = """General Audit Information about the current Revit project."""

from pyrevit import revit, DB, script, forms
import os

# Get the current document
doc = revit.doc

def project_info():
    print("=== PROJECT FILE INFORMATION ===|")

    # Document Title
    print("Document Title: {}|".format(doc.Title))

    # Check if file has been saved
    if doc.PathName:
        # Full path
        print("Full Path: {}|".format(doc.PathName))
        
        # File name only
        file_name = os.path.basename(doc.PathName)
        print("File Name: {}|".format(file_name))
        
        # Check if workshared
        if doc.IsWorkshared:
            print("Is Workshared: Yes|")
            try:
                # Get the central model path
                central_path = doc.GetWorksharingCentralModelPath()
                if central_path:
                    # Convert ModelPath to user visible path
                    central_path_string = central_path.CentralPath
                    if central_path_string:
                        central_file_name = os.path.basename(central_path_string)
                        print("Central Model File Name: {}|".format(central_file_name))
                    else:
                        print("Central Model Path: {}|".format(str(central_path)))
            except:
                print("Could not retrieve central model path|")
        else:
            print("Is Workshared: No|")
            
    else:
        print("File has not been saved yet - using document title: {}|".format(doc.Title))

def Views():
        print("=== Views Audit of Project ===|")
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
        
        # Main summary
        # output.print_md("## View Summary")
        
        # Create a summary table
        print(" Total Views : **{0}** |".format(total_views))
        print(" Views on Sheets :  **{0}** |".format(placed_views_count))
        print(" Views NOT on Sheets :  **{0}** |".format(unplaced_views_count))
        print(" Total Sheets :  **{0}** |".format(len(all_sheets)))
        print("")
        
        # # Percentage calculation
        # if total_views > 0:
        #     placed_percentage = (placed_views_count * 100.0) / total_views
        #     print("|Percentage of views are placed on sheets : **{0:.1f}%**".format(placed_percentage))
        #     print("")

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
        print("=== View Count Summary ===|")

        print("| View Type | On Sheets | Not on Sheets |")
        
        for view_type in sorted(view_types_breakdown.keys()):
            total_count = view_types_breakdown[view_type]
            placed_count = placed_view_types_breakdown.get(view_type, 0)
            unplaced_count = total_count - placed_count
            
            print("| {0} | {1} - {2} |".format(
                view_type, 
                placed_count, 
                unplaced_count
            ))

def Revit_Categories():
    print("=== Revit Categories in Project ===|")
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
    print("Total Model Elements: {}".format(total_count))
    print("Categories Found: {}".format(len(category_counts)))
    print("-" * 50)

    for category, count in sorted_categories:
        print("{}: {}|".format(category, count))


if __name__ == '__main__':
    project_info()
    Views()
    Views_Breakdown()
    Revit_Categories()



