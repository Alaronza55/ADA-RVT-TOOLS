# -*- coding: utf-8 -*-
__title__ = "Views Breakdown"
__author__ = "Almog Davidson"
__doc__ = "Export all views with their type and sheet placement to CSV"

from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List
import csv
import os
import codecs

# Get the current document
doc = revit.doc
folder_name = doc.Title

def get_view_category(view):
    """Get the view category (FloorPlan, Section, CeilingPlan, etc.)"""
    try:
        return view.ViewType.ToString()
    except:
        return "Unknown"

def build_sheet_lookup_tables():
    """Build lookup dictionaries for fast sheet placement checking"""
    view_to_sheet = {}
    schedule_to_sheet = {}

    # Get all sheets once
    sheets = DB.FilteredElementCollector(doc)\
        .OfClass(DB.ViewSheet)\
        .ToElements()

    # Build viewport lookup (for regular views)
    for sheet in sheets:
        sheet_info = u"{} - {}".format(sheet.SheetNumber, sheet.Name)
        viewport_ids = sheet.GetAllViewports()
        for vp_id in viewport_ids:
            viewport = doc.GetElement(vp_id)
            view_to_sheet[viewport.ViewId] = sheet_info

    # Build schedule instance lookup (for schedules)
    schedule_instances = DB.FilteredElementCollector(doc)\
        .OfClass(DB.ScheduleSheetInstance)\
        .ToElements()

    for instance in schedule_instances:
        sheet = doc.GetElement(instance.OwnerViewId)
        if sheet:
            sheet_info = u"{} - {}".format(sheet.SheetNumber, sheet.Name)
            schedule_to_sheet[instance.ScheduleId] = sheet_info

    return view_to_sheet, schedule_to_sheet

def main():
    # Build lookup tables once at the beginning
    view_to_sheet, schedule_to_sheet = build_sheet_lookup_tables()

    # Collect all views (excluding view templates)
    views_collector = DB.FilteredElementCollector(doc)\
        .OfClass(DB.View)\
        .ToElements()

    # Collect all schedules separately
    schedules_collector = DB.FilteredElementCollector(doc)\
        .OfClass(DB.ViewSchedule)\
        .ToElements()

    # Filter out view templates, system views, and DrawingSheet views
    views = [v for v in views_collector 
             if not v.IsTemplate 
             and v.CanBePrinted
             and v.ViewType != DB.ViewType.DrawingSheet]  # Exclude sheets

    # Filter out schedule templates and revision schedules
    schedules = [s for s in schedules_collector 
                 if not s.IsTemplate 
                 and not s.Name.startswith("<Revision Schedule>")]

    # Combine views and schedules
    all_views = list(views) + list(schedules)

    if not all_views:
        forms.alert("No views found in the project.", exitscript=True)

    # Prepare data for CSV
    data = []
    data.append([u"View Name", u"View Category", u"Placed On Sheet"])

    # Counters for summary
    views_on_sheets = 0
    views_not_placed = 0

    # Dictionary to track view type distribution
    view_type_distribution = {}

    # Process each view using lookup tables
    for view in all_views:
        view_name = view.Name
        view_category = get_view_category(view)

        # Use lookup tables for fast sheet info retrieval
        if isinstance(view, DB.ViewSchedule):
            sheet_info = schedule_to_sheet.get(view.Id, u"NOT PLACED ON SHEET")
        else:
            sheet_info = view_to_sheet.get(view.Id, u"NOT PLACED ON SHEET")

        # Count for summary
        if sheet_info == u"NOT PLACED ON SHEET":
            views_not_placed += 1
            is_on_sheet = False
        else:
            views_on_sheets += 1
            is_on_sheet = True

        # Track view type distribution
        if view_category not in view_type_distribution:
            view_type_distribution[view_category] = {'on_sheets': 0, 'not_on_sheets': 0}
        
        if is_on_sheet:
            view_type_distribution[view_category]['on_sheets'] += 1
        else:
            view_type_distribution[view_category]['not_on_sheets'] += 1

        data.append([view_name, view_category, sheet_info])

    # Get the output window
    output = script.get_output()

    # Print results to PyRevit output window
    output.print_md("# Views Report")
    output.print_md("---")
    output.print_md("**Project:** {}".format(doc.Title))
    output.print_md("**Total Views:** {}".format(len(all_views)))
    output.print_md("**Views on Sheets:** {}".format(views_on_sheets))
    output.print_md("**Views NOT Placed:** {}".format(views_not_placed))
    output.print_md("---")
    output.print_md("")

    # Create a formatted table
    output.print_md("| View Name | View Category | Placed On Sheet |")
    output.print_md("|-----------|---------------|-----------------|")

    for row in data[1:]:  # Skip header row
        view_name = row[0].replace("|", "\\|")  # Escape pipe characters
        view_category = row[1].replace("|", "\\|")
        sheet_info = row[2].replace("|", "\\|")
        output.print_md(u"| {} | {} | {} |".format(view_name, view_category, sheet_info))

    output.print_md("")
    output.print_md("---")
    output.print_md("")

    # Print View Type Distribution
    output.print_md("## View Type Distribution")
    output.print_md("")
    output.print_md("| View Type | Sum of On Sheets | Sum of Not on Sheets |")
    output.print_md("|-----------|------------------|----------------------|")
    
    # Sort by view type name for consistent output
    for view_type in sorted(view_type_distribution.keys()):
        counts = view_type_distribution[view_type]
        output.print_md("| {} | {} | {} |".format(
            view_type, 
            counts['on_sheets'], 
            counts['not_on_sheets']
        ))
    
    output.print_md("")
    output.print_md("---")
    output.print_md("")

    # Define the output folder - change this path as needed
    output_folder = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker\{}".format(folder_name)

    # Create the folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Save detailed element breakdown
    filename_detailed = "Views_Breakdown_Detailed.csv"
    filepath_detailed = os.path.join(output_folder, filename_detailed)

    # Save view type distribution
    filename_distribution = "ViewType_Distribution.csv"
    filepath_distribution = os.path.join(output_folder, filename_distribution)

    # Write to CSV with UTF-8 encoding including BOM for Excel compatibility
    try:
        import codecs
        
        # Write detailed CSV
        with codecs.open(filepath_detailed, 'w', encoding='utf-8-sig') as csvfile:
            writer = csv.writer(csvfile)
            # Write header
            writer.writerow(['View Name', 'View Category', 'Placed On Sheet'])
            # Write data
            writer.writerows(data[1:])  # Skip header row since we're writing it separately

        # Write view type distribution CSV
        with codecs.open(filepath_distribution, 'w', encoding='utf-8-sig') as csvfile:
            writer = csv.writer(csvfile)
            # Write header
            writer.writerow(['View Type', 'On Sheets', 'Not on Sheets'])
            # Write distribution data - sorted by view type
            for view_type in sorted(view_type_distribution.keys()):
                counts = view_type_distribution[view_type]
                writer.writerow([view_type, counts['on_sheets'], counts['not_on_sheets']])

        # Print success message in PyRevit output
        output.print_md("## :white_heavy_check_mark: Export Successful!")
        output.print_md("")
        output.print_md("**Files saved to:**")
        output.print_md("```")
        output.print_md(filepath_detailed)
        output.print_md(filepath_distribution)
        output.print_md("```")
        output.print_md("")
        output.print_md("[Click here to open folder]({})".format(output_folder))

    except Exception as e:
        # Print error message in PyRevit output
        output.print_md("## :x: Export Failed!")
        output.print_md("")
        output.print_md("**Error:**")
        output.print_md("```")
        output.print_md(str(e))
        output.print_md("```")

# Run the script
if __name__ == '__main__':
    main()
