# -*- coding: utf-8 -*-
"""
Add Multiple Views to a Sheet
This script adds selected views to a selected sheet at specified positions.
"""
__title__ = 'Add Views\nto Sheet'
__author__ = 'Assistant'

import clr
from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSheet, Viewport, 
    XYZ, ElementId, Transaction, BuiltInCategory, 
    ViewType, View
)
from pyrevit import revit, forms, script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

output = script.get_output()

def get_sheets():
    """Get all sheets in the project"""
    sheets = FilteredElementCollector(doc)\
        .OfClass(ViewSheet)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    sheet_dict = {"{0} - {1}".format(sheet.SheetNumber, sheet.Name): sheet for sheet in sheets}
    return sheet_dict

def get_placeable_views():
    """Get all views that can be placed on sheets"""
    all_views = FilteredElementCollector(doc)\
        .OfClass(View)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    # Filter out views that are already on sheets or cannot be placed
    valid_views = []
    for view in all_views:
        if not view.IsTemplate and view.CanBePrinted:
            # Skip view templates and system browser
            view_type = view.ViewType
            if view_type not in [ViewType.ProjectBrowser, ViewType.SystemBrowser]:
                # Use Name property instead of ViewName
                valid_views.append(view)
    
    view_dict = {"{0} - {1}".format(view.ViewType.ToString(), view.Name): view for view in valid_views}
    return view_dict

# Main execution
if __name__ == "__main__":
    # Get sheet to place views on
    sheet_dict = get_sheets()
    selected_sheet_name = forms.SelectFromList.show(
        sorted(sheet_dict.keys()),
        title="Select Sheet",
        multiselect=False
    )
    
    if not selected_sheet_name:
        script.exit()
    
    sheet = sheet_dict[selected_sheet_name]
    
    # Get views to place on sheet
    view_dict = get_placeable_views()
    
    if not view_dict:
        forms.alert("No placeable views found.")
        script.exit()
    
    selected_view_names = forms.SelectFromList.show(
        sorted(view_dict.keys()),
        title="Select Views to Place on Sheet",
        multiselect=True
    )
    
    if not selected_view_names:
        script.exit()
    
    selected_views = [view_dict[name] for name in selected_view_names]
    
    # Get starting point
    start_point_input = forms.ask_for_string(
        default="1.0, 1.0",
        prompt="Enter starting point (X, Y) in feet:",
        title="Starting Position"
    )
    
    try:
        x, y = map(float, start_point_input.split(','))
        start_point = XYZ(x, y, 0)
    except:
        start_point = XYZ(1.0, 1.0, 0)
        forms.alert("Using default starting point (1.0, 1.0)")
    
    # Get spacing
    spacing_input = forms.ask_for_string(
        default="0.5",
        prompt="Enter spacing between views (feet):",
        title="View Spacing"
    )
    
    try:
        spacing = float(spacing_input)
    except:
        spacing = 0.5
        forms.alert("Using default spacing of 0.5 feet")
    
    # Process the views
    t = Transaction(doc, "Add Views to Sheet")
    t.Start()
    
    current_x = start_point.X
    current_y = start_point.Y
    success_count = 0
    
    try:
        for i, view in enumerate(selected_views):
            # Skip views that can't be added to this sheet
            if not Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id):
                output.print_md("Warning: Skipping view: {0} (already on a sheet or can't be placed)".format(view.Name))
                continue
            
            # Position for the viewport
            position = XYZ(current_x, current_y, 0)
            
            # Create the viewport
            viewport_id = Viewport.Create(doc, sheet.Id, view.Id, position)
            
            # Count successful placements
            success_count += 1
            
            # Move down for next view
            current_y -= 1.0 + spacing
            
            # If we've placed 3 views vertically, start a new column
            if (i + 1) % 3 == 0:
                current_x += 3.0
                current_y = start_point.Y
        
        t.Commit()
        
        output.print_md("# Success!")
        output.print_md("Added {0} views to sheet {1} - {2}".format(
            success_count,
            sheet.SheetNumber, 
            sheet.Name
        ))
    
    except Exception as e:
        t.RollBack()
        output.print_md("# Failed to add views to sheet")
        output.print_md("Error: {0}".format(str(e)))