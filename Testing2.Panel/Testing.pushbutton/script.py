"""
Add Multiple Views to a Sheet
This script adds selected views to a selected sheet at specified positions.
"""
__title__ = 'Add Views\nto Sheet'
__author__ = 'AI Assistant'

from Autodesk.Revit.DB import (
    FilteredElementCollector, ViewSheet, Viewport, 
    XYZ, ElementId, Transaction, BuiltInCategory, 
    ViewType, View
)
from pyrevit import revit, forms, script
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

logger = script.get_logger()
output = script.get_output()

def get_sheets():
    """Get all sheets in the project"""
    sheets = FilteredElementCollector(doc)\
        .OfClass(ViewSheet)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    sheet_dict = {f"{sheet.SheetNumber} - {sheet.Name}": sheet for sheet in sheets}
    return sheet_dict

def get_placeable_views():
    """Get all views that can be placed on sheets"""
    all_views = FilteredElementCollector(doc)\
        .OfClass(View)\
        .WhereElementIsNotElementType()\
        .ToElements()
    
    valid_views = []
    for view in all_views:
        # Skip views already on sheets, sheets themselves, and non-placeable views
        if view.ViewType == ViewType.Schedule:
            valid_views.append(view)  # Schedules can be placed
        elif not view.IsTemplate and view.CanBePrinted:
            # Check if view is not already on a sheet
            if not Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id):
                continue
            valid_views.append(view)
    
    view_dict = {f"{view.ViewType.ToString()} - {view.Name}": view for view in valid_views}
    return view_dict

def add_views_to_sheet(sheet, views, start_point=(1.0, 1.0), spacing=1.0):
    """
    Add multiple views to a sheet with specified layout
    
    Args:
        sheet: The sheet to add views to
        views: List of views to add
        start_point: Starting point (x,y) in feet
        spacing: Space between views in feet
    """
    
    with Transaction(doc, "Add Views to Sheet") as t:
        t.Start()
        
        current_x, current_y = start_point
        max_width = 0
        
        try:
            for i, view in enumerate(views):
                # Create position for the viewport
                position = XYZ(current_x, current_y, 0)
                
                # Create the viewport
                viewport_id = Viewport.Create(
                    doc, 
                    sheet.Id, 
                    view.Id, 
                    position
                )
                
                viewport = doc.GetElement(viewport_id)
                
                # Get viewport dimensions to help with positioning
                bbox = viewport.get_BoundingBox(None)
                width = bbox.Max.X - bbox.Min.X
                height = bbox.Max.Y - bbox.Min.Y
                
                # Update max width
                max_width = max(max_width, width)
                
                # Move to next position (going down)
                current_y -= (height + spacing)
                
                # If we've placed 3 views vertically, start a new column
                if (i + 1) % 3 == 0:
                    current_x += (max_width + spacing)
                    current_y = start_point[1]
                    max_width = 0
            
            t.Commit()
            return True
        
        except Exception as e:
            t.RollBack()
            logger.error("Error adding views to sheet: {}".format(e))
            return False

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
        start_point = (x, y)
    except:
        start_point = (1.0, 1.0)
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
    
    # Add views to sheet
    result = add_views_to_sheet(sheet, selected_views, start_point, spacing)
    
    if result:
        output.print_md("# ✅ Success!")
        output.print_md(f"Added {len(selected_views)} views to sheet {sheet.SheetNumber} - {sheet.Name}")
    else:
        output.print_md("# ❌ Failed to add views to sheet")