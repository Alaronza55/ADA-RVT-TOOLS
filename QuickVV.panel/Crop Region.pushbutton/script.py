# -*- coding: utf-8 -*-
"""
Hide Crop Regions
----------------
This script allows you to select views on a sheet and hides their crop regions.
"""

from pyrevit import revit, script, forms
from Autodesk.Revit import DB
from Autodesk.Revit.UI import Selection

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

# Function to hide crop region of a view
def hide_crop_region(view):
    try:
        # Check if view has a crop box
        if not view.CropBoxActive:
            return False, "View does not have an active crop box"
        
        # Check if crop region is visible and hide it if it is
        if view.CropBoxVisible:
            with revit.Transaction("Hide Crop Region"):
                view.CropBoxVisible = False
            return True, "Crop region hidden successfully"
        else:
            return False, "Crop region is already hidden"
    except Exception as ex:
        return False, str(ex)

# Main execution
def main():
    try:
        # Ask user to select views on the sheet
        selection = uidoc.Selection.PickObjects(Selection.ObjectType.Element, 
                                             "Select views on sheet to hide crop regions")
        
        if not selection:
            forms.alert("No views selected. Operation canceled.", title="Warning")
            return
        
        successful_views = 0
        failed_views = 0
        errors = []
        
        # Process each selected element
        for selected_element in selection:
            element = doc.GetElement(selected_element)
            
            # Check if element is a viewport (view on sheet)
            if isinstance(element, DB.Viewport):
                view_id = element.ViewId
                view = doc.GetElement(view_id)
                
                # Try to hide crop region
                success, message = hide_crop_region(view)
                
                if success:
                    successful_views += 1
                else:
                    failed_views += 1
                    errors.append("View: {} - Error: {}".format(view.Name, message))
            else:
                failed_views += 1
                errors.append("Selected element is not a viewport")
        
        # Report results
        output.print_md("# Results")
        output.print_md("**Successfully processed: {} views**".format(successful_views))
        if failed_views > 0:
            output.print_md("**Failed: {} views**".format(failed_views))
            output.print_md("### Errors:")
            for error in errors:
                output.print_md("- {}".format(error))
        
        # Show summary in a dialog
        forms.alert("Operation completed.\n{} views processed successfully.\n{} views failed.".format(
            successful_views, failed_views), title="Hide Crop Regions")
    
    except Exception as ex:
        forms.alert("An error occurred: {}".format(str(ex)), title="Error")

# Execute main function
if __name__ == '__main__':
    main()
