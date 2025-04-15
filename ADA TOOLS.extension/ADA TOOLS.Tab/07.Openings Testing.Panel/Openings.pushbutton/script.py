# -*- coding: utf-8 -*-
"""
Structural Element Selector
Allows selection of specific structural elements
"""
__title__ = 'Select\nStructural'
__author__ = 'Assistant'

import clr
from pyrevit import forms, script
from pyrevit import revit, DB, UI
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import ElementId, Wall, Floor, FamilyInstance, BuiltInCategory

# Create selection filter for structural elements
class StructuralElementsFilter(ISelectionFilter):
    def AllowElement(self, element):
        # Allow Walls
        if isinstance(element, Wall):
            return True
        
        # Allow Floors
        if isinstance(element, Floor):
            return True
        
        # Check for Structural Framing and Structural Columns
        if isinstance(element, FamilyInstance):
            category_id = element.Category.Id
            if (category_id == ElementId(BuiltInCategory.OST_StructuralFraming) or
                category_id == ElementId(BuiltInCategory.OST_StructuralColumns)):
                return True
        
        return False

    def AllowReference(self, reference, position):
        return False

# Main script
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

# Show dialog to user
result = forms.alert(
    msg='Please Select Structural Elements',
    sub_msg='You can select Walls, Floors, Structural Framing, and Structural Columns',
    ok=True,
    cancel=True,
    title='Structural Element Selection'
)

# If user clicked OK, continue to selection
if result:
    try:
        # Create the selection filter
        selection_filter = StructuralElementsFilter()
        
        # Prompt user to select elements
        selection = uidoc.Selection.PickObjects(
            ObjectType.Element, 
            selection_filter, 
            "Select structural elements (Walls, Floors, Structural Framing, or Structural Columns)"
        )
        
        # Get the selected elements
        selected_elements = [doc.GetElement(reference) for reference in selection]
        
        # Display results
        output.print_md("# Selected Structural Elements:")
        
        for element in selected_elements:
            category_name = element.Category.Name
            element_id = element.Id.IntegerValue
            
            if isinstance(element, Wall):
                element_type = "Wall"
            elif isinstance(element, Floor):
                element_type = "Floor"
            elif element.Category.Id == ElementId(BuiltInCategory.OST_StructuralFraming):
                element_type = "Structural Framing"
            elif element.Category.Id == ElementId(BuiltInCategory.OST_StructuralColumns):
                element_type = "Structural Column"
            else:
                element_type = "Other"
            
            output.print_md("- {}: {} (ID: {})".format(element_type, category_name, element_id))
            
        output.print_md("\n**Total Elements Selected:** {}".format(len(selected_elements)))
        
    except Exception as e:
        if 'Canceled' in str(e):
            output.print_md("**Selection was canceled by user**")
        else:
            output.print_md("**Error:** {}".format(str(e)))
else:
    output.print_md("**Operation canceled by user**")
