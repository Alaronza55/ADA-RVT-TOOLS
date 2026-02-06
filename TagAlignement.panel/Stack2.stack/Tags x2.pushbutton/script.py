# -*- coding: utf-8 -*-
__title__ = "Double Label Type"
__doc__ = "Select label(s) in tag family to create new types with 2x font size"

from Autodesk.Revit.DB import *
from pyrevit import revit, forms

doc = revit.doc
uidoc = revit.uidoc

# Check if we're in a family document
if not doc.IsFamilyDocument:
    forms.alert("This script only works inside a tag family editor.", exitscript=True)

# Get currently selected elements
selection = uidoc.Selection
selected_ids = selection.GetElementIds()

if selected_ids.Count == 0:
    forms.alert("Please select one or more labels first, then run the script.", exitscript=True)

# Start transaction
t = Transaction(doc, "Create 2x Label Types")

try:
    t.Start()
    
    created_types = []
    
    # Process each selected element
    for elem_id in selected_ids:
        elem = doc.GetElement(elem_id)
        
        # Check if it's a TextElement (Label in family context)
        if isinstance(elem, TextElement):
            # Get the label's type
            label_symbol = doc.GetElement(elem.GetTypeId())
            
            # Get current text size
            text_size_param = label_symbol.LookupParameter("Text Size")
            if text_size_param is None:
                continue
            
            current_size = text_size_param.AsDouble()
            new_size = current_size * 2
            
            # Get the original type name to see what size it is
            original_name = label_symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            
            # Try to parse the original name as a number
            try:
                original_size_value = float(original_name)
                new_type_name = str(original_size_value * 2)
            except:
                # If original name is not a number, just use the internal size value
                # Convert from feet to mm and format
                new_size_mm = new_size * 304.8
                new_type_name = "{:.1f}".format(new_size_mm)
            
            # Check if type already exists
            existing_type = None
            collector = FilteredElementCollector(doc).OfClass(TextElementType)
            for existing in collector:
                existing_name = existing.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                if existing_name == new_type_name:
                    existing_type = existing
                    break
            
            if existing_type:
                created_types.append("Type '{}' already exists - skipped".format(new_type_name))
                continue
            
            # Duplicate the label type
            new_symbol = label_symbol.Duplicate(new_type_name)
            
            # Set the new text size
            new_text_size_param = new_symbol.LookupParameter("Text Size")
            new_text_size_param.Set(new_size)
            
            created_types.append("Created type '{}' from '{}'".format(new_type_name, original_name))
    
    # Commit transaction
    t.Commit()
    
    if created_types:
        forms.alert(
            "Label types processed:\n\n" + "\n".join(created_types),
            title="Success"
        )
    else:
        forms.alert("No label types were created. Make sure you selected labels.", title="Info")
    
except Exception as e:
    if t.HasStarted():
        t.RollBack()
    import traceback
    forms.alert("Error: {}\n\n{}".format(str(e), traceback.format_exc()), title="Error")