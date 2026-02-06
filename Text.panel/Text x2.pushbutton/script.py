# -*- coding: utf-8 -*-
__title__ = "Double Text Size"
__doc__ = "Select a text annotation to create a new type with 2x font size"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import revit, DB, forms

doc = revit.doc
uidoc = revit.uidoc

# Start transaction
t = Transaction(doc, "Create 2x Text Type")

try:
    # Prompt user to select a text annotation
    selection = uidoc.Selection
    
    # Select text element
    ref = selection.PickObject(ObjectType.Element, "Select a text annotation")
    text_elem = doc.GetElement(ref.ElementId)
    
    # Check if selected element is a text note
    if not isinstance(text_elem, TextNote):
        forms.alert("Selected element is not a text annotation. Please select a text note.", exitscript=True)
    
    # Get the text note type
    text_type_id = text_elem.GetTypeId()
    text_type = doc.GetElement(text_type_id)
    
    # Get current text size
    text_size_param = text_type.LookupParameter("Text Size")
    if text_size_param is None:
        forms.alert("Could not find Text Size parameter.", exitscript=True)
    
    current_size = text_size_param.AsDouble()
    new_size = current_size * 2
    
    # Start transaction
    t.Start()
    
    # Duplicate the text type
    original_name = text_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
    new_type_id = text_type.Duplicate("{}_2x".format(original_name))
    new_type = doc.GetElement(new_type_id)
    
    # Set the new text size
    new_text_size_param = new_type.LookupParameter("Text Size")
    new_text_size_param.Set(new_size)
    
    # Commit transaction
    t.Commit()
    
    # Convert sizes to display units (assumed to be in decimal feet, convert to inches or mm based on units)
    current_size_display = current_size * 12  # feet to inches for display
    new_size_display = new_size * 12
    
    forms.alert(
        "New text type created successfully!\n\n"
        "Original Type: {}\n"
        "Original Size: {:.2f}\"\n\n"
        "New Type: {}_2x\n"
        "New Size: {:.2f}\"".format(
            original_name, 
            current_size_display,
            original_name,
            new_size_display
        ),
        title="Success"
    )
    
except Exception as e:
    if t.HasStarted():
        t.RollBack()
    forms.alert("Error: {}".format(str(e)), title="Error")