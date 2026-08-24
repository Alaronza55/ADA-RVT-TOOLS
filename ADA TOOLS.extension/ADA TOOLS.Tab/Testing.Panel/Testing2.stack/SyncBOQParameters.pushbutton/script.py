# -*- coding: utf-8 -*-
__title__ = "Combine BOQ Parameters"
__author__ = "Alaronza"
__doc__ = """Combines Type Mark, Mark, Keynote, and ADBCode parameters 
into BESIX_BOQ_TypeMark_Mark_Keynote_ADB Code shared parameter 
for all elements visible in the active view."""

from Autodesk.Revit.DB import *
from pyrevit import revit, forms

doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView

# Parameter names to read (ADBCode is a shared parameter)
source_params = ["Type Mark", "Mark", "Keynote", "ADBCode"]
# Target shared parameter name
target_param = "BESIX_BOQ_TypeMark_Mark_Keynote_ADB Code"

def get_parameter_value(element, param_name):
    """Get parameter value as string from element or its type."""
    # Try instance parameter first
    param = element.LookupParameter(param_name)
    if param and param.HasValue:
        if param.StorageType == StorageType.String:
            return param.AsString() or ""
        elif param.StorageType == StorageType.Integer:
            return str(param.AsInteger())
        elif param.StorageType == StorageType.Double:
            return param.AsValueString() or str(param.AsDouble())
        else:
            return param.AsValueString() or ""
    
    # Try type parameter (especially for shared parameters)
    elem_type_id = element.GetTypeId()
    if elem_type_id != ElementId.InvalidElementId:
        elem_type = doc.GetElement(elem_type_id)
        if elem_type:
            param = elem_type.LookupParameter(param_name)
            if param and param.HasValue:
                if param.StorageType == StorageType.String:
                    return param.AsString() or ""
                elif param.StorageType == StorageType.Integer:
                    return str(param.AsInteger())
                elif param.StorageType == StorageType.Double:
                    return param.AsValueString() or str(param.AsDouble())
                else:
                    return param.AsValueString() or ""
    
    return ""

def set_parameter_value(element, param_name, value):
    """Set parameter value on element (shared parameter)."""
    param = element.LookupParameter(param_name)
    if param and not param.IsReadOnly:
        if param.StorageType == StorageType.String:
            param.Set(value)
            return True
    return False

# Collect all elements visible in active view
collector = FilteredElementCollector(doc, active_view.Id)\
    .WhereElementIsNotElementType()\
    .ToElements()

# Filter out view-specific elements and get elements with the target parameter
elements_to_process = []
for elem in collector:
    # Skip view-specific elements
    if isinstance(elem, (View, Viewport, Grid, Level, ReferencePlane)):
        continue
    
    # Check if element has the target shared parameter
    target = elem.LookupParameter(target_param)
    if target and target.StorageType == StorageType.String:
        elements_to_process.append(elem)

if not elements_to_process:
    forms.alert("No elements found in the active view with the shared parameter '{}'.\n\n"
                "Make sure the shared parameter is added to the elements in this view.".format(target_param),
                exitscript=True)

# Show confirmation
result = forms.alert(
    "Found {} elements in the active view.\n\n"
    "This will combine:\n"
    "• Type Mark (built-in)\n"
    "• Mark (built-in)\n"
    "• Keynote (built-in)\n"
    "• ADBCode (shared parameter)\n\n"
    "Into shared parameter: {}\n\n"
    "Separator: ' // '\n\n"
    "Continue?".format(len(elements_to_process), target_param),
    yes=True,
    no=True
)

if not result:
    import script
    script.exit()

# Process elements
modified_count = 0
skipped_count = 0
error_count = 0

t = Transaction(doc, "Combine BOQ Parameters")
t.Start()

try:
    for elem in elements_to_process:
        try:
            # Collect values from source parameters
            values = []
            for param_name in source_params:
                value = get_parameter_value(elem, param_name)
                values.append(value if value else "")
            
            # Combine with separator
            combined_value = " // ".join(values)
            
            # Set the target shared parameter
            if set_parameter_value(elem, target_param, combined_value):
                modified_count += 1
            else:
                skipped_count += 1
        except Exception as elem_error:
            error_count += 1
            continue
    
    t.Commit()
    
    # Show results
    message = "Operation completed!\n\n" \
              "Modified: {} elements\n" \
              "Skipped: {} elements (read-only parameter)".format(modified_count, skipped_count)
    
    if error_count > 0:
        message += "\nErrors: {} elements".format(error_count)
    
    forms.alert(message, title="Success")

except Exception as e:
    t.RollBack()
    forms.alert("Error: {}".format(str(e)), title="Error")