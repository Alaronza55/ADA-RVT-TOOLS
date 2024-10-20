# Import necessary Revit API modules and pyRevit modules
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import revit

# Get the active Revit document and UI document
uidoc = revit.uidoc
doc = revit.doc

# Prompt the user to select multiple elements
try:
    selection = uidoc.Selection.PickObjects(ObjectType.Element, "Select multiple elements and press 'Finish' when done.")
except Exception as e:
    print("Selection canceled.")
    selection = None

# Check if the user made a selection
if not selection:
    print("No elements selected.")
else:
    # Create a list to store element IDs
    element_ids = []

    # Loop through the selected elements to retrieve their IDs
    for sel in selection:
        # Get the element by its Reference
        element = doc.GetElement(sel.ElementId)
        if element:
            # Append the ElementId's IntegerValue (the actual ID number)
            element_ids.append(sel.ElementId.IntegerValue)

    # Print the selected element IDs
    print("Selected Element IDs:", element_ids)

