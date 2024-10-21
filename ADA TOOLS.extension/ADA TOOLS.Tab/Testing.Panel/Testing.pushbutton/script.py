import clr
clr.AddReference("RevitServices")
clr.AddReference("RevitAPI")
clr.AddReference("RevitNodes")

from RevitServices.Persistence import DocumentManager
from Autodesk.Revit.DB import FilteredElementCollector, ElementIntersectsElementFilter

# Get the active document
doc = DocumentManager.Instance.CurrentDBDocument

# Define a function to check for intersection between two elements
def elements_intersect(element1_id, element2_id):
    # Get elements by their IDs
    element1 = doc.GetElement(element1_id)
    element2 = doc.GetElement(element2_id)
    
    # Create an ElementIntersectsElementFilter
    intersection_filter = ElementIntersectsElementFilter(element2)
    
    # Check if the first element intersects with the second element
    return intersection_filter.PassesFilter(element1)

# Example element IDs (replace with actual IDs from your Revit model)
element1_id = 12345  # Replace with the actual element ID
element2_id = 67890  # Replace with the actual element ID

# Call the function and print the result
if elements_intersect(element1_id, element2_id):
    print("The elements intersect.")
else:
    print("The elements do not intersect.")
