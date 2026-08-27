__doc__ = """Draw Cross at Combined Center
Select multiple labels and draw one red cross at their combined center point"""
__title__ = "Draw Red Cross\nat Combined Center"
__author__ = "Alaronza"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import revit, DB, forms

doc = revit.doc
uidoc = revit.uidoc

# Check if we're in a family document
if not doc.IsFamilyDocument:
    forms.alert("This script must be run in a family document.", exitscript=True)

# Prompt user to select labels
selected_refs = None
try:
    selected_refs = uidoc.Selection.PickObjects(
        ObjectType.Element,
        "Select labels"
    )
except OperationCanceledException:
    # User cancelled selection
    import sys
    sys.exit()
except Exception as e:
    forms.alert("Error during selection: {}".format(str(e)), exitscript=True)

if not selected_refs or len(selected_refs) == 0:
    forms.alert("No elements selected.", exitscript=True)

# Get selected elements
selected_elements = [doc.GetElement(ref.ElementId) for ref in selected_refs]

# Function to get or create a red line style
def get_red_line_style(doc):
    """Find or create a red line style"""
    # Try to find existing red line style
    collector = FilteredElementCollector(doc).OfClass(GraphicsStyle)
    for gs in collector:
        if gs.GraphicsStyleCategory and gs.GraphicsStyleCategory.LineColor:
            color = gs.GraphicsStyleCategory.LineColor
            if color.Red == 255 and color.Green == 0 and color.Blue == 0:
                return gs
    
    # If not found, we'll just return None and skip color setting
    return None

# Collect all label points
label_points = []
active_view = doc.ActiveView

for element in selected_elements:
    label_point = None
    
    # Try to get location from LocationPoint
    location = element.Location
    if location and isinstance(location, LocationPoint):
        label_point = location.Point
    
    # If no LocationPoint, try to get bounding box center
    if not label_point:
        try:
            bbox = element.get_BoundingBox(active_view)
            if bbox:
                label_point = (bbox.Min + bbox.Max) / 2.0
        except:
            pass
    
    if label_point:
        label_points.append(label_point)

if not label_points:
    forms.alert("Could not find valid locations for the selected elements.", exitscript=True)

# Calculate the center point of all labels
sum_x = sum(pt.X for pt in label_points)
sum_y = sum(pt.Y for pt in label_points)
sum_z = sum(pt.Z for pt in label_points)
count = len(label_points)

center_point = XYZ(sum_x / count, sum_y / count, sum_z / count)

# Start transaction
t = Transaction(doc, "Draw Red Cross at Combined Center")
t.Start()

try:
    # Define cross size
    length = 0.05  # Small cross, adjust as needed
    
    # Get red line style if available
    red_style = get_red_line_style(doc)
    
    # Create horizontal line
    start_h = XYZ(center_point.X - length, center_point.Y, center_point.Z)
    end_h = XYZ(center_point.X + length, center_point.Y, center_point.Z)
    line_h = Line.CreateBound(start_h, end_h)
    
    # Create vertical line
    start_v = XYZ(center_point.X, center_point.Y - length, center_point.Z)
    end_v = XYZ(center_point.X, center_point.Y + length, center_point.Z)
    line_v = Line.CreateBound(start_v, end_v)
    
    # Create detail lines
    detail_line_h = doc.FamilyCreate.NewDetailCurve(active_view, line_h)
    detail_line_v = doc.FamilyCreate.NewDetailCurve(active_view, line_v)
    
    # Try to set line style if we found a red one
    if red_style:
        try:
            detail_line_h.LineStyle = red_style
            detail_line_v.LineStyle = red_style
        except:
            pass
    
    t.Commit()
    
    print("Created red cross at combined center of {} labels".format(count))
    if red_style:
        forms.alert("Successfully created red cross at the center of {} labels!".format(count))
    else:
        forms.alert("Successfully created cross at the center of {} labels! (Note: Could not find red line style)".format(count))
        
except Exception as e:
    t.RollBack()
    print("Error: {}".format(e))
    import traceback
    print(traceback.format_exc())
    forms.alert("Error: {}".format(str(e)))