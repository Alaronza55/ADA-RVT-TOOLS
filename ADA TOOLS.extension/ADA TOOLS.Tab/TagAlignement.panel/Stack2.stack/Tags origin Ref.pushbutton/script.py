__doc__ = """Draw Cross at Origin
Creates two perpendicular annotation lines intersecting at 0,0,0 in a family document"""
__title__ = "Draw Cross\nat Origin"
__author__ = "Alaronza"

from Autodesk.Revit.DB import *
from pyrevit import revit, DB

doc = revit.doc

# Check if we're in a family document
if not doc.IsFamilyDocument:
    TaskDialog.Show("Error", "This script must be run in a family document.")
    raise Exception("Not a family document")

# Start transaction
t = Transaction(doc, "Draw Cross at Origin")
t.Start()

try:
    # Define the origin point
    origin = XYZ(0, 0, 0)
    
    # Define line length (adjust as needed)
    length = 1.0  # 1 foot
    
    # Create horizontal line (along X-axis)
    start_h = XYZ(-length, 0, 0)
    end_h = XYZ(length, 0, 0)
    line_h = Line.CreateBound(start_h, end_h)
    
    # Create vertical line (along Y-axis)
    start_v = XYZ(0, -length, 0)
    end_v = XYZ(0, length, 0)
    line_v = Line.CreateBound(start_v, end_v)
    
    # Get the active view
    active_view = doc.ActiveView
    
    # Create detail lines (annotation lines)
    detail_line_h = doc.FamilyCreate.NewDetailCurve(active_view, line_h)
    detail_line_v = doc.FamilyCreate.NewDetailCurve(active_view, line_v)
    
    t.Commit()
    print("Cross drawn successfully at origin (0,0,0)")
        
except Exception as e:
    t.RollBack()
    print("Error: {}".format(e))
    TaskDialog.Show("Error", str(e))