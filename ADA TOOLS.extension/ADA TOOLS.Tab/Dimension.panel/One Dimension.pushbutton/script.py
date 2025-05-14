# -*- coding: utf-8 -*-
"""
Family Instance Dimensioning

Places one dimension from one detail line to another
"""
__title__ = "One-Dim\nDetail lines"
__author__ = "ADA"

# -*- coding: utf-8 -*-
import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
import Autodesk.Revit.DB as DB

from System.Collections.Generic import List
import System
from System.Windows.Forms import *

from pyrevit import revit, forms, script, HOST_APP

doc = HOST_APP.doc
uidoc = HOST_APP.uidoc
active_view = doc.ActiveView

# Define a class filter for detail lines
class DetailLineFilter(ISelectionFilter):
    def AllowElement(self, element):
        try:
            if element.Category.Id.IntegerValue == int(BuiltInCategory.OST_Lines):
                return True
            return False
        except:
            return False
        
    def AllowReference(self, ref, point):
        return True

# Function to display dimension type selection dialog
def select_dimension_type():
    # Get all dimension types
    dim_types = FilteredElementCollector(doc).OfClass(DimensionType).ToElements()
    
    # Format for display
    options = {Element.Name.GetValue(dim_type): dim_type for dim_type in dim_types}
    
    # Show selection dialog
    selected = forms.SelectFromList.show(
        options.keys(),
        title='Select Dimension Type',
        button_name='Select',
        multiselect=False
    )
    
    if selected:
        return options[selected]
    return None

# Main function to create dimensions
def create_dimensions_between_lines():
    try:
        # Get selected detail lines
        try:
            filter = DetailLineFilter()
            references = uidoc.Selection.PickObjects(ObjectType.Element, filter, "Select detail lines for dimensioning")
            
            if len(list(references)) < 2:
                forms.alert('Please select at least two detail lines.', exitscript=True)
        except Exception as e:
            script.exit(str(e))
        
        # Get all elements from the references
        lines = [doc.GetElement(ref) for ref in references]
        
        # Get dimension type from user
        dimension_type = select_dimension_type()
        
        if not dimension_type:
            script.exit()
        
        # Start transaction
        with revit.Transaction("Create Dimensions Between Lines"):
            # Create a reference array
            reference_array = ReferenceArray()
            for line in lines:
                ref = Reference(line)
                reference_array.Append(ref)
            
            # Get the first line's direction vector
            first_line = lines[0].GeometryCurve
            line_direction = first_line.Direction
            
            # Determine the "top" direction based on the view orientation
            # In plan views, top is usually the positive Y direction
            # In elevations or sections, top is usually the positive Z direction
            if active_view.ViewType == ViewType.FloorPlan or active_view.ViewType == ViewType.CeilingPlan:
                top_direction = XYZ(0, 0, 1)  # Z is up in plan views
            else:
                # For elevations, sections, etc.
                top_direction = active_view.UpDirection
            
            # Calculate a direction for the dimension offset
            # We want to offset in the direction that's most "upward" in the current view
            offset_direction = top_direction.CrossProduct(line_direction).Normalize()
            
            # If the resulting vector points "down", reverse it
            if offset_direction.DotProduct(top_direction) < 0:
                offset_direction = -offset_direction
            
            # Calculate midpoint of the first line
            midpoint = (first_line.GetEndPoint(0) + first_line.GetEndPoint(1)) / 2.0
            
            # Convert 10mm to feet (Revit's internal unit)
            # 1 foot = 304.8mm, so 10mm = 0.0328 feet
            offset_distance = 10.0 / 304.8  # 10mm in feet
            
            # Offset to create dimension line
            dimension_line_point = midpoint + offset_direction * offset_distance
            
            # Create dimension line
            dim_line = Line.CreateBound(
                dimension_line_point,
                dimension_line_point + line_direction * 100.0  # Making line longer than needed
            )
            
            # Create dimension
            dim = doc.Create.NewDimension(active_view, dim_line, reference_array)
            
            # Set dimension type
            dim.DimensionType = dimension_type
        
    except Exception as e:
        forms.alert(str(e), 'Error')

# Run the script
create_dimensions_between_lines()
