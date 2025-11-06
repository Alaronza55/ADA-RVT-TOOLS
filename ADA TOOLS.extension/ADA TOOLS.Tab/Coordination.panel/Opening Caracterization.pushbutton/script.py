"""
Openings Caracterization
Finds all generic models with "RESA" or "Opening" in type name
- Sets their absolute bottom elevation (relative to Survey Point Z coordinate) 
  to OPE_ABSOLUTE LEVEL parameter
- Assigns sequential numbers (1 to N) to OPE_NUMBER parameter based on distance 
  from origin (bigger to smaller), skipping elements that already have a number
Uses actual geometry (not bounding box) for precise elevation calculation
Works with metric units (meters)
"""

__title__ = "Openings\nCaracterization"
__author__ = "Your Name"

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    Transaction,
    BoundingBoxXYZ,
    BasePoint
)
from pyrevit import revit, DB, forms
import sys

# Get current document
doc = revit.doc
uidoc = revit.uidoc


def get_survey_point_position():
    """
    Get the Survey Point's Z coordinate position
    """
    try:
        # Collect all BasePoint elements
        collector = FilteredElementCollector(doc).OfClass(BasePoint)
        
        for bp in collector:
            # Check if it's the Survey Point (not Project Base Point)
            if bp.IsShared:  # Survey Point is shared
                # Get the actual position of the Survey Point
                survey_point = bp.Position
                return survey_point.Z
        
        print("Warning: Survey Point not found, using 0.0 as Z coordinate")
        return 0.0
        
    except Exception as e:
        print("Error getting Survey Point position: {}".format(str(e)))
        return 0.0


def get_bottom_elevation(element, survey_z):
    """
    Get the absolute bottom elevation of an element's geometry
    relative to Survey Point Z coordinate
    Uses actual geometry instead of bounding box
    """
    try:
        # Get geometry options
        options = DB.Options()
        options.ComputeReferences = True
        options.DetailLevel = DB.ViewDetailLevel.Fine
        options.IncludeNonVisibleObjects = False
        
        # Get the geometry element
        geom_elem = element.get_Geometry(options)
        
        if geom_elem is None:
            print("No geometry found for element ID: {}".format(element.Id))
            return None
        
        min_z = None
        
        # Iterate through geometry objects
        for geom_obj in geom_elem:
            # Handle geometry instances (for families)
            if isinstance(geom_obj, DB.GeometryInstance):
                inst_geom = geom_obj.GetInstanceGeometry()
                if inst_geom:
                    for inst_obj in inst_geom:
                        z_value = process_geometry_object(inst_obj)
                        if z_value is not None:
                            if min_z is None or z_value < min_z:
                                min_z = z_value
            else:
                z_value = process_geometry_object(geom_obj)
                if z_value is not None:
                    if min_z is None or z_value < min_z:
                        min_z = z_value
        
        if min_z is not None:
            # Calculate elevation relative to Survey Point Z coordinate
            return min_z - survey_z
        else:
            return None
            
    except Exception as e:
        print("Error getting geometry for element {}: {}".format(
            element.Id, str(e)))
        return None


def get_center_point(element):
    """
    Get the center point (X, Y) coordinates of an element
    Returns tuple (X, Y) or None if unable to calculate
    """
    try:
        # Try to get location point first (for point-based families)
        location = element.Location
        if isinstance(location, DB.LocationPoint):
            point = location.Point
            return (point.X, point.Y)
        
        # Otherwise, use bounding box center
        bbox = element.get_BoundingBox(None)
        if bbox is not None:
            center_x = (bbox.Min.X + bbox.Max.X) / 2.0
            center_y = (bbox.Min.Y + bbox.Max.Y) / 2.0
            return (center_x, center_y)
        
        return None
        
    except Exception as e:
        print("Error getting center point for element {}: {}".format(
            element.Id, str(e)))
        return None


def process_geometry_object(geom_obj):
    """
    Process a geometry object and return its minimum Z coordinate
    """
    min_z = None
    
    try:
        if isinstance(geom_obj, DB.Solid):
            # Get all vertices from faces
            for face in geom_obj.Faces:
                mesh = face.Triangulate()
                for i in range(mesh.NumTriangles):
                    triangle = mesh.get_Triangle(i)
                    for j in range(3):
                        vertex = triangle.get_Vertex(j)
                        if min_z is None or vertex.Z < min_z:
                            min_z = vertex.Z
                            
        elif isinstance(geom_obj, DB.Mesh):
            # Process mesh vertices
            for i in range(geom_obj.NumTriangles):
                triangle = geom_obj.get_Triangle(i)
                for j in range(3):
                    vertex = triangle.get_Vertex(j)
                    if min_z is None or vertex.Z < min_z:
                        min_z = vertex.Z
                        
        elif isinstance(geom_obj, DB.Curve):
            # Get curve endpoints and tessellation points
            if geom_obj.IsBound:
                start = geom_obj.GetEndPoint(0)
                end = geom_obj.GetEndPoint(1)
                if min_z is None or start.Z < min_z:
                    min_z = start.Z
                if min_z is None or end.Z < min_z:
                    min_z = end.Z
                    
        elif isinstance(geom_obj, DB.Point):
            coord = geom_obj.Coord
            if min_z is None or coord.Z < min_z:
                min_z = coord.Z
                
    except Exception as e:
        # Silently continue if there's an issue with a specific geometry object
        pass
    
    return min_z


def main():
    # Collect all generic models in the project
    collector = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_GenericModel) \
        .WhereElementIsNotElementType()
    
    # Filter elements by type name containing "RESA" or "Opening"
    filtered_elements = []
    for elem in collector:
        elem_type = doc.GetElement(elem.GetTypeId())
        if elem_type:
            type_name = elem_type.get_Parameter(
                DB.BuiltInParameter.SYMBOL_NAME_PARAM
            ).AsString()
            
            if type_name and ("RESA" in type_name.upper() or 
                            "OPENING" in type_name.upper()):
                filtered_elements.append(elem)
    
    if not filtered_elements:
        forms.alert("No generic models found with 'RESA' or 'Opening' "
                   "in their type name.", exitscript=True)
    
    print("Found {} elements to process".format(len(filtered_elements)))
    
    # Get Survey Point Z coordinate
    survey_z = get_survey_point_position()
    print("Survey Point Z coordinate: {} m".format(survey_z))
    
    # Calculate center points and create sorting data
    element_data = []
    for elem in filtered_elements:
        center = get_center_point(elem)
        if center is not None:
            x, y = center
            # Calculate distance from origin (0,0) for sorting
            # Bigger distance = further from origin
            distance = (x * x + y * y) ** 0.5
            element_data.append({
                'element': elem,
                'x': x,
                'y': y,
                'distance': distance
            })
        else:
            print("Could not get center point for element ID: {}".format(elem.Id))
    
    # Sort by distance from bigger to smaller
    element_data.sort(key=lambda item: item['distance'], reverse=True)
    
    print("Elements sorted by distance from origin (biggest to smallest)")
    
    # Process elements
    success_count = 0
    error_count = 0
    no_param_count = 0
    already_numbered_count = 0
    numbering_success_count = 0
    numbering_no_param_count = 0
    
    # Start transaction
    t = Transaction(doc, "Set Opening Absolute Levels and Numbers")
    t.Start()
    
    try:
        # Assign numbers sequentially
        for index, item in enumerate(element_data, start=1):
            elem = item['element']
            
            # Handle OPE_NUMBER parameter
            ope_number_param = elem.LookupParameter("OPE_NUMBER")
            if ope_number_param is not None:
                # Check if already has a number
                if ope_number_param.HasValue and ope_number_param.AsInteger() != 0:
                    existing_num = ope_number_param.AsInteger()
                    print("Element ID {} already has OPE_NUMBER: {} - skipping".format(
                        elem.Id, existing_num))
                    already_numbered_count += 1
                else:
                    # Assign new number
                    if not ope_number_param.IsReadOnly:
                        try:
                            ope_number_param.Set(index)
                            numbering_success_count += 1
                            print("Assigned OPE_NUMBER {} to element ID {} (X:{:.2f}, Y:{:.2f})".format(
                                index, elem.Id, item['x'], item['y']))
                        except Exception as e:
                            print("Error setting OPE_NUMBER for element {}: {}".format(
                                elem.Id, str(e)))
                    else:
                        print("OPE_NUMBER parameter is read-only for element ID: {}".format(elem.Id))
            else:
                print("Element ID {} does not have 'OPE_NUMBER' parameter".format(elem.Id))
                numbering_no_param_count += 1
            
            # Handle OPE_ABSOLUTE LEVEL parameter
            # Get bottom elevation relative to Survey Point Z coordinate
            bottom_elev = get_bottom_elevation(elem, survey_z)
            
            if bottom_elev is None:
                print("Could not get elevation for element ID: {}".format(
                    elem.Id))
                error_count += 1
                continue
            
            # Try to set the parameter
            param = elem.LookupParameter("OPE_ABSOLUTE LEVEL")
            
            if param is None:
                print("Element ID {} does not have 'OPE_ABSOLUTE LEVEL' "
                     "parameter".format(elem.Id))
                no_param_count += 1
                continue
            
            if not param.IsReadOnly:
                try:
                    param.Set(bottom_elev)
                    success_count += 1
                except Exception as e:
                    print("Error setting parameter for element {}: {}".format(
                        elem.Id, str(e)))
                    error_count += 1
            else:
                print("Parameter is read-only for element ID: {}".format(
                    elem.Id))
                error_count += 1
        
        t.Commit()
        
    except Exception as e:
        t.RollBack()
        forms.alert("Error during transaction: {}".format(str(e)))
        sys.exit()
    
    # Report results
    print("\n" + "="*50)
    print("RESULTS:")
    print("="*50)
    print("Total elements found: {}".format(len(filtered_elements)))
    print("\nOPE_ABSOLUTE LEVEL:")
    print("  Successfully updated: {}".format(success_count))
    print("  Missing parameter: {}".format(no_param_count))
    print("  Errors: {}".format(error_count))
    print("\nOPE_NUMBER:")
    print("  Successfully numbered: {}".format(numbering_success_count))
    print("  Already numbered (skipped): {}".format(already_numbered_count))
    print("  Missing parameter: {}".format(numbering_no_param_count))
    print("="*50)
    
    message = "Process completed!\n\n"
    message += "OPE_ABSOLUTE LEVEL: {} updated\n".format(success_count)
    message += "OPE_NUMBER: {} newly assigned, {} already numbered".format(
        numbering_success_count, already_numbered_count)
    
    if no_param_count > 0 or numbering_no_param_count > 0:
        message += "\n\nWarning: Some elements are missing parameters."
        message += "\nCheck the output window for details."
        forms.alert(message, warn_icon=True)
    else:
        forms.alert(message)


if __name__ == "__main__":
    main()