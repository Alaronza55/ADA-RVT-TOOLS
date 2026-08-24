"""
Openings Absolute Level
Calculates the absolute bottom elevation (relative to Survey Point Z coordinate) 
for all generic models in the active view and sets it to OPE_ABSOLUTE LEVEL parameter
Uses actual geometry (not bounding box) for precise elevation calculation
Works with metric units (meters)
"""

__title__ = "Openings\nAbsolute Level"
__author__ = "ADA"

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    Transaction,
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
        collector = FilteredElementCollector(doc).OfClass(BasePoint)
        
        for bp in collector:
            if bp.IsShared:  # Survey Point is shared
                survey_point = bp.Position
                return survey_point.Z
        
        return 0.0
        
    except Exception as e:
        return 0.0


def get_bottom_elevation(element, survey_z):
    """
    Get the absolute bottom elevation of an element's geometry
    relative to Survey Point Z coordinate
    Uses actual geometry instead of bounding box
    """
    try:
        options = DB.Options()
        options.ComputeReferences = True
        options.DetailLevel = DB.ViewDetailLevel.Fine
        options.IncludeNonVisibleObjects = False
        
        geom_elem = element.get_Geometry(options)
        
        if geom_elem is None:
            return None
        
        min_z = None
        
        for geom_obj in geom_elem:
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
            return min_z - survey_z
        else:
            return None
            
    except Exception as e:
        return None


def process_geometry_object(geom_obj):
    """
    Process a geometry object and return its minimum Z coordinate
    """
    min_z = None
    
    try:
        if isinstance(geom_obj, DB.Solid):
            for face in geom_obj.Faces:
                mesh = face.Triangulate()
                for i in range(mesh.NumTriangles):
                    triangle = mesh.get_Triangle(i)
                    for j in range(3):
                        vertex = triangle.get_Vertex(j)
                        if min_z is None or vertex.Z < min_z:
                            min_z = vertex.Z
                            
        elif isinstance(geom_obj, DB.Mesh):
            for i in range(geom_obj.NumTriangles):
                triangle = geom_obj.get_Triangle(i)
                for j in range(3):
                    vertex = triangle.get_Vertex(j)
                    if min_z is None or vertex.Z < min_z:
                        min_z = vertex.Z
                        
        elif isinstance(geom_obj, DB.Curve):
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
        pass
    
    return min_z


def main():
    # Get active view
    active_view = doc.ActiveView
    
    # Collect all generic models in the active view
    collector = FilteredElementCollector(doc, active_view.Id) \
        .OfCategory(BuiltInCategory.OST_GenericModel) \
        .WhereElementIsNotElementType()
    
    elements = list(collector)
    
    if not elements:
        forms.alert("No generic models found in active view.", exitscript=True)
    
    # Get Survey Point Z coordinate
    survey_z = get_survey_point_position()
    
    # Process elements
    success_count = 0
    error_count = 0
    no_param_count = 0
    
    # Start transaction
    t = Transaction(doc, "Set Opening Absolute Levels")
    t.Start()
    
    try:
        for elem in elements:
            bottom_elev = get_bottom_elevation(elem, survey_z)
            
            if bottom_elev is None:
                error_count += 1
                continue
            
            param = elem.LookupParameter("OPE_ABSOLUTE LEVEL")
            
            if param is None:
                no_param_count += 1
                continue
            
            if not param.IsReadOnly:
                try:
                    param.Set(bottom_elev)
                    success_count += 1
                except Exception as e:
                    error_count += 1
            else:
                error_count += 1
        
        t.Commit()
        
    except Exception as e:
        t.RollBack()
        forms.alert("Error during transaction: {}".format(str(e)))
        sys.exit()
    
    # Report results
    message = "Process completed!\n\n"
    message += "{} elements updated successfully".format(success_count)
    
    if no_param_count > 0:
        message += "\n\nWarning: {} elements missing OPE_ABSOLUTE LEVEL parameter".format(no_param_count)
        forms.alert(message, warn_icon=True)
    else:
        forms.alert(message)


if __name__ == "__main__":
    main()