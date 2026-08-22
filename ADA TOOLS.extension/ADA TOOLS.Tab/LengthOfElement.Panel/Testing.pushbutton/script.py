"""
Select Elements and Get Total Geometric Length
"""
__title__ = "Get Length"
__author__ = "Your Name"

from pyrevit import revit, DB, UI
from pyrevit import forms

# Get the active document
doc = revit.doc
uidoc = revit.uidoc

def get_element_length(element):
    """Get length from element geometry intelligently"""
    
    # Method 1: Try Location Curve first (most reliable for linear elements)
    try:
        location = element.Location
        if isinstance(location, DB.LocationCurve):
            curve = location.Curve
            return curve.Length
    except:
        pass
    
    # Method 2: Check for common length parameters
    try:
        # Try OPE_THICKNESS parameter (for your openings)
        thickness_param = element.LookupParameter("OPE_THICKNESS")
        if thickness_param and thickness_param.HasValue:
            return thickness_param.AsDouble()
        
        # Try standard Length parameter
        length_param = element.LookupParameter("Length")
        if length_param and length_param.HasValue:
            return length_param.AsDouble()
    except:
        pass
    
    # Method 3: Use bounding box to determine orientation and get length
    try:
        bbox = element.get_BoundingBox(None)
        if bbox:
            min_pt = bbox.Min
            max_pt = bbox.Max
            
            # Calculate dimensions
            width = abs(max_pt.X - min_pt.X)
            depth = abs(max_pt.Y - min_pt.Y)
            height = abs(max_pt.Z - min_pt.Z)
            
            # Determine which dimension is the "length" (longest dimension)
            dimensions = [
                ('X', width),
                ('Y', depth),
                ('Z', height)
            ]
            
            # Sort by value to find the longest dimension
            dimensions.sort(key=lambda x: x[1], reverse=True)
            longest_dimension = dimensions[0][1]
            
            # Determine if element is vertical or horizontal
            # If Z is clearly the longest, it's vertical
            if height > width * 1.5 and height > depth * 1.5:
                return height  # Vertical element
            # If X or Y is clearly the longest, it's horizontal
            elif width > height * 1.5 or depth > height * 1.5:
                return max(width, depth)  # Horizontal element
            else:
                # If dimensions are similar, return the longest
                return longest_dimension
    except:
        pass
    
    # Method 4: Try geometry curves as last resort
    try:
        options = DB.Options()
        options.ComputeReferences = True
        options.DetailLevel = DB.ViewDetailLevel.Fine
        options.IncludeNonVisibleObjects = True
        
        geom_element = element.get_Geometry(options)
        
        if geom_element:
            all_curves = []
            
            for geom_obj in geom_element:
                if isinstance(geom_obj, DB.Curve):
                    all_curves.append(geom_obj.Length)
                
                elif isinstance(geom_obj, DB.GeometryInstance):
                    inst_geom = geom_obj.GetInstanceGeometry()
                    if inst_geom:
                        for inst_obj in inst_geom:
                            if isinstance(inst_obj, DB.Curve):
                                all_curves.append(inst_obj.Length)
                
                elif isinstance(geom_obj, DB.Solid):
                    if geom_obj.Edges.Size > 0:
                        for edge in geom_obj.Edges:
                            curve = edge.AsCurve()
                            all_curves.append(curve.Length)
            
            if all_curves:
                return max(all_curves)
    except:
        pass
    
    return None

try:
    # Get selection
    selection = uidoc.Selection
    selected_ids = selection.GetElementIds()
    
    if not selected_ids or selected_ids.Count == 0:
        refs = uidoc.Selection.PickObjects(
            UI.Selection.ObjectType.Element,
            "Select elements to calculate their length"
        )
        selected_ids = [ref.ElementId for ref in refs]
    
    if not selected_ids:
        forms.alert("No elements selected.", exitscript=True)
    
    total_length = 0.0
    elements_with_length = 0
    elements_without_length = 0
    element_details = []
    
    print("=" * 70)
    print("CALCULATING ELEMENT LENGTHS")
    print("=" * 70)
    
    # Process each selected element
    for elem_id in selected_ids:
        element = doc.GetElement(elem_id)
        
        # Get element info
        element_name = "Unnamed"
        try:
            element_name = element.Name
        except:
            pass
        
        element_category = "No Category"
        try:
            if element.Category:
                element_category = element.Category.Name
        except:
            pass
        
        # Get length
        length = get_element_length(element)
        
        if length and length > 0:
            total_length += length
            elements_with_length += 1
            element_details.append({
                'id': elem_id.IntegerValue,
                'name': element_name,
                'category': element_category,
                'length': length
            })
            print("\nElement ID {}: {:.2f} m".format(
                elem_id.IntegerValue, length * 0.3048))
        else:
            elements_without_length += 1
            print("\nElement ID {}: No length found".format(elem_id.IntegerValue))
    
    # Convert to display units
    total_length_mm = total_length * 304.8
    total_length_m = total_length * 0.3048
    
    # Display summary
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print("Total elements selected: {}".format(len(selected_ids)))
    print("Elements with length: {}".format(elements_with_length))
    print("Elements without length: {}".format(elements_without_length))
    print("-" * 70)
    print("TOTAL LENGTH: {:.2f} feet".format(total_length))
    print("TOTAL LENGTH: {:.2f} meters".format(total_length_m))
    print("TOTAL LENGTH: {:.2f} mm".format(total_length_mm))
    print("=" * 70)

except Exception as e:
    print("Error: {}".format(e))
    import traceback
    traceback.print_exc()