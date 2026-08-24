"""
Select Elements and Get Total Surface Area
"""
__title__ = "Get Surface"
__author__ = "ADA"

from pyrevit import revit, DB, UI
from pyrevit import forms

# Get the active document
doc = revit.doc
uidoc = revit.uidoc

def get_solid_face_area(solid_or_geom, areas):
    """Recursively collect face areas from a geometry object"""
    if isinstance(solid_or_geom, DB.Solid):
        if solid_or_geom.Faces.Size > 0:
            for face in solid_or_geom.Faces:
                if face.Area > 0:
                    areas.append(face.Area)
    elif isinstance(solid_or_geom, DB.GeometryInstance):
        inst_geom = solid_or_geom.GetInstanceGeometry()
        if inst_geom:
            for inst_obj in inst_geom:
                get_solid_face_area(inst_obj, areas)

def get_element_surface_area(element):
    """Get surface area from element intelligently"""

    # Method 1: Try the built-in "Area" parameter first
    try:
        area_param = element.LookupParameter("Area")
        if area_param and area_param.HasValue:
            value = area_param.AsDouble()
            if value > 0:
                return value
    except:
        pass

    # Method 2: Sum all face areas from geometry (total surface area)
    try:
        options = DB.Options()
        options.ComputeReferences = True
        options.DetailLevel = DB.ViewDetailLevel.Fine
        options.IncludeNonVisibleObjects = True

        geom_element = element.get_Geometry(options)

        if geom_element:
            areas = []
            for geom_obj in geom_element:
                get_solid_face_area(geom_obj, areas)

            if areas:
                return sum(areas)
    except:
        pass

    return None

try:
    # Ask the user where to pick elements from
    source = forms.CommandSwitchWindow.show(
        ["Current Model", "Linked Model"],
        message="Select elements in:"
    )

    if not source:
        forms.alert("Cancelled.", exitscript=True)

    # Each entry is (element, source_doc, display_id)
    picked_elements = []

    if source == "Current Model":
        # Reuse existing selection if there is one, otherwise prompt
        selection = uidoc.Selection
        selected_ids = selection.GetElementIds()

        if not selected_ids or selected_ids.Count == 0:
            refs = uidoc.Selection.PickObjects(
                UI.Selection.ObjectType.Element,
                "Select elements to calculate their surface area"
            )
            selected_ids = [ref.ElementId for ref in refs]

        for elem_id in selected_ids:
            element = doc.GetElement(elem_id)
            picked_elements.append((element, doc, elem_id.IntegerValue))

    else:  # Linked Model
        refs = uidoc.Selection.PickObjects(
            UI.Selection.ObjectType.LinkedElement,
            "Select elements in the linked model to calculate their surface area"
        )

        for ref in refs:
            link_instance = doc.GetElement(ref.ElementId)
            linked_doc = link_instance.GetLinkDocument()
            linked_element = linked_doc.GetElement(ref.LinkedElementId)
            picked_elements.append(
                (linked_element, linked_doc, ref.LinkedElementId.IntegerValue))

    if not picked_elements:
        forms.alert("No elements selected.", exitscript=True)

    total_area = 0.0
    elements_with_area = 0
    elements_without_area = 0
    element_details = []

    print("=" * 70)
    print("CALCULATING ELEMENT SURFACE AREAS")
    print("=" * 70)

    # Process each selected element
    for element, element_doc, display_id in picked_elements:
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

        # Get surface area
        area = get_element_surface_area(element)

        if area and area > 0:
            total_area += area
            elements_with_area += 1
            element_details.append({
                'id': display_id,
                'name': element_name,
                'category': element_category,
                'area': area
            })
            print("\nElement ID {}: {:.3f} m2".format(
                display_id, area * 0.09290304))
        else:
            elements_without_area += 1
            print("\nElement ID {}: No surface area found".format(display_id))

    # Convert to display units (square feet -> square meters)
    total_area_m2 = total_area * 0.09290304

    # Display summary
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print("Total elements selected: {}".format(len(picked_elements)))
    print("Elements with surface area: {}".format(elements_with_area))
    print("Elements without surface area: {}".format(elements_without_area))
    print("-" * 70)
    print("TOTAL SURFACE AREA: {:.3f} square feet".format(total_area))
    print("TOTAL SURFACE AREA: {:.3f} square meters".format(total_area_m2))
    print("=" * 70)

except Exception as e:
    print("Error: {}".format(e))
    import traceback
    traceback.print_exc()
