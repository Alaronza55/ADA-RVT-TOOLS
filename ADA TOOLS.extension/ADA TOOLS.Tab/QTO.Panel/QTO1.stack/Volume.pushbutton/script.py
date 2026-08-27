__doc__ = """Select one or more elements (in the current model or in a linked
model) and get their total volume.

Measurement method, per element:
1. If the element has a built-in "Volume" parameter with a value
   greater than 0 (walls, floors, generic models... when "Volumes"
   is enabled under Area and Volume Computations), that value is
   used as-is.
2. Otherwise, the element's solid geometry is collected (recursing
   into nested geometry instances) and every solid's Volume is
   summed - i.e. the total volume of all solids that make up the
   element, not just the first/largest one.

The "Volume" parameter is only populated when volume computation is
turned on for the project; most elements will fall through to the
geometry method. See the hover diagram for a visual comparison.

Results are printed per element and as a running total, in cubic
feet, cubic meters and liters."""
__title__ = "Get Volume"
__author__ = "ADA"

from pyrevit import revit, DB, UI
from pyrevit import forms

# Get the active document
doc = revit.doc
uidoc = revit.uidoc

def get_solid_volume(solid_or_geom, volumes):
    """Recursively collect solid volumes from a geometry object"""
    if isinstance(solid_or_geom, DB.Solid):
        if solid_or_geom.Volume > 0:
            volumes.append(solid_or_geom.Volume)
    elif isinstance(solid_or_geom, DB.GeometryInstance):
        inst_geom = solid_or_geom.GetInstanceGeometry()
        if inst_geom:
            for inst_obj in inst_geom:
                get_solid_volume(inst_obj, volumes)

def get_element_volume(element):
    """Get volume from element intelligently"""

    # Method 1: Try the built-in "Volume" parameter first
    try:
        volume_param = element.LookupParameter("Volume")
        if volume_param and volume_param.HasValue:
            value = volume_param.AsDouble()
            if value > 0:
                return value
    except:
        pass

    # Method 2: Sum solid volumes from geometry
    try:
        options = DB.Options()
        options.ComputeReferences = True
        options.DetailLevel = DB.ViewDetailLevel.Fine
        options.IncludeNonVisibleObjects = True

        geom_element = element.get_Geometry(options)

        if geom_element:
            volumes = []
            for geom_obj in geom_element:
                get_solid_volume(geom_obj, volumes)

            if volumes:
                return sum(volumes)
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
                "Select elements to calculate their volume"
            )
            selected_ids = [ref.ElementId for ref in refs]

        for elem_id in selected_ids:
            element = doc.GetElement(elem_id)
            picked_elements.append((element, doc, elem_id.IntegerValue))

    else:  # Linked Model
        refs = uidoc.Selection.PickObjects(
            UI.Selection.ObjectType.LinkedElement,
            "Select elements in the linked model to calculate their volume"
        )

        for ref in refs:
            link_instance = doc.GetElement(ref.ElementId)
            linked_doc = link_instance.GetLinkDocument()
            linked_element = linked_doc.GetElement(ref.LinkedElementId)
            picked_elements.append(
                (linked_element, linked_doc, ref.LinkedElementId.IntegerValue))

    if not picked_elements:
        forms.alert("No elements selected.", exitscript=True)

    total_volume = 0.0
    elements_with_volume = 0
    elements_without_volume = 0
    element_details = []

    print("=" * 70)
    print("CALCULATING ELEMENT VOLUMES")
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

        # Get volume
        volume = get_element_volume(element)

        if volume and volume > 0:
            total_volume += volume
            elements_with_volume += 1
            element_details.append({
                'id': display_id,
                'name': element_name,
                'category': element_category,
                'volume': volume
            })
            print("\nElement ID {}: {:.3f} m3".format(
                display_id, volume * 0.0283168))
        else:
            elements_without_volume += 1
            print("\nElement ID {}: No volume found".format(display_id))

    # Convert to display units (cubic feet -> cubic meters / liters)
    total_volume_m3 = total_volume * 0.0283168
    total_volume_l = total_volume_m3 * 1000.0

    # Display summary
    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print("Total elements selected: {}".format(len(picked_elements)))
    print("Elements with volume: {}".format(elements_with_volume))
    print("Elements without volume: {}".format(elements_without_volume))
    print("-" * 70)
    print("TOTAL VOLUME: {:.3f} cubic feet".format(total_volume))
    print("TOTAL VOLUME: {:.3f} cubic meters".format(total_volume_m3))
    print("TOTAL VOLUME: {:.2f} liters".format(total_volume_l))
    print("=" * 70)

except Exception as e:
    print("Error: {}".format(e))
    import traceback
    traceback.print_exc()
