"""
TEST VARIANT of Get Surface - measures picked FACES, not elements.

Pick one or more faces directly, in the current model or inside a
linked model, and get their total area.

Unlike "Get Surface", which measures whole elements and has to guess
at a method (parameter vs geometry), this variant measures exactly
the face(s) you click on - what you see highlighted while picking is
exactly what gets summed, with no ambiguity and no separate
"show me what was measured" step required.

Revit's face-picking (ObjectType.Face) only works on the current
model - it does not let you click into a linked model at all. To
support links too, this script asks up front which one you're
picking from, and uses ObjectType.LinkedElement for the linked case
instead. A linked pick's Reference cannot be resolved back to a Face
directly (GetGeometryObjectFromReference only works within the
document that actually owns the reference), so instead the clicked
point (Reference.GlobalPoint) is transformed into the linked
document's own coordinate system via the link's placement transform,
and matched against the linked element's faces by closest projection.
"""
__title__ = "Get Surface\n(Test - Faces)"
__author__ = "ADA"

import math

from pyrevit import revit, DB, UI
from pyrevit import forms
from System.Collections.Generic import List

doc = revit.doc
uidoc = revit.uidoc

MARKER_NAME = "ADA_QTO_FaceMarker"
MARKER_LENGTH = 3.0  # feet
MARKER_COLOR = DB.Color(255, 140, 0)

def build_arrow_lines(tip, normal, length=MARKER_LENGTH):
    """Build a simple 3-line leader arrow: a shaft from `tip` outward
    along `normal`, plus a small V-shaped arrowhead at `tip`."""
    normal = normal.Normalize()
    tail = tip.Add(normal.Multiply(length))
    lines = [DB.Line.CreateBound(tail, tip)]

    arbitrary = DB.XYZ(0, 0, 1) if abs(normal.Z) < 0.9 else DB.XYZ(1, 0, 0)
    side = normal.CrossProduct(arbitrary).Normalize()

    head_len = length * 0.25
    angle = math.radians(25)
    for sign in (1.0, -1.0):
        head_dir = normal.Multiply(math.cos(angle)).Add(
            side.Multiply(math.sin(angle) * sign))
        head_dir = head_dir.Normalize()
        head_end = tip.Add(head_dir.Multiply(head_len))
        lines.append(DB.Line.CreateBound(tip, head_end))

    return lines

def clear_old_markers():
    old_ids = []
    for ds in DB.FilteredElementCollector(doc).OfClass(DB.DirectShape):
        try:
            if ds.Name == MARKER_NAME:
                old_ids.append(ds.Id)
        except Exception:
            pass
    if old_ids:
        doc.Delete(List[DB.ElementId](old_ids))

def create_marker(tip, normal):
    category_id = DB.ElementId(DB.BuiltInCategory.OST_Lines)
    ds = DB.DirectShape.CreateElement(doc, category_id)
    ds.SetShape(List[DB.GeometryObject](build_arrow_lines(tip, normal)))
    ds.Name = MARKER_NAME

    ogs = DB.OverrideGraphicSettings()
    ogs.SetProjectionLineColor(MARKER_COLOR)
    ogs.SetProjectionLineWeight(6)
    doc.ActiveView.SetElementOverrides(ds.Id, ogs)
    return ds

def collect_faces(geom_obj, faces):
    """Recursively collect Face objects from a geometry object"""
    if isinstance(geom_obj, DB.Solid):
        if geom_obj.Faces.Size > 0:
            for f in geom_obj.Faces:
                faces.append(f)
    elif isinstance(geom_obj, DB.GeometryInstance):
        inst_geom = geom_obj.GetInstanceGeometry()
        if inst_geom:
            for g in inst_geom:
                collect_faces(g, faces)

def find_face_at_point(element, point):
    """Find which face of element's geometry the given point lies on,
    by projecting the point onto every face and keeping the closest
    match. point must already be in the element's own document's
    coordinate system (i.e. transformed out of any link transform).
    """
    options = DB.Options()
    options.DetailLevel = DB.ViewDetailLevel.Fine
    options.IncludeNonVisibleObjects = True

    geom = element.get_Geometry(options)
    if not geom:
        return None

    all_faces = []
    for g in geom:
        collect_faces(g, all_faces)

    best_face = None
    best_dist = None
    for f in all_faces:
        try:
            result = f.Project(point)
        except Exception:
            result = None
        if result is not None:
            d = result.Distance
            if best_dist is None or d < best_dist:
                best_dist = d
                best_face = f

    return best_face

try:
    source = forms.CommandSwitchWindow.show(
        ["Current Model", "Linked Model"],
        message="Pick face(s) in:"
    )

    if not source:
        forms.alert("Cancelled.", exitscript=True)

    if source == "Current Model":
        refs = uidoc.Selection.PickObjects(
            UI.Selection.ObjectType.Face,
            "Select one or more faces to measure their area"
        )
    else:
        refs = uidoc.Selection.PickObjects(
            UI.Selection.ObjectType.LinkedElement,
            "Select one or more faces in the linked model to measure their area"
        )

    if not refs:
        forms.alert("No faces selected.", exitscript=True)

    total_area = 0.0
    face_details = []
    markers = []  # list of (tip_point, normal_vector), both in host coords

    print("=" * 70)
    print("CALCULATING FACE SURFACE AREAS")
    print("=" * 70)

    for ref in refs:
        is_linked = ref.LinkedElementId != DB.ElementId.InvalidElementId
        face = None

        try:
            if is_linked:
                link_instance = doc.GetElement(ref.ElementId)
                linked_doc = link_instance.GetLinkDocument()
                linked_element = linked_doc.GetElement(ref.LinkedElementId)
                display_id = ref.LinkedElementId.IntegerValue
                location_note = "in link: {}".format(link_instance.Name)

                # ref.GlobalPoint is the actual clicked point, already
                # in host/world coordinates. Transform it into the
                # linked document's own coordinate system using the
                # link's placement transform, then find which face of
                # the linked element's geometry that point lies on.
                transform = link_instance.GetTotalTransform()
                local_point = transform.Inverse.OfPoint(ref.GlobalPoint)

                face = find_face_at_point(linked_element, local_point)
            else:
                element = doc.GetElement(ref.ElementId)
                display_id = ref.ElementId.IntegerValue
                location_note = "current model"
                face = element.GetGeometryObjectFromReference(ref)
        except Exception as resolve_err:
            print("\nElement ID {}: could not resolve picked face ({})".format(
                ref.LinkedElementId.IntegerValue if is_linked else ref.ElementId.IntegerValue,
                resolve_err))
            continue

        if not isinstance(face, DB.Face):
            print("\nElement ID {} ({}): picked reference resolved to {}, "
                  "not a face, skipped".format(
                      display_id, location_note,
                      type(face).__name__ if face is not None else "None"))
            continue

        area = face.Area
        total_area += area
        face_details.append({
            'id': display_id,
            'area': area,
            'location': location_note
        })

        print("\nElement ID {} ({}): {:.3f} m2".format(
            display_id, location_note, area * 0.09290304))

        # Work out the marker point/normal in host (world) coordinates
        try:
            if is_linked:
                uv = face.Project(local_point).UVPoint
                normal_host = transform.OfVector(face.ComputeNormal(uv))
            else:
                uv = face.Project(ref.GlobalPoint).UVPoint
                normal_host = face.ComputeNormal(uv)
            markers.append((ref.GlobalPoint, normal_host))
        except Exception as marker_err:
            print("  (could not compute marker for this face: {})".format(marker_err))

    total_area_m2 = total_area * 0.09290304

    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print("Total faces measured: {}".format(len(face_details)))
    print("-" * 70)
    print("TOTAL SURFACE AREA: {:.3f} square feet".format(total_area))
    print("TOTAL SURFACE AREA: {:.3f} square meters".format(total_area_m2))
    print("=" * 70)

    # Revit's own selection highlight only renders at whole-element
    # granularity for linked content, no matter how the Reference is
    # built - confirmed by testing. So instead of relying on
    # SetReferences, draw our own orange leader-arrow marker at each
    # measured face's picked point, pointing along its normal. Old
    # markers from a previous run are cleared first so repeated tests
    # don't clutter the model.
    if markers:
        try:
            with revit.Transaction("QTO Face Marker"):
                clear_old_markers()
                for tip, normal in markers:
                    create_marker(tip, normal)
            uidoc.RefreshActiveView()
            print("\n{} face marker(s) drawn in the view (orange arrows).".format(
                len(markers)))
        except Exception as marker_err:
            print("\nCould not draw face markers: {}".format(marker_err))

except Exception as e:
    if 'cancel' not in str(e).lower():
        print("Error: {}".format(e))
        import traceback
        traceback.print_exc()
