# -*- coding: utf-8 -*-
__doc__ = """Select one or more faces (in the current model or in a
linked model) and get their total area.

Pick faces directly rather than whole elements - what you see
highlighted while picking is exactly what gets summed, with no
ambiguity about which face/method was used.

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

After measuring, a yellow 3D arrow (cone) is drawn at each picked
face, pointing at it - similar to pyRevit's built-in 3D Measure tool
- so you can see exactly which faces were measured, even inside
links where Revit's own selection highlight can't show a single
face.
"""
__title__ = "Get Surface"
__author__ = "ADA"

import math

from pyrevit import revit, DB, UI
from pyrevit import forms
from System.Collections.Generic import List

doc = revit.doc
uidoc = revit.uidoc

MARKER_NAME = "ADA_QTO_FaceMarker"
MARKER_LENGTH = 1.6        # feet, cone height
MARKER_RADIUS = 0.4        # feet, cone base radius
MARKER_SIDES = 20          # facets around the cone
MARKER_COLOR = DB.Color(255, 205, 0)   # yellow, like the 3D Measure tool
MARKER_LINE_COLOR = DB.Color(0, 0, 0)  # black cone edges
MARKER_TRANSPARENCY = 25               # % transparent -> 75% opacity


def build_cone_faces(tip, normal, length=MARKER_LENGTH,
                      radius=MARKER_RADIUS, sides=MARKER_SIDES):
    """Build the face loops of a solid cone: apex at `tip`, base
    centered further out along `normal` by `length`. Returns a list
    of vertex loops (each an IList[XYZ]), suitable for
    TessellatedShapeBuilder - one loop for the base cap, one
    triangle per side facet."""
    normal = normal.Normalize()
    base_center = tip.Add(normal.Multiply(length))

    arbitrary = DB.XYZ(0, 0, 1) if abs(normal.Z) < 0.9 else DB.XYZ(1, 0, 0)
    u = normal.CrossProduct(arbitrary).Normalize()
    v = normal.CrossProduct(u).Normalize()

    ring = []
    for i in range(sides):
        theta = 2.0 * math.pi * i / sides
        offset = u.Multiply(radius * math.cos(theta)).Add(
            v.Multiply(radius * math.sin(theta)))
        ring.append(base_center.Add(offset))

    loops = [list(ring)]  # base cap
    for i in range(sides):
        p0 = ring[i]
        p1 = ring[(i + 1) % sides]
        loops.append([p0, p1, tip])

    return loops


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


def get_solid_fill_pattern_id():
    for fp in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
        try:
            if fp.GetFillPattern().IsSolidFill:
                return fp.Id
        except Exception:
            continue
    return DB.ElementId.InvalidElementId


def build_cone_geometry(tip, normal):
    """Build the cone as a Solid (falls back to a Mesh if the facet
    winding isn't clean enough for a strict BRep solid - either way
    TessellatedShapeBuilder returns something paintable)."""
    builder = DB.TessellatedShapeBuilder()
    builder.OpenConnectedFaceSet(False)
    for loop in build_cone_faces(tip, normal):
        builder.AddFace(DB.TessellatedFace(
            List[DB.XYZ](loop), DB.ElementId.InvalidElementId))
    builder.CloseConnectedFaceSet()
    builder.Target = DB.TessellatedShapeBuilderTarget.AnyGeometry
    builder.Fallback = DB.TessellatedShapeBuilderFallback.Mesh
    builder.Build()
    result = builder.GetBuildResult()
    return list(result.GetGeometricalObjects())


def create_marker(tip, normal):
    category_id = DB.ElementId(DB.BuiltInCategory.OST_GenericModel)
    ds = DB.DirectShape.CreateElement(doc, category_id)
    ds.SetShape(List[DB.GeometryObject](build_cone_geometry(tip, normal)))
    ds.Name = MARKER_NAME

    ogs = DB.OverrideGraphicSettings()
    ogs.SetProjectionLineColor(MARKER_LINE_COLOR)
    ogs.SetProjectionLineWeight(4)
    ogs.SetSurfaceTransparency(MARKER_TRANSPARENCY)
    fill_id = get_solid_fill_pattern_id()
    if fill_id != DB.ElementId.InvalidElementId:
        ogs.SetSurfaceForegroundPatternVisible(True)
        ogs.SetSurfaceForegroundPatternColor(MARKER_COLOR)
        ogs.SetSurfaceForegroundPatternId(fill_id)
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
    # built. So instead of relying on SetReferences, draw our own
    # yellow 3D arrow (cone) at each measured face's picked point,
    # pointing along its normal - like pyRevit's 3D Measure tool. Old
    # markers from a previous run are cleared first so repeated runs
    # don't clutter the model.
    if markers:
        try:
            with revit.Transaction("QTO Face Marker"):
                clear_old_markers()
                for tip, normal in markers:
                    create_marker(tip, normal)
            uidoc.RefreshActiveView()
            print("\n{} face marker(s) drawn in the view (yellow arrows).".format(
                len(markers)))
        except Exception as marker_err:
            print("\nCould not draw face markers: {}".format(marker_err))

except Exception as e:
    if 'cancel' not in str(e).lower():
        print("Error: {}".format(e))
        import traceback
        traceback.print_exc()
