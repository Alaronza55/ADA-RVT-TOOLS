# -*- coding: utf-8 -*-
__doc__ = """TEST VARIANT of Get Surface - draws a duplicate plane and
a 3D digital-readout area label instead of an arrow marker.

Pick one or more faces directly, in the current model or inside a
linked model, and get their total area - same picking logic as
"Get Surface".

Instead of a pointer arrow, this variant recreates each picked
face's exact shape as a flat orange, 75%-opacity generic model
(via TessellatedShapeBuilder, from Face.Triangulate()), offset
slightly off the real face so it doesn't z-fight with it - the
marker's own area matches the measured face's area exactly, since
it IS that face's shape.

It also builds the area value (in square meters) as real 3D
geometry near the face's centroid: each digit is drawn as a
blocky, 7-segment/digital-display-style numeral out of small
raised boxes (same TessellatedShapeBuilder technique as the plane -
real Model Text turned out to require creating a whole temporary
family document under the hood, so this sidesteps that entirely).
"""
__title__ = "Get Surface\n(Test - Plane)"
__author__ = "ADA"

from pyrevit import revit, DB, UI
from pyrevit import forms
from System.Collections.Generic import List

doc = revit.doc
uidoc = revit.uidoc

# --- 7-segment digit geometry --------------------------------------------
# Each digit is drawn in a local 2D cell (x: 0..DIGIT_W, y: 0..DIGIT_H,
# both in feet) using up to 7 rectangular segments (A..G, standard
# 7-segment layout), each extruded a small depth along the face normal.
DIGIT_W = 0.95
DIGIT_H = 1.75
STROKE = 0.24
DIGIT_GAP = 0.30
DOT_W = 0.42
DEPTH = 0.13  # feet, how far the digits stick out past the plane marker
DIGIT_COLOR = DB.Color(30, 90, 220)  # blue

SEGMENT_RECTS = {
    'A': (STROKE * 0.5, DIGIT_H - STROKE, DIGIT_W - STROKE * 0.5, DIGIT_H),
    'G': (STROKE * 0.5, DIGIT_H / 2.0 - STROKE / 2.0, DIGIT_W - STROKE * 0.5, DIGIT_H / 2.0 + STROKE / 2.0),
    'D': (STROKE * 0.5, 0.0, DIGIT_W - STROKE * 0.5, STROKE),
    'F': (0.0, DIGIT_H / 2.0, STROKE, DIGIT_H - STROKE * 0.5),
    'B': (DIGIT_W - STROKE, DIGIT_H / 2.0, DIGIT_W, DIGIT_H - STROKE * 0.5),
    'E': (0.0, STROKE * 0.5, STROKE, DIGIT_H / 2.0),
    'C': (DIGIT_W - STROKE, STROKE * 0.5, DIGIT_W, DIGIT_H / 2.0),
}

DIGIT_SEGMENTS = {
    '0': 'ABCDEF', '1': 'BC', '2': 'ABGED', '3': 'ABGCD',
    '4': 'FGBC', '5': 'AFGCD', '6': 'AFGECD', '7': 'ABC',
    '8': 'ABCDEFG', '9': 'ABCDFG',
}

MARKER_NAME = "ADA_QTO_FacePlaneMarker"
TEXT_MARKER_NAME = "ADA_QTO_FaceAreaText"
MARKER_OFFSET = 0.03                    # feet, lift off the real face to avoid z-fighting
MARKER_COLOR = DB.Color(255, 140, 0)    # orange
MARKER_LINE_COLOR = DB.Color(0, 0, 0)   # black edges
MARKER_TRANSPARENCY = 25                # % transparent -> 75% opacity


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


def face_triangles_host(face, transform):
    """Triangulate a face and return its triangles as a list of
    (p0, p1, p2) tuples in host coordinates. transform is applied to
    every vertex if given (linked faces); pass None for host faces."""
    mesh = face.Triangulate()
    triangles = []
    for i in range(mesh.NumTriangles):
        tri = mesh.get_Triangle(i)
        pts = [tri.get_Vertex(j) for j in range(3)]
        if transform is not None:
            pts = [transform.OfPoint(p) for p in pts]
        triangles.append(tuple(pts))
    return triangles


def offset_triangles(triangles, normal, offset):
    shift = normal.Multiply(offset)
    return [tuple(p.Add(shift) for p in tri) for tri in triangles]


def clear_old_markers():
    old_ids = []
    for ds in DB.FilteredElementCollector(doc).OfClass(DB.DirectShape):
        try:
            if ds.Name in (MARKER_NAME, TEXT_MARKER_NAME):
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


def create_plane_marker(triangles):
    builder = DB.TessellatedShapeBuilder()
    builder.OpenConnectedFaceSet(False)
    for tri in triangles:
        builder.AddFace(DB.TessellatedFace(
            List[DB.XYZ](list(tri)), DB.ElementId.InvalidElementId))
    builder.CloseConnectedFaceSet()
    builder.Target = DB.TessellatedShapeBuilderTarget.AnyGeometry
    builder.Fallback = DB.TessellatedShapeBuilderFallback.Mesh
    builder.Build()
    result = builder.GetBuildResult()
    geom_objs = list(result.GetGeometricalObjects())

    category_id = DB.ElementId(DB.BuiltInCategory.OST_GenericModel)
    ds = DB.DirectShape.CreateElement(doc, category_id)
    ds.SetShape(List[DB.GeometryObject](geom_objs))
    ds.Name = MARKER_NAME

    ogs = DB.OverrideGraphicSettings()
    ogs.SetProjectionLineColor(MARKER_LINE_COLOR)
    ogs.SetProjectionLineWeight(3)
    ogs.SetSurfaceTransparency(MARKER_TRANSPARENCY)
    fill_id = get_solid_fill_pattern_id()
    if fill_id != DB.ElementId.InvalidElementId:
        ogs.SetSurfaceForegroundPatternVisible(True)
        ogs.SetSurfaceForegroundPatternColor(MARKER_COLOR)
        ogs.SetSurfaceForegroundPatternId(fill_id)
    doc.ActiveView.SetElementOverrides(ds.Id, ogs)
    return ds


def box_faces(origin, u, v, n, x0, x1, y0, y1, z0, z1):
    """Return the 6 quad faces of a box, in the local (u, v, n) frame
    rooted at `origin`: x along u, y along v, z along n."""
    def pt(x, y, z):
        return origin.Add(u.Multiply(x)).Add(v.Multiply(y)).Add(n.Multiply(z))

    p = {}
    for xi in (x0, x1):
        for yi in (y0, y1):
            for zi in (z0, z1):
                p[(xi, yi, zi)] = pt(xi, yi, zi)

    return [
        [p[(x0, y0, z0)], p[(x1, y0, z0)], p[(x1, y1, z0)], p[(x0, y1, z0)]],  # bottom (-n)
        [p[(x0, y0, z1)], p[(x0, y1, z1)], p[(x1, y1, z1)], p[(x1, y0, z1)]],  # top (+n)
        [p[(x0, y0, z0)], p[(x0, y1, z0)], p[(x0, y1, z1)], p[(x0, y0, z1)]],  # -u side
        [p[(x1, y0, z0)], p[(x1, y0, z1)], p[(x1, y1, z1)], p[(x1, y1, z0)]],  # +u side
        [p[(x0, y0, z0)], p[(x0, y0, z1)], p[(x1, y0, z1)], p[(x1, y0, z0)]],  # -v side
        [p[(x0, y1, z0)], p[(x1, y1, z0)], p[(x1, y1, z1)], p[(x0, y1, z1)]],  # +v side
    ]


def build_number_faces(text, origin, u, v, n):
    """Build face loops for `text` (digits and '.' only) as raised
    7-segment-style blocks, reading left to right along `u` starting
    at `origin`, sticking out along `n` by DEPTH."""
    faces = []
    cursor = 0.0
    for ch in text:
        if ch == '.':
            faces.extend(box_faces(
                origin, u, v, n,
                cursor, cursor + DOT_W, 0.0, STROKE,
                0.0, DEPTH))
            cursor += DOT_W + DIGIT_GAP
            continue

        segments = DIGIT_SEGMENTS.get(ch)
        if not segments:
            cursor += DIGIT_W + DIGIT_GAP
            continue

        for seg in segments:
            sx0, sy0, sx1, sy1 = SEGMENT_RECTS[seg]
            faces.extend(box_faces(
                origin, u, v, n,
                cursor + sx0, cursor + sx1, sy0, sy1,
                0.0, DEPTH))
        cursor += DIGIT_W + DIGIT_GAP

    return faces


def face_reading_basis(normal):
    """Pick (u, v) = (right, up) for text lying on a face with this
    normal, biased so `v` is as close to world-up as the face allows
    (projecting world Z onto the face plane) - so text on a wall
    reads right-side up instead of whatever direction an arbitrary
    cross product happens to land on. For near-horizontal faces
    (floors/ceilings, where world Z can't project onto the plane),
    world Y is used as the reference "up" instead.

    u is derived as v x normal, which keeps (u, v, normal) a
    right-handed frame - i.e. text reads correctly (not mirrored)
    when viewed from the same side the normal points to."""
    normal = normal.Normalize()
    world_ref = DB.XYZ(0, 0, 1) if abs(normal.Z) < 0.999 else DB.XYZ(0, 1, 0)
    v = world_ref.Subtract(normal.Multiply(world_ref.DotProduct(normal)))
    if v.GetLength() < 1e-6:
        world_ref = DB.XYZ(1, 0, 0)
        v = world_ref.Subtract(normal.Multiply(world_ref.DotProduct(normal)))
    v = v.Normalize()
    u = v.CrossProduct(normal).Normalize()
    return u, v


def create_area_digits(centroid, normal, area_m2):
    u, v = face_reading_basis(normal)

    text = "{:.2f}".format(area_m2)
    total_w = sum((DOT_W if c == '.' else DIGIT_W) + DIGIT_GAP for c in text) - DIGIT_GAP
    origin = centroid.Add(u.Multiply(-total_w / 2.0)).Add(normal.Multiply(MARKER_OFFSET))

    face_loops = build_number_faces(text, origin, u, v, normal)

    builder = DB.TessellatedShapeBuilder()
    builder.OpenConnectedFaceSet(False)
    for loop in face_loops:
        builder.AddFace(DB.TessellatedFace(
            List[DB.XYZ](loop), DB.ElementId.InvalidElementId))
    builder.CloseConnectedFaceSet()
    builder.Target = DB.TessellatedShapeBuilderTarget.AnyGeometry
    builder.Fallback = DB.TessellatedShapeBuilderFallback.Mesh
    builder.Build()
    result = builder.GetBuildResult()
    geom_objs = list(result.GetGeometricalObjects())

    category_id = DB.ElementId(DB.BuiltInCategory.OST_GenericModel)
    ds = DB.DirectShape.CreateElement(doc, category_id)
    ds.SetShape(List[DB.GeometryObject](geom_objs))
    ds.Name = TEXT_MARKER_NAME

    ogs = DB.OverrideGraphicSettings()
    ogs.SetProjectionLineColor(MARKER_LINE_COLOR)
    ogs.SetProjectionLineWeight(2)
    fill_id = get_solid_fill_pattern_id()
    if fill_id != DB.ElementId.InvalidElementId:
        ogs.SetSurfaceForegroundPatternVisible(True)
        ogs.SetSurfaceForegroundPatternColor(DIGIT_COLOR)
        ogs.SetSurfaceForegroundPatternId(fill_id)
    doc.ActiveView.SetElementOverrides(ds.Id, ogs)
    return ds


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
    markers = []  # list of (triangles_host, centroid_host, normal_host, area_ft2)

    print("=" * 70)
    print("CALCULATING FACE SURFACE AREAS")
    print("=" * 70)

    for ref in refs:
        is_linked = ref.LinkedElementId != DB.ElementId.InvalidElementId
        face = None
        transform = None

        try:
            if is_linked:
                link_instance = doc.GetElement(ref.ElementId)
                linked_doc = link_instance.GetLinkDocument()
                linked_element = linked_doc.GetElement(ref.LinkedElementId)
                display_id = ref.LinkedElementId.IntegerValue
                location_note = "in link: {}".format(link_instance.Name)

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

        try:
            triangles = face_triangles_host(face, transform)
            if not triangles:
                raise Exception("face triangulation returned no triangles")

            if is_linked:
                uv = face.Project(local_point).UVPoint
                normal_host = transform.OfVector(face.ComputeNormal(uv))
            else:
                uv = face.Project(ref.GlobalPoint).UVPoint
                normal_host = face.ComputeNormal(uv)
            normal_host = normal_host.Normalize()

            offset_tris = offset_triangles(triangles, normal_host, MARKER_OFFSET)

            all_pts = [p for tri in offset_tris for p in tri]
            centroid = DB.XYZ(
                sum(p.X for p in all_pts) / len(all_pts),
                sum(p.Y for p in all_pts) / len(all_pts),
                sum(p.Z for p in all_pts) / len(all_pts))

            markers.append((offset_tris, centroid, normal_host, area * 0.09290304))
        except Exception as marker_err:
            print("  (could not build plane marker for this face: {})".format(marker_err))

    total_area_m2 = total_area * 0.09290304

    print("\n" + "=" * 70)
    print("SUMMARY")
    print("=" * 70)
    print("Total faces measured: {}".format(len(face_details)))
    print("-" * 70)
    print("TOTAL SURFACE AREA: {:.3f} square feet".format(total_area))
    print("TOTAL SURFACE AREA: {:.3f} square meters".format(total_area_m2))
    print("=" * 70)

    if markers:
        drawn = 0
        text_failed = 0
        try:
            with revit.Transaction("QTO Face Plane Marker"):
                clear_old_markers()
                for triangles, centroid, normal, area_m2 in markers:
                    create_plane_marker(triangles)
                    drawn += 1
                    try:
                        create_area_digits(centroid, normal, area_m2)
                    except Exception as text_err:
                        text_failed += 1
                        print("  (could not build 3D area digits: {})".format(text_err))
            uidoc.RefreshActiveView()
            print("\n{} plane marker(s) drawn in the view (orange, 75% opacity).".format(drawn))
            if text_failed:
                print("{} of {} 3D area digit label(s) failed to create - see errors above.".format(
                    text_failed, drawn))
        except Exception as marker_err:
            print("\nCould not draw plane markers: {}".format(marker_err))

except Exception as e:
    if 'cancel' not in str(e).lower():
        print("Error: {}".format(e))
        import traceback
        traceback.print_exc()
