# -*- coding: utf-8 -*-
__doc__ = """Select one or more elements (in the current model or in a
linked model) and get their total volume.

Measurement method, per element (for the reported number):
1. If the element has a built-in "Volume" parameter with a value
   greater than 0 (walls, floors, generic models... when "Volumes"
   is enabled under Area and Volume Computations), that value is
   used as-is.
2. Otherwise, the element's solid geometry is collected (recursing
   into nested geometry instances) and every solid's Volume is
   summed - i.e. the total volume of all solids that make up the
   element, not just the first/largest one.

Visualization: regardless of which method produced the number, every
face of every solid making up the element is duplicated as a green,
80%-opacity generic model (via TessellatedShapeBuilder, from
Face.Triangulate()), each offset slightly outward along its own
face normal so the whole element appears wrapped in a green shell -
i.e. every face that contributed to "all faces of the solid(s)" is
shown, not just one. A black 3D digital-readout of the volume (cubic
meters, built from raised block segments) is placed flat on whichever
of the element's own faces is most visible from the current view
(i.e. faces the camera the most), offset just outside the green
shell - rather than floating at the element's centroid, which is
usually buried inside the solid and hard to read.

Results are printed per element and as a running total, in cubic
meters and liters.
"""
__title__ = "Get Volume"
__version__ = "Version 1.0"
__author__ = "ADA"

from pyrevit import revit, DB, UI
from pyrevit import forms
from System.Collections.Generic import List

# Shared ADA-Tools dark/gold themed report (see lib/GUI/ReportTheme.py) and
# small button-choice popup (see lib/GUI/SelectFromButtons.py)
from GUI.ReportTheme import ADAReport
from GUI.forms import select_from_buttons

# Get the active document
doc = revit.doc
uidoc = revit.uidoc

MARKER_NAME = "ADA_QTO_VolumeFaceMarker"
TEXT_MARKER_NAME = "ADA_QTO_VolumeText"
MARKER_OFFSET = 0.04                    # feet, how far each face is pushed outward
MARKER_COLOR = DB.Color(50, 160, 90)    # green
MARKER_LINE_COLOR = DB.Color(0, 0, 0)   # black edges
MARKER_TRANSPARENCY = 20                # % transparent -> 80% opacity
MIN_LABEL_FACE_AREA_RATIO = 0.05        # ignore sliver faces below this fraction of the largest face

# --- 7-segment digit geometry (same technique as Get Surface) -----------
DIGIT_W = 0.95
DIGIT_H = 1.75
STROKE = 0.24
DIGIT_GAP = 0.30
DOT_W = 0.42
DEPTH = 0.13
DIGIT_COLOR = DB.Color(0, 0, 0)        # black
DIGIT_OFFSET = 0.09                     # feet, standoff beyond the green shell so digits don't z-fight

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
    except Exception:
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
    except Exception:
        pass

    return None


def collect_solids(geom_obj, solids):
    """Recursively collect Solid objects (with volume > 0) from a
    geometry object"""
    if isinstance(geom_obj, DB.Solid):
        if geom_obj.Volume > 0:
            solids.append(geom_obj)
    elif isinstance(geom_obj, DB.GeometryInstance):
        inst_geom = geom_obj.GetInstanceGeometry()
        if inst_geom:
            for inst_obj in inst_geom:
                collect_solids(inst_obj, solids)


def get_element_solids(element):
    options = DB.Options()
    options.DetailLevel = DB.ViewDetailLevel.Fine
    options.IncludeNonVisibleObjects = True

    geom = element.get_Geometry(options)
    solids = []
    if geom:
        for g in geom:
            collect_solids(g, solids)
    return solids


def face_offset_triangles(face, transform, offset):
    """Triangulate a face and return its triangles, offset outward
    along the face's own normal, in host coordinates. transform is
    applied to every vertex/normal if given (linked elements); pass
    None for host elements."""
    mesh = face.Triangulate()
    if mesh is None:
        return []

    tris = []
    for i in range(mesh.NumTriangles):
        tri = mesh.get_Triangle(i)
        pts = [tri.get_Vertex(j) for j in range(3)]
        if transform is not None:
            pts = [transform.OfPoint(p) for p in pts]
        tris.append(pts)

    if not tris:
        return []

    try:
        bbox = face.GetBoundingBox()
        uv_mid = DB.UV((bbox.Min.U + bbox.Max.U) / 2.0, (bbox.Min.V + bbox.Max.V) / 2.0)
        normal = face.ComputeNormal(uv_mid)
        if transform is not None:
            normal = transform.OfVector(normal)
        normal = normal.Normalize()
        shift = normal.Multiply(offset)
        return [tuple(p.Add(shift) for p in tri) for tri in tris]
    except Exception:
        return [tuple(tri) for tri in tris]


def face_center_and_normal(face, transform):
    """Return (center_point, normal, area) for a face in host
    coordinates - the average of its triangulated vertices and its
    normal at the face's UV midpoint. Used to pick which of an
    element's faces is most visible from the current view, to place
    the volume label flat on it instead of at the element's (often
    buried) centroid."""
    mesh = face.Triangulate()
    if mesh is None or mesh.NumTriangles == 0:
        return None, None, 0.0

    pts = []
    for i in range(mesh.NumTriangles):
        tri = mesh.get_Triangle(i)
        for j in range(3):
            p = tri.get_Vertex(j)
            if transform is not None:
                p = transform.OfPoint(p)
            pts.append(p)

    if not pts:
        return None, None, 0.0

    center = DB.XYZ(
        sum(p.X for p in pts) / len(pts),
        sum(p.Y for p in pts) / len(pts),
        sum(p.Z for p in pts) / len(pts))

    try:
        bbox = face.GetBoundingBox()
        uv_mid = DB.UV((bbox.Min.U + bbox.Max.U) / 2.0, (bbox.Min.V + bbox.Max.V) / 2.0)
        normal = face.ComputeNormal(uv_mid)
        if transform is not None:
            normal = transform.OfVector(normal)
        normal = normal.Normalize()
    except Exception:
        return None, None, 0.0

    try:
        area = face.Area
    except Exception:
        area = 0.0

    return center, normal, area


def pick_visible_face(face_candidates, view_normal):
    """From a list of (center, normal, area), pick the one that best
    faces the current view (highest alignment with view_normal),
    ignoring sliver faces well below the largest face's area."""
    if not face_candidates:
        return None, None

    max_area = max(fc[2] for fc in face_candidates)
    significant = [fc for fc in face_candidates
                   if fc[2] >= max_area * MIN_LABEL_FACE_AREA_RATIO] or face_candidates

    center, normal, area = max(significant, key=lambda fc: fc[1].DotProduct(view_normal))
    return center, normal


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


def create_volume_marker(all_triangles):
    builder = DB.TessellatedShapeBuilder()
    builder.OpenConnectedFaceSet(False)
    for tri in all_triangles:
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
    """Pick (u, v) = (right, up) for text facing along `normal`,
    biased so `v` is as close to world-up as possible (projecting
    world Z onto the plane perpendicular to normal) - falls back to
    world Y if normal is itself near-vertical."""
    normal = normal.Normalize()
    world_ref = DB.XYZ(0, 0, 1) if abs(normal.Z) < 0.999 else DB.XYZ(0, 1, 0)
    v = world_ref.Subtract(normal.Multiply(world_ref.DotProduct(normal)))
    if v.GetLength() < 1e-6:
        world_ref = DB.XYZ(1, 0, 0)
        v = world_ref.Subtract(normal.Multiply(world_ref.DotProduct(normal)))
    v = v.Normalize()
    u = v.CrossProduct(normal).Normalize()
    return u, v


def create_volume_digits(label_point, label_normal, volume_m3):
    u, v = face_reading_basis(label_normal)

    text = "{:.2f}".format(volume_m3)
    total_w = sum((DOT_W if c == '.' else DIGIT_W) + DIGIT_GAP for c in text) - DIGIT_GAP
    origin = label_point.Add(u.Multiply(-total_w / 2.0)).Add(label_normal.Multiply(DIGIT_OFFSET))

    face_loops = build_number_faces(text, origin, u, v, label_normal)

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
    # Ask the user where to pick elements from
    source = select_from_buttons(
        ["Current Model", "Linked Model"],
        title=__title__,
        label="Select elements in:",
        version=__version__
    )

    if not source:
        forms.alert("Cancelled.", exitscript=True)

    # Each entry is (element, source_doc, display_id, link_transform)
    # link_transform is None for elements in the current (host) model
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
            picked_elements.append((element, doc, elem_id.IntegerValue, None))

    else:  # Linked Model
        refs = uidoc.Selection.PickObjects(
            UI.Selection.ObjectType.LinkedElement,
            "Select elements in the linked model to calculate their volume"
        )

        for ref in refs:
            link_instance = doc.GetElement(ref.ElementId)
            linked_doc = link_instance.GetLinkDocument()
            linked_element = linked_doc.GetElement(ref.LinkedElementId)
            transform = link_instance.GetTotalTransform()
            picked_elements.append(
                (linked_element, linked_doc, ref.LinkedElementId.IntegerValue, transform))

    if not picked_elements:
        forms.alert("No elements selected.", exitscript=True)

    report = ADAReport(__title__)

    # Computed once up front (view-only, no transaction needed) so each
    # element's label can be placed on whichever of its own faces is
    # most visible from here, rather than always facing the camera.
    try:
        facing_normal = doc.ActiveView.ViewDirection.Normalize()
    except Exception:
        facing_normal = DB.XYZ(0, 0, 1)

    total_volume = 0.0
    elements_with_volume = 0
    elements_without_volume = 0
    element_details = []
    markers = []      # list of (all_triangles, label_point, label_normal, volume_m3), host coords
    table_rows = []   # rows for the "Measured Elements" table
    not_found_ids = []

    # Process each selected element
    for element, element_doc, display_id, transform in picked_elements:
        # Get element info
        element_name = "Unnamed"
        try:
            element_name = element.Name
        except Exception:
            pass

        element_category = "No Category"
        try:
            if element.Category:
                element_category = element.Category.Name
        except Exception:
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
            table_rows.append([str(display_id), element_name, element_category,
                                "{:.3f} m3".format(volume * 0.0283168)])

            # Build the visualization from the element's actual solids,
            # independent of whether the number above came from the
            # Volume parameter or the geometry sum.
            try:
                solids = get_element_solids(element)
                all_triangles = []
                face_candidates = []  # (center, normal, area) per face, host coords
                for solid in solids:
                    if solid.Faces.Size == 0:
                        continue
                    for face in solid.Faces:
                        all_triangles.extend(
                            face_offset_triangles(face, transform, MARKER_OFFSET))
                        center, normal, area = face_center_and_normal(face, transform)
                        if center is not None and normal is not None:
                            face_candidates.append((center, normal, area))

                if all_triangles:
                    # Prefer placing the label flat on whichever real face
                    # is most visible from the current view; fall back to
                    # the overall centroid (old behavior) only if no face
                    # info could be resolved.
                    label_point, label_normal = pick_visible_face(face_candidates, facing_normal)
                    if label_point is None:
                        all_pts = [p for tri in all_triangles for p in tri]
                        label_point = DB.XYZ(
                            sum(p.X for p in all_pts) / len(all_pts),
                            sum(p.Y for p in all_pts) / len(all_pts),
                            sum(p.Z for p in all_pts) / len(all_pts))
                        label_normal = facing_normal
                    markers.append((all_triangles, label_point, label_normal, volume * 0.0283168))
                else:
                    report.warn("Element ID <b>{}</b>: no solid geometry found to visualize".format(display_id))
            except Exception as viz_err:
                report.warn("Element ID <b>{}</b>: could not build volume visualization ({})".format(
                    display_id, viz_err))
        else:
            elements_without_volume += 1
            not_found_ids.append(display_id)

    if table_rows:
        report.subheader("Measured Elements")
        report.table(["Element ID", "Name", "Category", "Volume"], table_rows)

    if not_found_ids:
        report.warn("No volume found for element ID(s): {}".format(
            ", ".join(str(i) for i in not_found_ids)))

    # Convert to display units (cubic feet -> cubic meters / liters)
    total_volume_m3 = total_volume * 0.0283168
    total_volume_l = total_volume_m3 * 1000.0

    report.subheader("Summary")
    report.line("Total elements selected: <b>{}</b>".format(len(picked_elements)))
    report.line("Elements with volume: <b>{}</b> / without: <b>{}</b>".format(
        elements_with_volume, elements_without_volume))
    report.line("Total volume: <b>{:.3f} m3</b> ({:.2f} L)".format(
        total_volume_m3, total_volume_l))

    # Wrap every solid face in a green, 80%-opacity shell, and place a
    # black 3D volume readout per element, flat on whichever of its own
    # faces is most visible from the current view.
    if markers:
        drawn = 0
        text_failed = 0
        try:
            with revit.Transaction("QTO Volume Marker"):
                clear_old_markers()

                for all_triangles, label_point, label_normal, volume_m3 in markers:
                    create_volume_marker(all_triangles)
                    drawn += 1
                    try:
                        create_volume_digits(label_point, label_normal, volume_m3)
                    except Exception as text_err:
                        text_failed += 1
                        report.warn("Could not build 3D volume digits: {}".format(text_err))
            uidoc.RefreshActiveView()
            report.success("{} volume marker(s) drawn (green, 80% opacity).".format(drawn))
            if text_failed:
                report.warn(
                    "{} of {} 3D volume digit label(s) failed to create - see warnings above.".format(
                        text_failed, drawn))
        except Exception as marker_err:
            report.error("Could not draw volume markers: {}".format(marker_err))

    report.flush()

except Exception as e:
    report = report if 'report' in globals() else ADAReport(__title__)
    report.error("Error: {}".format(e))
    report.flush()
    import traceback
    print(traceback.format_exc())
