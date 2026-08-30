# -*- coding: utf-8 -*-
__doc__ = """Select one or more curves (in the current model or in a linked
model) and get their total length.

Pick curves directly rather than whole elements - what you see
highlighted while picking is exactly what gets summed, with no
ambiguity about which measurement method was used. This includes
edges of solid geometry (walls, pipes, framing...) as well as
free-standing Model Lines / Detail Lines / Reference Lines. Compare
with "Get Length", which instead picks whole elements and has to
guess at a method (location curve, parameter, bounding box, longest
edge) depending on what kind of element it is.

Revit's edge-picking (ObjectType.Edge) only works on the current
model - it does not let you click into a linked model at all. To
support links too, this script asks up front which one you're picking
from, and uses ObjectType.LinkedElement for the linked case instead.
A linked pick's Reference cannot be resolved back to a curve directly,
so instead the clicked point (Reference.GlobalPoint) is transformed
into the linked document's own coordinate system via the link's
placement transform, and matched against the linked element's own
curves/edges by closest projection - the same technique used by "Get
Surface" for linked faces.

Visualization: a red double-headed dimension-style arrow is drawn
along each picked curve's exact endpoints, offset toward the current
view so it stays readable regardless of the curve's own orientation,
with a red 3D digit readout of its length in meters placed next to it
- same technique as "Get Length". If several picked curves turn out
to sit on the same vector (same direction, same line - e.g. one edge
that Revit happens to have split into multiple collinear fragments),
they are merged into a single combined arrow spanning the group, with
one label showing the SUM of their lengths, instead of one arrow per
fragment.

Results are printed per curve and as a running total, in meters
and millimeters."""
__title__ = "Get Length\n(Curve)"
__version__ = "Version = 1.0"
__author__ = "ADA"

import math

from pyrevit import revit, DB, UI
from pyrevit import forms, script
from System.Collections.Generic import List

# Custom ADA GUI - small button-choice popup (same style as the picker
# used by Sections.panel/Openings.pushbutton) and the shared dark/gold
# themed report (see lib/GUI/ReportTheme.py)
from GUI.forms import select_from_buttons
from GUI.ReportTheme import ADAReport

# Get the active document
doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

MARKER_NAME = "ADA_QTO_CurveLengthArrowMarker"
TEXT_MARKER_NAME = "ADA_QTO_CurveLengthText"

ARROWHEAD_LENGTH = 0.35        # feet
ARROWHEAD_RADIUS = 0.13        # feet
ARROWHEAD_SIDES = 16
SHAFT_RADIUS = 0.045           # feet
SHAFT_SIDES = 12
TICK_LENGTH = 0.55             # feet, perpendicular tick mark at each end
TICK_THICKNESS = 0.05
TICK_DEPTH = 0.05
ARROW_OFFSET_MIN = 0.5         # feet, standoff toward the viewer off the measured line
ARROW_OFFSET_MAX = 3.0         # feet, cap for very long curves
ARROW_OFFSET_FRACTION = 0.12   # offset grows with curve length, up to the cap above
TEXT_STANDOFF = 0.6            # extra toward-viewer offset for the digit label, beyond the arrow

MARKER_COLOR = DB.Color(210, 30, 30)     # red
MARKER_LINE_COLOR = DB.Color(0, 0, 0)    # black edges
DIGIT_COLOR = DB.Color(210, 30, 30)      # red, matches the arrow
DIGIT_OFFSET = 0.05                      # feet, nudge digits toward the viewer

# --- 7-segment digit geometry (same technique as Get Length / Get Surface) -
DIGIT_W = 0.95
DIGIT_H = 1.75
STROKE = 0.24
DIGIT_GAP = 0.30
DOT_W = 0.42
DEPTH = 0.13

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


def get_solid_fill_pattern_id():
    for fp in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
        try:
            if fp.GetFillPattern().IsSolidFill:
                return fp.Id
        except Exception:
            continue
    return DB.ElementId.InvalidElementId


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


def face_reading_basis(normal):
    """Pick (u, v) = (right, up) perpendicular to `normal`, biased so
    `v` is as close to world-up as possible (projecting world Z onto
    the plane perpendicular to normal) - falls back to world Y if
    normal is itself near-vertical."""
    normal = normal.Normalize()
    world_ref = DB.XYZ(0, 0, 1) if abs(normal.Z) < 0.999 else DB.XYZ(0, 1, 0)
    v = world_ref.Subtract(normal.Multiply(world_ref.DotProduct(normal)))
    if v.GetLength() < 1e-6:
        world_ref = DB.XYZ(1, 0, 0)
        v = world_ref.Subtract(normal.Multiply(world_ref.DotProduct(normal)))
    v = v.Normalize()
    u = v.CrossProduct(normal).Normalize()
    return u, v


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


def build_cone_faces(tip, normal, length, radius, sides):
    """Build the face loops of a solid cone: apex at `tip`, base
    centered further out along `normal` by `length`."""
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


def build_cylinder_faces(p0, p1, radius, sides):
    """Build the face loops of a solid cylinder between p0 and p1."""
    axis = p1.Subtract(p0)
    length = axis.GetLength()
    if length < 1e-6:
        return []
    normal = axis.Normalize()

    arbitrary = DB.XYZ(0, 0, 1) if abs(normal.Z) < 0.9 else DB.XYZ(1, 0, 0)
    u = normal.CrossProduct(arbitrary).Normalize()
    v = normal.CrossProduct(u).Normalize()

    ring0, ring1 = [], []
    for i in range(sides):
        theta = 2.0 * math.pi * i / sides
        offset = u.Multiply(radius * math.cos(theta)).Add(
            v.Multiply(radius * math.sin(theta)))
        ring0.append(p0.Add(offset))
        ring1.append(p1.Add(offset))

    loops = [list(reversed(ring0)), list(ring1)]  # end caps
    for i in range(sides):
        j = (i + 1) % sides
        loops.append([ring0[i], ring1[i], ring1[j], ring0[j]])  # side quad

    return loops


def offset_arrow_points(p0, p1, view_normal):
    """Push the measured segment straight toward the viewer (along
    view_normal - the active view's ViewDirection, which already
    points toward the camera) so the drawn arrow pops out in front of
    the curve instead of sitting right on top of it - staying
    readable no matter how the curve itself is oriented. Shifting both
    endpoints by the same vector doesn't change the segment's
    direction/length, so this doesn't need to be perpendicular to it.
    The offset scales with the segment's length (clamped between
    ARROW_OFFSET_MIN and ARROW_OFFSET_MAX) so it stays visible on both
    short and long curves. Returns
    (p0_draw, p1_draw, direction, offset_dir), or all-None if p0/p1
    coincide."""
    axis = p1.Subtract(p0)
    dist = axis.GetLength()
    if dist < 1e-6:
        return None, None, None, None
    direction = axis.Normalize()
    offset_dir = view_normal.Normalize()
    offset = max(ARROW_OFFSET_MIN, min(ARROW_OFFSET_MAX, dist * ARROW_OFFSET_FRACTION))
    shift = offset_dir.Multiply(offset)
    return p0.Add(shift), p1.Add(shift), direction, offset_dir


def build_arrow_faces(p0, p1):
    """Build a double-headed dimension-style arrow from p0 to p1: a
    thin shaft with a solid arrowhead (cone) pointing at each
    endpoint, plus a short perpendicular tick mark at each end - the
    same convention as a Revit dimension line. p0/p1 are expected to
    already be the (offset) draw points."""
    axis = p1.Subtract(p0)
    dist = axis.GetLength()
    if dist < 1e-6:
        return []
    direction = axis.Normalize()
    head_len = min(ARROWHEAD_LENGTH, dist * 0.4)

    faces = []
    faces.extend(build_cone_faces(p0, direction, head_len, ARROWHEAD_RADIUS, ARROWHEAD_SIDES))
    faces.extend(build_cone_faces(p1, direction.Negate(), head_len, ARROWHEAD_RADIUS, ARROWHEAD_SIDES))

    shaft_p0 = p0.Add(direction.Multiply(head_len))
    shaft_p1 = p1.Subtract(direction.Multiply(head_len))
    if shaft_p1.Subtract(shaft_p0).DotProduct(direction) > 1e-6:
        faces.extend(build_cylinder_faces(shaft_p0, shaft_p1, SHAFT_RADIUS, SHAFT_SIDES))

    tick_u, tick_v = face_reading_basis(direction)
    for p in (p0, p1):
        origin = (p.Subtract(tick_u.Multiply(TICK_THICKNESS / 2.0))
                   .Subtract(tick_v.Multiply(TICK_LENGTH / 2.0))
                   .Subtract(direction.Multiply(TICK_DEPTH / 2.0)))
        faces.extend(box_faces(
            origin, tick_u, tick_v, direction,
            0.0, TICK_THICKNESS, 0.0, TICK_LENGTH, 0.0, TICK_DEPTH))

    return faces


def create_arrow_marker(p0, p1, view_normal):
    """p0/p1 are the true measured endpoints (host coords); the drawn
    arrow is offset toward the viewer from them via offset_arrow_points."""
    p0_draw, p1_draw, direction, _ = offset_arrow_points(p0, p1, view_normal)
    if direction is None:
        return None

    face_loops = build_arrow_faces(p0_draw, p1_draw)
    if not face_loops:
        return None

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
    ds.Name = MARKER_NAME

    ogs = DB.OverrideGraphicSettings()
    ogs.SetProjectionLineColor(MARKER_LINE_COLOR)
    ogs.SetProjectionLineWeight(3)
    fill_id = get_solid_fill_pattern_id()
    if fill_id != DB.ElementId.InvalidElementId:
        ogs.SetSurfaceForegroundPatternVisible(True)
        ogs.SetSurfaceForegroundPatternColor(MARKER_COLOR)
        ogs.SetSurfaceForegroundPatternId(fill_id)
    doc.ActiveView.SetElementOverrides(ds.Id, ogs)
    return ds


def create_length_digits(origin_point, facing_normal, length_m):
    u, v = face_reading_basis(facing_normal)

    text = "{:.2f}".format(length_m)
    total_w = sum((DOT_W if c == '.' else DIGIT_W) + DIGIT_GAP for c in text) - DIGIT_GAP
    origin = origin_point.Add(u.Multiply(-total_w / 2.0)).Add(facing_normal.Multiply(DIGIT_OFFSET))

    face_loops = build_number_faces(text, origin, u, v, facing_normal)

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


def collect_curves(geom_obj, curves):
    """Recursively collect Curve objects from a geometry object -
    standalone curves (model/detail lines), every edge of every solid,
    and anything nested inside a GeometryInstance."""
    if isinstance(geom_obj, DB.Curve):
        curves.append(geom_obj)
    elif isinstance(geom_obj, DB.Solid):
        if geom_obj.Edges.Size > 0:
            for edge in geom_obj.Edges:
                curves.append(edge.AsCurve())
    elif isinstance(geom_obj, DB.GeometryInstance):
        inst_geom = geom_obj.GetInstanceGeometry()
        if inst_geom:
            for g in inst_geom:
                collect_curves(g, curves)


def find_curve_at_point(element, point):
    """Find which curve of element's geometry the given point lies
    closest to, by projecting the point onto every curve and keeping
    the closest match. point must already be in the element's own
    document's coordinate system (i.e. transformed out of any link
    transform)."""
    options = DB.Options()
    options.DetailLevel = DB.ViewDetailLevel.Fine
    options.IncludeNonVisibleObjects = True

    geom = element.get_Geometry(options)
    if not geom:
        return None

    all_curves = []
    for g in geom:
        collect_curves(g, all_curves)

    best_curve = None
    best_dist = None
    for c in all_curves:
        try:
            result = c.Project(point)
        except Exception:
            result = None
        if result is not None:
            d = result.Distance
            if best_dist is None or d < best_dist:
                best_dist = d
                best_curve = c

    return best_curve


COLLINEAR_ANGLE_TOL = 0.01   # sine of the max angle between directions to still count as "same vector"
COLLINEAR_DIST_TOL = 0.05    # feet, max perpendicular distance between the two lines


def segments_collinear(p0a, p1a, p0b, p1b):
    """True if segment a and segment b lie on (near enough) the same
    infinite line - same direction (either way) and no perpendicular
    offset between them - regardless of any gap between them along
    that line."""
    da = p1a.Subtract(p0a)
    la = da.GetLength()
    db = p1b.Subtract(p0b)
    lb = db.GetLength()
    if la < 1e-9 or lb < 1e-9:
        return False
    da = da.Normalize()
    db = db.Normalize()

    if da.CrossProduct(db).GetLength() > COLLINEAR_ANGLE_TOL:
        return False

    w = p0b.Subtract(p0a)
    perp = w.Subtract(da.Multiply(w.DotProduct(da)))
    return perp.GetLength() <= COLLINEAR_DIST_TOL


def group_collinear_markers(markers):
    """Merge picked curves that sit on the same vector (same direction,
    same line) into a single combined marker: one arrow spanning the
    group's overall extent along that line, labeled with the SUM of
    the individual curve lengths - so an edge that Revit happens to
    have split into several collinear curves reads as one dimension
    instead of one per fragment. Returns a new marker list; each
    surviving group is (p0, p1, length_m, member_count)."""
    n = len(markers)
    parent = list(range(n))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i, j):
        ri, rj = find(i), find(j)
        if ri != rj:
            parent[ri] = rj

    for i in range(n):
        for j in range(i + 1, n):
            p0a, p1a, _ = markers[i]
            p0b, p1b, _ = markers[j]
            if segments_collinear(p0a, p1a, p0b, p1b):
                union(i, j)

    groups = {}
    for i in range(n):
        groups.setdefault(find(i), []).append(i)

    merged = []
    for members in groups.values():
        if len(members) == 1:
            p0, p1, length_m = markers[members[0]]
            merged.append((p0, p1, length_m, 1))
            continue

        p0_ref, p1_ref, _ = markers[members[0]]
        direction = p1_ref.Subtract(p0_ref).Normalize()

        total_len = 0.0
        min_t, max_t = None, None
        min_pt, max_pt = None, None
        for idx in members:
            p0, p1, length_m = markers[idx]
            total_len += length_m
            for pt in (p0, p1):
                t = pt.Subtract(p0_ref).DotProduct(direction)
                if min_t is None or t < min_t:
                    min_t, min_pt = t, pt
                if max_t is None or t > max_t:
                    max_t, max_pt = t, pt

        merged.append((min_pt, max_pt, total_len, len(members)))

    return merged


try:
    report = ADAReport()

    source = select_from_buttons(
        ["Current Model", "Linked Model"],
        title=__title__,
        label="Pick curve(s) in:",
        version=__version__
    )

    if not source:
        forms.alert("Cancelled.", exitscript=True)

    if source == "Current Model":
        refs = uidoc.Selection.PickObjects(
            UI.Selection.ObjectType.Edge,
            "Select one or more curves/edges to measure their length"
        )
    else:
        refs = uidoc.Selection.PickObjects(
            UI.Selection.ObjectType.LinkedElement,
            "Select one or more curves/edges in the linked model to measure their length"
        )

    if not refs:
        forms.alert("No curves selected.", exitscript=True)

    total_length = 0.0
    curve_details = []
    markers = []      # list of (p0_host, p1_host, length_m)
    table_rows = []   # rows for the "Measured Curves" table

    report.header(__title__.replace(chr(10), " "))

    for ref in refs:
        is_linked = ref.LinkedElementId != DB.ElementId.InvalidElementId
        curve = None

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
                # link's placement transform, then find which curve of
                # the linked element's geometry that point lies on.
                transform = link_instance.GetTotalTransform()
                local_point = transform.Inverse.OfPoint(ref.GlobalPoint)

                curve = find_curve_at_point(linked_element, local_point)
            else:
                element = doc.GetElement(ref.ElementId)
                display_id = ref.ElementId.IntegerValue
                location_note = "current model"
                transform = None

                geom_obj = element.GetGeometryObjectFromReference(ref)
                if isinstance(geom_obj, DB.Edge):
                    curve = geom_obj.AsCurve()
                elif isinstance(geom_obj, DB.Curve):
                    curve = geom_obj
        except Exception as resolve_err:
            report.warn("Element ID <b>{}</b>: could not resolve picked curve ({})".format(
                ref.LinkedElementId.IntegerValue if is_linked else ref.ElementId.IntegerValue,
                resolve_err))
            continue

        if curve is None:
            report.warn(
                "Element ID <b>{}</b> ({}): picked reference did not resolve to "
                "a curve, skipped".format(display_id, location_note))
            continue

        length = curve.Length
        total_length += length
        curve_details.append({
            'id': display_id,
            'length': length,
            'location': location_note
        })

        id_cell = str(display_id) if is_linked else output.linkify(ref.ElementId, title=str(display_id))
        table_rows.append([id_cell, "{:.2f} m".format(length * 0.3048)])

        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        if transform is not None:
            p0 = transform.OfPoint(p0)
            p1 = transform.OfPoint(p1)
        markers.append((p0, p1, length * 0.3048))

    if table_rows:
        report.subheader("Measured Curves")
        report.table(["Element ID", "Length"], table_rows)

    # Convert to display units
    total_length_mm = total_length * 304.8
    total_length_m = total_length * 0.3048

    report.subheader("Summary")
    report.line("Total curves measured: <b>{}</b>".format(len(curve_details)))
    report.line(
        "Total length: <b>{:.2f} m</b> ({:.0f} mm)".format(
            total_length_m, total_length_mm))

    # Curves that Revit happens to have split into several collinear
    # fragments (e.g. an edge broken up by extra vertices) would
    # otherwise draw one arrow per fragment; merge those into a single
    # combined arrow/label first, spanning the group and summing their
    # lengths.
    grouped_markers = group_collinear_markers(markers)
    merged_groups = [g for g in grouped_markers if g[3] > 1]
    if merged_groups:
        report.subheader("Merged Collinear Curves")
        for _, _, length_m, count in merged_groups:
            report.line(
                "{} curve(s) combined into one <b>{:.2f} m</b> dimension".format(count, length_m))

    # Draw a red dimension-style arrow along each measured curve (or
    # merged group of collinear curves), with a red 3D digit readout
    # of its length (meters) next to it.
    if grouped_markers:
        drawn = 0
        text_failed = 0
        try:
            with revit.Transaction("QTO Curve Length Marker"):
                clear_old_markers()
                try:
                    facing_normal = doc.ActiveView.ViewDirection.Normalize()
                except Exception:
                    facing_normal = DB.XYZ(0, 0, 1)

                for p0, p1, length_m, _count in grouped_markers:
                    marker = create_arrow_marker(p0, p1, facing_normal)
                    if marker is not None:
                        drawn += 1
                    try:
                        p0_draw, p1_draw, direction, offset_dir = offset_arrow_points(p0, p1, facing_normal)
                        if direction is not None:
                            midpoint = DB.XYZ(
                                (p0_draw.X + p1_draw.X) / 2.0,
                                (p0_draw.Y + p1_draw.Y) / 2.0,
                                (p0_draw.Z + p1_draw.Z) / 2.0)
                            text_origin = midpoint.Add(offset_dir.Multiply(TEXT_STANDOFF))
                            create_length_digits(text_origin, facing_normal, length_m)
                    except Exception as text_err:
                        text_failed += 1
                        report.warn("Could not build 3D length digits: {}".format(text_err))
            uidoc.RefreshActiveView()
            report.success("{} length arrow marker(s) drawn (red).".format(drawn))
            if text_failed:
                report.warn(
                    "{} of {} 3D length digit label(s) failed to create - see warnings above.".format(
                        text_failed, drawn))
        except Exception as marker_err:
            report.error("Could not draw length markers: {}".format(marker_err))

    report.flush()

except Exception as e:
    if 'cancel' not in str(e).lower():
        report = report if 'report' in globals() else ADAReport()
        report.error("Error: {}".format(e))
        report.flush()
        import traceback
        print(traceback.format_exc())
