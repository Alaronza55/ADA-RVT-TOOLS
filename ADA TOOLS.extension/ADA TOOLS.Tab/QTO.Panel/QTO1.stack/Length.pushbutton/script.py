# -*- coding: utf-8 -*-
__doc__ = """Select one or more elements (in the current model or in a linked
model) and get their total length.

Measurement method, per element, tried in order until one succeeds:
1. Location Curve: if the element's Location is a curve (pipes,
   ducts, conduits, structural framing, walls...), that curve's
   actual length is used - the most reliable method, and correct
   even for sloped or curved elements.
2. Parameter: the "OPE_THICKNESS" parameter if present (for
   openings), otherwise a "Length" parameter if the element has one.
3. Bounding box: the element's bounding box is measured along X, Y
   and Z, and classified as vertical or horizontal based on how much
   one dimension dominates the other two; the corresponding dimension
   is used. This is an approximation - it can be wrong for elements
   that are diagonal or have no clearly dominant dimension.
4. Geometry edges: as a last resort, every curve/edge in the
   element's geometry (including nested geometry instances and solid
   edges) is collected and the SINGLE LONGEST one is used - not a
   sum, just the longest edge found.

Because the method varies per element, mixing element types in one
selection can mix measurement methods too. See the hover diagram for
a visual comparison of methods 1 and 3, the two most common.

Visualization: for each measured element, a red double-headed
dimension-style arrow is drawn along the exact segment that was
measured (the location curve's endpoints, the longest edge found, or
a segment along the dominant bounding-box axis), with a short
perpendicular tick mark at each end - and a red 3D digit readout of
that element's length in meters is placed next to it, facing the
current view. This makes it obvious which span each reported number
actually corresponds to.

Results are printed per element and as a running total, in meters
and millimeters."""
__title__ = "Get Length"
__version__ = "Version 1.0"
__author__ = "ADA"

import math

from pyrevit import revit, DB, UI
from pyrevit import forms, script
from System.Collections.Generic import List

# Shared ADA-Tools dark/gold themed report (see lib/GUI/ReportTheme.py) and
# small button-choice popup (see lib/GUI/SelectFromButtons.py)
from GUI.ReportTheme import ADAReport
from GUI.forms import select_from_buttons

# Get the active document
doc = revit.doc
uidoc = revit.uidoc

MARKER_NAME = "ADA_QTO_LengthArrowMarker"
TEXT_MARKER_NAME = "ADA_QTO_LengthText"

ARROWHEAD_LENGTH = 0.35        # feet
ARROWHEAD_RADIUS = 0.13        # feet
ARROWHEAD_SIDES = 16
SHAFT_RADIUS = 0.045           # feet
SHAFT_SIDES = 12
TICK_LENGTH = 0.55             # feet, perpendicular tick mark at each end
TICK_THICKNESS = 0.05
TICK_DEPTH = 0.05
ARROW_OFFSET_MIN = 0.5         # feet, standoff toward the viewer off the measured line
ARROW_OFFSET_MAX = 3.0         # feet, cap for very long elements
ARROW_OFFSET_FRACTION = 0.12   # offset grows with segment length, up to the cap above
TEXT_STANDOFF = 0.6            # extra toward-viewer offset for the digit label, beyond the arrow

MARKER_COLOR = DB.Color(210, 30, 30)     # red
MARKER_LINE_COLOR = DB.Color(0, 0, 0)    # black edges
DIGIT_COLOR = DB.Color(210, 30, 30)      # red, matches the arrow
DIGIT_OFFSET = 0.05                      # feet, nudge digits toward the viewer

# --- 7-segment digit geometry (same technique as Get Surface / Get Volume) -
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
    the element instead of sitting right on top of it - staying
    readable no matter how the element itself is oriented. Shifting
    both endpoints by the same vector doesn't change the segment's
    direction/length, so this doesn't need to be perpendicular to it.
    The offset scales with the segment's length (clamped between
    ARROW_OFFSET_MIN and ARROW_OFFSET_MAX) so it stays visible on both
    short and long elements. Returns
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


def bbox_axis_line(element, span_override=None):
    """Return (p0, p1, span) along the dominant bounding-box axis,
    centered at the bbox center - same axis-classification logic as
    the old bounding-box length method. If span_override is given,
    the segment uses that length instead of the raw bbox dimension
    along the chosen axis (used for the parameter-based method, which
    has a length value but no natural line of its own)."""
    try:
        bbox = element.get_BoundingBox(None)
        if not bbox:
            return None
        min_pt, max_pt = bbox.Min, bbox.Max
        width = abs(max_pt.X - min_pt.X)
        depth = abs(max_pt.Y - min_pt.Y)
        height = abs(max_pt.Z - min_pt.Z)
        center = DB.XYZ((min_pt.X + max_pt.X) / 2.0,
                         (min_pt.Y + max_pt.Y) / 2.0,
                         (min_pt.Z + max_pt.Z) / 2.0)

        if height > width * 1.5 and height > depth * 1.5:
            raw_span, axis_vec = height, DB.XYZ(0, 0, 1)
        elif width > height * 1.5 or depth > height * 1.5:
            if width >= depth:
                raw_span, axis_vec = width, DB.XYZ(1, 0, 0)
            else:
                raw_span, axis_vec = depth, DB.XYZ(0, 1, 0)
        else:
            dims = [(width, DB.XYZ(1, 0, 0)), (depth, DB.XYZ(0, 1, 0)), (height, DB.XYZ(0, 0, 1))]
            dims.sort(key=lambda d: d[0], reverse=True)
            raw_span, axis_vec = dims[0]

        span = span_override if span_override is not None else raw_span
        half = axis_vec.Multiply(span / 2.0)
        return center.Subtract(half), center.Add(half), span
    except Exception:
        return None


def get_element_length(element):
    """Get length from element geometry intelligently. Returns
    (length, p0, p1) - p0/p1 are the endpoints of the segment that
    was actually measured, in the element's own document's
    coordinates - or None if no length could be determined."""

    # Method 1: Try Location Curve first (most reliable for linear elements)
    try:
        location = element.Location
        if isinstance(location, DB.LocationCurve):
            curve = location.Curve
            return curve.Length, curve.GetEndPoint(0), curve.GetEndPoint(1)
    except Exception:
        pass

    # Method 2: Check for common length parameters
    try:
        thickness_param = element.LookupParameter("OPE_THICKNESS")
        length_param = thickness_param if (thickness_param and thickness_param.HasValue) \
            else element.LookupParameter("Length")
        if length_param and length_param.HasValue:
            value = length_param.AsDouble()
            line = bbox_axis_line(element, span_override=value)
            if line:
                p0, p1, _ = line
                return value, p0, p1
            return value, None, None
    except Exception:
        pass

    # Method 3: Use bounding box to determine orientation and get length
    try:
        line = bbox_axis_line(element)
        if line:
            p0, p1, span = line
            return span, p0, p1
    except Exception:
        pass

    # Method 4: Try geometry curves as last resort - the single longest one
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
                    all_curves.append(geom_obj)

                elif isinstance(geom_obj, DB.GeometryInstance):
                    inst_geom = geom_obj.GetInstanceGeometry()
                    if inst_geom:
                        for inst_obj in inst_geom:
                            if isinstance(inst_obj, DB.Curve):
                                all_curves.append(inst_obj)

                elif isinstance(geom_obj, DB.Solid):
                    if geom_obj.Edges.Size > 0:
                        for edge in geom_obj.Edges:
                            all_curves.append(edge.AsCurve())

            if all_curves:
                best = max(all_curves, key=lambda c: c.Length)
                return best.Length, best.GetEndPoint(0), best.GetEndPoint(1)
    except Exception:
        pass

    return None


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
                "Select elements to calculate their length"
            )
            selected_ids = [ref.ElementId for ref in refs]

        for elem_id in selected_ids:
            element = doc.GetElement(elem_id)
            picked_elements.append((element, doc, elem_id.IntegerValue, None))

    else:  # Linked Model
        refs = uidoc.Selection.PickObjects(
            UI.Selection.ObjectType.LinkedElement,
            "Select elements in the linked model to calculate their length"
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

    total_length = 0.0
    elements_with_length = 0
    elements_without_length = 0
    element_details = []
    markers = []      # list of (p0_host, p1_host, length_m)
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

        # Get length
        result = get_element_length(element)
        length, p0_local, p1_local = result if result else (None, None, None)

        if length and length > 0:
            total_length += length
            elements_with_length += 1
            element_details.append({
                'id': display_id,
                'name': element_name,
                'category': element_category,
                'length': length
            })
            table_rows.append([str(display_id), element_name, element_category,
                                "{:.2f} m".format(length * 0.3048)])

            if p0_local is not None and p1_local is not None:
                if transform is not None:
                    p0_host = transform.OfPoint(p0_local)
                    p1_host = transform.OfPoint(p1_local)
                else:
                    p0_host, p1_host = p0_local, p1_local
                markers.append((p0_host, p1_host, length * 0.3048))
        else:
            elements_without_length += 1
            not_found_ids.append(display_id)

    if table_rows:
        report.subheader("Measured Elements")
        report.table(["Element ID", "Name", "Category", "Length"], table_rows)

    if not_found_ids:
        report.warn("No length found for element ID(s): {}".format(
            ", ".join(str(i) for i in not_found_ids)))

    # Convert to display units
    total_length_mm = total_length * 304.8
    total_length_m = total_length * 0.3048

    report.subheader("Summary")
    report.line("Total elements selected: <b>{}</b>".format(len(picked_elements)))
    report.line("Elements with length: <b>{}</b> / without: <b>{}</b>".format(
        elements_with_length, elements_without_length))
    report.line("Total length: <b>{:.2f} m</b> ({:.0f} mm)".format(
        total_length_m, total_length_mm))

    # Draw a red dimension-style arrow along each measured segment,
    # with a red 3D digit readout of its length (meters) next to it.
    if markers:
        drawn = 0
        text_failed = 0
        try:
            with revit.Transaction("QTO Length Marker"):
                clear_old_markers()
                try:
                    facing_normal = doc.ActiveView.ViewDirection.Normalize()
                except Exception:
                    facing_normal = DB.XYZ(0, 0, 1)

                for p0, p1, length_m in markers:
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
        report = report if 'report' in globals() else ADAReport(__title__)
        report.error("Error: {}".format(e))
        report.flush()
        import traceback
        print(traceback.format_exc())
