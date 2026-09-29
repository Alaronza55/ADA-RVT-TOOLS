# -*- coding: utf-8 -*-
__doc__ = """Pick two curves (in the current model and/or in a linked model)
and get the distance between them.

Curves are picked the same way as in "Get Length (Curve)": edges of
solid geometry (walls, pipes, framing...) as well as Model Lines /
Detail Lines / Reference Lines. The distance is measured from the
MIDPOINT of the shorter of the two curves to the closest point of the
other curve - for two parallel edges that is their perpendicular
distance, with the arrow centred on the shorter edge.

Visualization: a blue double-headed arrow is drawn from that midpoint
to the other curve (lifted slightly toward the view so it is not
hidden in the surface), with short witness lines back to the curves
and a blue 3D digit readout of the distance in meters next to it.

Markers from earlier runs are kept - measure as many pairs as you
like; use "Clear QTO Markers" to remove them all."""
__title__ = "Distance\n2 Curves"
__version__ = "Version = 1.0"
__author__ = "ADA"

import math

from pyrevit import revit, DB, UI
from pyrevit import forms, script
from System.Collections.Generic import List

# Custom ADA GUI - small button-choice popup and the shared dark/gold
# themed report (see lib/GUI/ReportTheme.py)
from GUI.forms import select_from_buttons
from GUI.ReportTheme import ADAReport

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

MARKER_NAME = "ADA_QTO_DistanceArrowMarker"
TEXT_MARKER_NAME = "ADA_QTO_DistanceText"

SOURCE_CURRENT = "Both in Current Model"
SOURCE_LINKED = "Both in Linked Model"
SOURCE_MIXED = "1st Current, 2nd Linked"

# Size of the whole marker (arrow + digits) relative to the other QTO tools
MARKER_SCALE = 0.5

ARROWHEAD_LENGTH = 0.35 * MARKER_SCALE   # feet
ARROWHEAD_RADIUS = 0.13 * MARKER_SCALE   # feet
ARROWHEAD_SIDES = 16
SHAFT_RADIUS = 0.045 * MARKER_SCALE      # feet
SHAFT_SIDES = 12
WITNESS_RADIUS = 0.03 * MARKER_SCALE     # feet
WITNESS_SIDES = 8
ARROW_LIFT = 0.15 * MARKER_SCALE         # feet, toward the viewer so the arrow is not buried in the face
TEXT_STANDOFF = 0.6 * MARKER_SCALE       # feet, sideways (in the view plane) from the arrow's middle
TOUCH_TOL = 1e-4                         # feet, below this the curves are considered touching

MARKER_COLOR = DB.Color(30, 90, 210)     # blue
MARKER_LINE_COLOR = DB.Color(0, 0, 0)    # black edges
DIGIT_COLOR = DB.Color(30, 90, 210)      # blue, matches the arrow
DIGIT_OFFSET = 0.05 * MARKER_SCALE       # feet, nudge digits toward the viewer

# --- 7-segment digit geometry (same technique as Get Length / Get Surface) -
DIGIT_W = 0.95 * MARKER_SCALE
DIGIT_H = 1.75 * MARKER_SCALE
STROKE = 0.24 * MARKER_SCALE
DIGIT_GAP = 0.30 * MARKER_SCALE
DOT_W = 0.42 * MARKER_SCALE
DEPTH = 0.13 * MARKER_SCALE

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


# ---------------------------------------------------------------------------
# marker geometry (same as Get Length (Curve))
# ---------------------------------------------------------------------------

def get_solid_fill_pattern_id():
    for fp in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
        try:
            if fp.GetFillPattern().IsSolidFill:
                return fp.Id
        except Exception:
            continue
    return DB.ElementId.InvalidElementId


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
        [p[(x0, y0, z0)], p[(x1, y0, z0)], p[(x1, y1, z0)], p[(x0, y1, z0)]],
        [p[(x0, y0, z1)], p[(x0, y1, z1)], p[(x1, y1, z1)], p[(x1, y0, z1)]],
        [p[(x0, y0, z0)], p[(x0, y1, z0)], p[(x0, y1, z1)], p[(x0, y0, z1)]],
        [p[(x1, y0, z0)], p[(x1, y0, z1)], p[(x1, y1, z1)], p[(x1, y1, z0)]],
        [p[(x0, y0, z0)], p[(x0, y0, z1)], p[(x1, y0, z1)], p[(x1, y0, z0)]],
        [p[(x0, y1, z0)], p[(x1, y1, z0)], p[(x1, y1, z1)], p[(x0, y1, z1)]],
    ]


def face_reading_basis(normal):
    """(u, v) = (right, up) perpendicular to `normal`, with v as close to
    world-up as possible."""
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
    """Raised 7-segment-style blocks for `text` (digits and '.')."""
    faces = []
    cursor = 0.0
    for ch in text:
        if ch == '.':
            faces.extend(box_faces(origin, u, v, n,
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
            faces.extend(box_faces(origin, u, v, n,
                                   cursor + sx0, cursor + sx1, sy0, sy1,
                                   0.0, DEPTH))
        cursor += DIGIT_W + DIGIT_GAP

    return faces


def build_cone_faces(tip, normal, length, radius, sides):
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

    loops = [list(ring)]
    for i in range(sides):
        loops.append([ring[i], ring[(i + 1) % sides], tip])
    return loops


def build_cylinder_faces(p0, p1, radius, sides):
    axis = p1.Subtract(p0)
    if axis.GetLength() < 1e-6:
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

    loops = [list(reversed(ring0)), list(ring1)]
    for i in range(sides):
        j = (i + 1) % sides
        loops.append([ring0[i], ring1[i], ring1[j], ring0[j]])
    return loops


def build_arrow_faces(p0, p1, p0_true, p1_true):
    """Double-headed arrow p0 -> p1 plus a witness line from each end
    back to the true measured point."""
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

    for p_draw, p_true in ((p0, p0_true), (p1, p1_true)):
        if p_draw.Subtract(p_true).GetLength() > 1e-6:
            faces.extend(build_cylinder_faces(p_true, p_draw, WITNESS_RADIUS, WITNESS_SIDES))
    return faces


def create_marker_shape(face_loops, name, color, line_weight):
    builder = DB.TessellatedShapeBuilder()
    builder.OpenConnectedFaceSet(False)
    for loop in face_loops:
        builder.AddFace(DB.TessellatedFace(
            List[DB.XYZ](loop), DB.ElementId.InvalidElementId))
    builder.CloseConnectedFaceSet()
    builder.Target = DB.TessellatedShapeBuilderTarget.AnyGeometry
    builder.Fallback = DB.TessellatedShapeBuilderFallback.Mesh
    builder.Build()
    geom_objs = list(builder.GetBuildResult().GetGeometricalObjects())

    ds = DB.DirectShape.CreateElement(
        doc, DB.ElementId(DB.BuiltInCategory.OST_GenericModel))
    ds.SetShape(List[DB.GeometryObject](geom_objs))
    ds.Name = name

    ogs = DB.OverrideGraphicSettings()
    ogs.SetProjectionLineColor(MARKER_LINE_COLOR)
    ogs.SetProjectionLineWeight(line_weight)
    fill_id = get_solid_fill_pattern_id()
    if fill_id != DB.ElementId.InvalidElementId:
        ogs.SetSurfaceForegroundPatternVisible(True)
        ogs.SetSurfaceForegroundPatternColor(color)
        ogs.SetSurfaceForegroundPatternId(fill_id)
    doc.ActiveView.SetElementOverrides(ds.Id, ogs)
    return ds


def create_distance_digits(origin_point, facing_normal, distance_m):
    u, v = face_reading_basis(facing_normal)
    text = "{:.2f}".format(distance_m)
    total_w = sum((DOT_W if c == '.' else DIGIT_W) + DIGIT_GAP for c in text) - DIGIT_GAP
    origin = origin_point.Add(u.Multiply(-total_w / 2.0))\
        .Add(v.Multiply(-DIGIT_H / 2.0))\
        .Add(facing_normal.Multiply(DIGIT_OFFSET))
    faces = build_number_faces(text, origin, u, v, facing_normal)
    return create_marker_shape(faces, TEXT_MARKER_NAME, DIGIT_COLOR, 2)


# ---------------------------------------------------------------------------
# picking (same technique as Get Length (Curve))
# ---------------------------------------------------------------------------

def collect_curves(geom_obj, curves):
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
    """Curve of element's geometry closest to `point` (element's own
    document coordinates)."""
    options = DB.Options()
    options.DetailLevel = DB.ViewDetailLevel.Fine
    options.IncludeNonVisibleObjects = True
    geom = element.get_Geometry(options)
    if not geom:
        return None

    all_curves = []
    for g in geom:
        collect_curves(g, all_curves)

    best_curve, best_dist = None, None
    for c in all_curves:
        try:
            result = c.Project(point)
        except Exception:
            result = None
        if result is not None and (best_dist is None or result.Distance < best_dist):
            best_dist, best_curve = result.Distance, c
    return best_curve


def pick_curve(linked, ordinal):
    """Pick one curve; returns (points_in_host_coords, midpoint_in_host_coords,
    curve_length_ft, description)."""
    if linked:
        ref = uidoc.Selection.PickObject(
            UI.Selection.ObjectType.LinkedElement,
            "Select the {} curve/edge in the linked model".format(ordinal))
        link_instance = doc.GetElement(ref.ElementId)
        linked_doc = link_instance.GetLinkDocument()
        element = linked_doc.GetElement(ref.LinkedElementId)
        transform = link_instance.GetTotalTransform()
        curve = find_curve_at_point(element, transform.Inverse.OfPoint(ref.GlobalPoint))
        desc = "{} (in link: {})".format(ref.LinkedElementId.IntegerValue,
                                         link_instance.Name)
    else:
        ref = uidoc.Selection.PickObject(
            UI.Selection.ObjectType.Edge,
            "Select the {} curve/edge".format(ordinal))
        element = doc.GetElement(ref.ElementId)
        geom_obj = element.GetGeometryObjectFromReference(ref)
        curve = None
        if isinstance(geom_obj, DB.Edge):
            curve = geom_obj.AsCurve()
        elif isinstance(geom_obj, DB.Curve):
            curve = geom_obj
        transform = None
        desc = output.linkify(ref.ElementId, title=str(ref.ElementId.IntegerValue))

    if curve is None:
        forms.alert("The {} pick did not resolve to a curve.".format(ordinal),
                    exitscript=True)

    points = list(curve.Tessellate())
    midpoint = curve.Evaluate(0.5, True)
    if transform is not None:
        points = [transform.OfPoint(p) for p in points]
        midpoint = transform.OfPoint(midpoint)
    return points, midpoint, curve.Length, desc


# ---------------------------------------------------------------------------
# closest points between two polylines
# ---------------------------------------------------------------------------

def closest_points_segments(p1, q1, p2, q2):
    """Closest points between segments p1-q1 and p2-q2 (Ericson,
    Real-Time Collision Detection 5.1.9). Returns (c1, c2)."""
    d1 = q1.Subtract(p1)
    d2 = q2.Subtract(p2)
    r = p1.Subtract(p2)
    a = d1.DotProduct(d1)
    e = d2.DotProduct(d2)
    f = d2.DotProduct(r)
    eps = 1e-12

    if a <= eps and e <= eps:
        return p1, p2
    if a <= eps:
        s = 0.0
        t = min(max(f / e, 0.0), 1.0)
    else:
        c = d1.DotProduct(r)
        if e <= eps:
            t = 0.0
            s = min(max(-c / a, 0.0), 1.0)
        else:
            b = d1.DotProduct(d2)
            denom = a * e - b * b
            s = min(max((b * f - c * e) / denom, 0.0), 1.0) if denom > eps else 0.0
            t = (b * s + f) / e
            if t < 0.0:
                t = 0.0
                s = min(max(-c / a, 0.0), 1.0)
            elif t > 1.0:
                t = 1.0
                s = min(max((b - c) / a, 0.0), 1.0)

    return p1.Add(d1.Multiply(s)), p2.Add(d2.Multiply(t))


def closest_point_on_polyline(point, pts):
    """Closest point to `point` on the polyline `pts`."""
    best, best_d = None, None
    for i in range(len(pts) - 1):
        _, c = closest_points_segments(point, point, pts[i], pts[i + 1])
        d = point.DistanceTo(c)
        if best_d is None or d < best_d:
            best, best_d = c, d
    return best


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

try:
    report = ADAReport()
    title = __title__.replace(chr(10), " ")

    source = select_from_buttons(
        [SOURCE_CURRENT, SOURCE_LINKED, SOURCE_MIXED],
        title=title,
        label="Where are the two curves?",
        version=__version__)
    if not source:
        forms.alert("Cancelled.", exitscript=True)

    first_linked = source == SOURCE_LINKED
    second_linked = source in (SOURCE_LINKED, SOURCE_MIXED)

    pts_a, mid_a, len_a, desc_a = pick_curve(first_linked, "first")
    pts_b, mid_b, len_b, desc_b = pick_curve(second_linked, "second")

    # measure from the midpoint of the shorter curve to the other curve
    if len_a <= len_b:
        pa, from_label = mid_a, "1st"
        pb = closest_point_on_polyline(pa, pts_b)
    else:
        pa, from_label = mid_b, "2nd"
        pb = closest_point_on_polyline(pa, pts_a)
    dist = pa.DistanceTo(pb)
    dist_m = dist * 0.3048
    delta = pb.Subtract(pa)
    horizontal_m = math.sqrt(delta.X ** 2 + delta.Y ** 2) * 0.3048
    vertical_m = abs(delta.Z) * 0.3048

    report.header(title)
    report.subheader("Picked Curves")
    report.table(["Curve", "Element ID", "Curve length"],
                 [["1st", desc_a, "{:.2f} m".format(len_a * 0.3048)],
                  ["2nd", desc_b, "{:.2f} m".format(len_b * 0.3048)]])

    report.subheader("Distance")
    report.line("Measured from the midpoint of the shorter curve ({}) to the "
                "closest point of the other curve.".format(from_label))
    report.line("Distance: <b>{:.3f} m</b> ({:.0f} mm)".format(dist_m, dist * 304.8))
    report.line("Horizontal component: {:.3f} m &nbsp;|&nbsp; Vertical component: "
                "{:.3f} m".format(horizontal_m, vertical_m))

    if dist < TOUCH_TOL:
        report.warn("The two curves touch or cross - no arrow drawn.")
        report.flush()
        script.exit()

    view = doc.ActiveView
    try:
        toward_camera = view.ViewDirection.Normalize()
        up = view.UpDirection.Normalize()
    except Exception:
        toward_camera, up = DB.XYZ(0, 0, 1), DB.XYZ(0, 1, 0)

    lift = toward_camera.Multiply(ARROW_LIFT)
    pa_draw, pb_draw = pa.Add(lift), pb.Add(lift)

    direction = delta.Normalize()
    side = toward_camera.CrossProduct(direction)
    side = side.Normalize() if side.GetLength() > 1e-6 else up
    midpoint = DB.XYZ((pa_draw.X + pb_draw.X) / 2.0,
                      (pa_draw.Y + pb_draw.Y) / 2.0,
                      (pa_draw.Z + pb_draw.Z) / 2.0)
    text_origin = midpoint.Add(side.Multiply(TEXT_STANDOFF + DIGIT_H / 2.0))

    try:
        with revit.Transaction("QTO Distance Marker"):
            faces = build_arrow_faces(pa_draw, pb_draw, pa, pb)
            if faces:
                create_marker_shape(faces, MARKER_NAME, MARKER_COLOR, 3)
            try:
                create_distance_digits(text_origin, toward_camera, dist_m)
            except Exception as text_err:
                report.warn("Could not build 3D distance digits: {}".format(text_err))
        uidoc.RefreshActiveView()
        report.success("Distance arrow drawn (blue). Earlier markers were kept.")
    except Exception as marker_err:
        report.error("Could not draw the distance marker: {}".format(marker_err))

    report.flush()

except Exception as e:
    if 'cancel' not in str(e).lower():
        report = report if 'report' in globals() else ADAReport()
        report.error("Error: {}".format(e))
        report.flush()
        import traceback
        print(traceback.format_exc())
