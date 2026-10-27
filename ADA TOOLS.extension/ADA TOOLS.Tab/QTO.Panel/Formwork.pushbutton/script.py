# -*- coding: utf-8 -*-
__doc__ = """Formwork surface of structural elements (current model or a link).

  1. choose where the elements are (current model / linked model),
  2. pick the elements - only structural elements can be picked
     (structural framing, structural columns, foundations, floors,
     walls, stairs).

Formwork faces = every face of the element's solid geometry EXCEPT the
faces looking up (the top of the pour needs no shuttering): for a cube,
the face below + the 4 sides. Sloped soffits count; upward slopes count
only when steeper than 60 deg from horizontal. Faces hidden by joins
(e.g. a beam end inside a column) are already removed by Revit's join
geometry.

  3. answer "Is this the first pour?":
     - First pour: nothing is built around yet - every face except the
       top is formwork.
     - Next pour: the parts of those faces touching a neighbouring
       structural element (beam end against a wall, slab under a beam,
       column side against a wall...) are not poured against formwork,
       so they are deducted. Each flat face is cut into small triangles
       and probed outward: a neighbour closer than 3 cm counts as
       contact. Neighbours are searched in the same model as the
       element (inside the link for linked elements). Curved faces are
       counted in full.

Each element's formwork faces are drawn as a semi-transparent surface
(no text) - they stay in the model like the other QTO markers; use
"Clear QTO Markers" to remove them. The report gives the bottom, side
and total formwork area per element, in m2."""
__title__ = "Formwork\nSurface"
__version__ = "Version 1.0"
__author__ = "ADA"

from pyrevit import revit, DB, UI
from pyrevit import forms
from System.Collections.Generic import List

# Shared ADA-Tools dark/gold themed report (lib/GUI/ReportTheme.py) and
# small button-choice popup (lib/GUI/SelectFromButtons.py)
from GUI.ReportTheme import ADAReport
from GUI.forms import select_from_buttons

doc = revit.doc
uidoc = revit.uidoc

TITLE = __title__.replace("\n", " ")
SQFT_TO_M2 = 0.09290304

MARKER_NAME = "ADA_QTO_FormworkMarker"
MARKER_OFFSET = 0.03                    # feet, lift off the real face to avoid z-fighting
MARKER_COLOR = DB.Color(0, 170, 170)    # teal (Get Surface is orange)
MARKER_LINE_COLOR = DB.Color(0, 90, 90)
MARKER_TRANSPARENCY = 50                # % transparent

UP_LIMIT = 0.5          # normal.Z above this = top face (within 60 deg of up): no formwork
DOWN_LIMIT = -0.5       # normal.Z below this = bottom / soffit face
CONTACT_GAP = 0.1       # feet (3 cm): a neighbour closer than this to a face touches it
PROBE_START = 0.003     # feet, probes start just off the face
MAX_EDGE = 1.0          # feet, triangles are split down to this size before being tested
MIN_EDGE = 0.1          # feet (3 cm), finest split along the contact boundaries

FIRST_POUR = "First pour"
NEXT_POUR = "Next pour (deduct contacts)"

STRUCTURAL_CATEGORIES = set(int(c) for c in (
    DB.BuiltInCategory.OST_StructuralFraming,
    DB.BuiltInCategory.OST_StructuralColumns,
    DB.BuiltInCategory.OST_StructuralFoundation,
    DB.BuiltInCategory.OST_Floors,
    DB.BuiltInCategory.OST_Walls,
    DB.BuiltInCategory.OST_Stairs,
))


def eid_value(eid):
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue


def is_structural(element):
    cat = element.Category if element is not None else None
    return cat is not None and eid_value(cat.Id) in STRUCTURAL_CATEGORIES


class HostStructuralFilter(UI.Selection.ISelectionFilter):
    def AllowElement(self, element):
        return is_structural(element)

    def AllowReference(self, reference, point):
        return False


class LinkStructuralFilter(UI.Selection.ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, DB.RevitLinkInstance)

    def AllowReference(self, reference, point):
        link = doc.GetElement(reference.ElementId)
        if not isinstance(link, DB.RevitLinkInstance):
            return False
        link_doc = link.GetLinkDocument()
        return link_doc is not None and is_structural(
            link_doc.GetElement(reference.LinkedElementId))


# ---------------------------------------------------------------------------
# geometry
# ---------------------------------------------------------------------------

def collect_solids(geom_obj, solids):
    if isinstance(geom_obj, DB.Solid):
        if geom_obj.Faces.Size > 0 and geom_obj.Volume > 1e-9:
            solids.append(geom_obj)
    elif isinstance(geom_obj, DB.GeometryInstance):
        inst = geom_obj.GetInstanceGeometry()
        if inst:
            for g in inst:
                collect_solids(g, solids)


def element_solids(element):
    options = DB.Options()
    options.DetailLevel = DB.ViewDetailLevel.Fine
    geom = element.get_Geometry(options)
    solids = []
    if geom:
        for g in geom:
            collect_solids(g, solids)
    return solids


def face_normal(face):
    if isinstance(face, DB.PlanarFace):
        return face.FaceNormal
    box = face.GetBoundingBox()
    mid = DB.UV((box.Min.U + box.Max.U) / 2.0, (box.Min.V + box.Max.V) / 2.0)
    return face.ComputeNormal(mid)


def points_box(points):
    return (min(p.X for p in points), min(p.Y for p in points), min(p.Z for p in points),
            max(p.X for p in points), max(p.Y for p in points), max(p.Z for p in points))


def boxes_overlap(a, b, margin=0.0):
    return all(a[i] - margin <= b[i + 3] and b[i] - margin <= a[i + 3] for i in range(3))


def solid_box(solid):
    bb = solid.GetBoundingBox()
    t = bb.Transform
    return points_box([t.OfPoint(DB.XYZ(x, y, z))
                       for x in (bb.Min.X, bb.Max.X)
                       for y in (bb.Min.Y, bb.Max.Y)
                       for z in (bb.Min.Z, bb.Max.Z)])


def neighbour_solids(element):
    """[(solid, box)] of the structural elements around `element`, in the
    same document (coordinates of that document)."""
    document = element.Document
    bb = element.get_BoundingBox(None)
    if bb is None:
        return []
    m = DB.XYZ(CONTACT_GAP * 2, CONTACT_GAP * 2, CONTACT_GAP * 2)
    near = DB.BoundingBoxIntersectsFilter(DB.Outline(bb.Min - m, bb.Max + m))
    result = []
    for other in (DB.FilteredElementCollector(document).WherePasses(near)
                  .WhereElementIsNotElementType()):
        if other.Id == element.Id or not is_structural(other):
            continue
        for solid in element_solids(other):
            try:
                result.append((solid, solid_box(solid)))
            except Exception:
                continue
    return result


class ContactTester(object):
    """Tells if a neighbouring element sits right outside a planar face at a
    given point: a short segment is probed from just off the face, along
    its normal, up to CONTACT_GAP (so small gaps count as contact)."""

    def __init__(self, neighbours, normal):
        self.neighbours = neighbours
        self.start = normal.Multiply(PROBE_START)
        self.end = normal.Multiply(CONTACT_GAP)
        self.options = DB.SolidCurveIntersectionOptions()

    def covered(self, p):
        a, b = p.Add(self.start), p.Add(self.end)
        seg_box = points_box([a, b])
        line = None
        for solid, box in self.neighbours:
            if not boxes_overlap(seg_box, box):
                continue
            if line is None:
                line = DB.Line.CreateBound(a, b)
            try:
                if solid.IntersectWithCurve(line, self.options).SegmentCount > 0:
                    return True
            except Exception:
                continue
        return False


def bisect(tri):
    """Split a triangle in two at the middle of its longest edge."""
    a, b, c = tri
    edges = [(a.DistanceTo(b), a, b, c), (b.DistanceTo(c), b, c, a),
             (c.DistanceTo(a), c, a, b)]
    _, p, q, r = max(edges, key=lambda e: e[0])
    m = p.Add(q).Multiply(0.5)
    return [(p, m, r), (m, q, r)]


def longest_edge(tri):
    a, b, c = tri
    return max(a.DistanceTo(b), b.DistanceTo(c), c.DistanceTo(a))


def triangle_area(tri):
    a, b, c = tri
    return b.Subtract(a).CrossProduct(c.Subtract(a)).GetLength() / 2.0


def uncovered_triangles(triangles, tester):
    """Triangles (refined) of a planar face not touched by a neighbour:
    each triangle is probed at its centre and near its corners; mixed
    results are split again, down to MIN_EDGE along the contact edge."""
    keep = []
    stack = list(triangles)
    while stack:
        tri = stack.pop()
        edge = longest_edge(tri)
        if edge > MAX_EDGE:
            stack.extend(bisect(tri))
            continue
        a, b, c = tri
        centre = a.Add(b).Add(c).Divide(3.0)
        probes = [centre] + [v.Add(centre.Subtract(v).Multiply(0.2)) for v in tri]
        states = [tester.covered(p) for p in probes]
        if not any(states):
            keep.append(tri)
        elif all(states):
            continue
        elif edge <= MIN_EDGE:
            if not states[0]:
                keep.append(tri)
        else:
            stack.extend(bisect(tri))
    return keep


def triangle_normal(face, tri, default):
    if isinstance(face, DB.PlanarFace):
        return default
    try:
        a, b, c = tri
        uv = face.Project(a.Add(b).Add(c).Divide(3.0)).UVPoint
        return face.ComputeNormal(uv).Normalize()
    except Exception:
        return default


def formwork_faces(element, transform, deduct):
    """[(kind, gross area, net area, [(a, b, c, normal)])] for every
    formwork face of the element - kind 'bottom' or 'side'. With `deduct`,
    the parts touching neighbouring elements are removed from the net
    area and from the triangles (element's document coordinates)."""
    neighbours = neighbour_solids(element) if deduct else []
    result = []
    for solid in element_solids(element):
        for face in solid.Faces:
            n_doc = face_normal(face).Normalize()
            n = transform.OfVector(n_doc).Normalize() if transform is not None else n_doc
            if n.Z > UP_LIMIT:
                continue                       # top of the pour
            kind = "bottom" if n.Z < DOWN_LIMIT else "side"
            mesh = face.Triangulate()
            tris = [tuple(mesh.get_Triangle(i).get_Vertex(j) for j in range(3))
                    for i in range(mesh.NumTriangles)]
            net, kept = face.Area, tris
            if neighbours and tris and isinstance(face, DB.PlanarFace):
                face_box = points_box([p for t in tris for p in t])
                near = [nb for nb in neighbours
                        if boxes_overlap(face_box, nb[1], CONTACT_GAP)]
                if near:
                    try:
                        kept = uncovered_triangles(tris, ContactTester(near, n_doc))
                        net = sum(triangle_area(t) for t in kept)
                    except Exception:
                        net, kept = face.Area, tris   # contact test failed: count in full
            result.append((kind, face.Area, net,
                           [t + (triangle_normal(face, t, n_doc),) for t in kept]))
    return result


def marker_triangles(triangles, transform):
    """Triangles in host coords, lifted along their normal."""
    out = []
    for a, b, c, normal in triangles:
        pts = [a, b, c]
        if transform is not None:
            pts = [transform.OfPoint(p) for p in pts]
            normal = transform.OfVector(normal)
        shift = normal.Multiply(MARKER_OFFSET)
        out.append([p.Add(shift) for p in pts])
    return out


def get_solid_fill_pattern_id():
    for fp in DB.FilteredElementCollector(doc).OfClass(DB.FillPatternElement):
        try:
            if fp.GetFillPattern().IsSolidFill:
                return fp.Id
        except Exception:
            continue
    return DB.ElementId.InvalidElementId


def create_marker(triangles, fill_id):
    builder = DB.TessellatedShapeBuilder()
    builder.OpenConnectedFaceSet(False)
    for tri in triangles:
        builder.AddFace(DB.TessellatedFace(List[DB.XYZ](tri),
                                           DB.ElementId.InvalidElementId))
    builder.CloseConnectedFaceSet()
    builder.Target = DB.TessellatedShapeBuilderTarget.AnyGeometry
    builder.Fallback = DB.TessellatedShapeBuilderFallback.Mesh
    builder.Build()
    geom_objs = list(builder.GetBuildResult().GetGeometricalObjects())

    ds = DB.DirectShape.CreateElement(
        doc, DB.ElementId(DB.BuiltInCategory.OST_GenericModel))
    ds.SetShape(List[DB.GeometryObject](geom_objs))
    ds.Name = MARKER_NAME

    ogs = DB.OverrideGraphicSettings()
    ogs.SetProjectionLineColor(MARKER_LINE_COLOR)
    ogs.SetSurfaceTransparency(MARKER_TRANSPARENCY)
    if fill_id != DB.ElementId.InvalidElementId:
        ogs.SetSurfaceForegroundPatternVisible(True)
        ogs.SetSurfaceForegroundPatternColor(MARKER_COLOR)
        ogs.SetSurfaceForegroundPatternId(fill_id)
    doc.ActiveView.SetElementOverrides(ds.Id, ogs)
    return ds


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def pick_elements():
    """[(element, transform or None, location text, host id or None)]"""
    source = select_from_buttons(["Current Model", "Linked Model"],
                                 title=TITLE,
                                 label="Pick structural element(s) in:",
                                 version=__version__)
    if not source:
        forms.alert("Cancelled.", exitscript=True)

    picked = []
    if source == "Current Model":
        pre = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
        pre = [e for e in pre if is_structural(e)]
        if pre:
            return [(e, None, "current model", e.Id) for e in pre]
        refs = uidoc.Selection.PickObjects(
            UI.Selection.ObjectType.Element, HostStructuralFilter(),
            "Select structural elements, then click Finish")
        for ref in refs:
            e = doc.GetElement(ref.ElementId)
            picked.append((e, None, "current model", e.Id))
    else:
        refs = uidoc.Selection.PickObjects(
            UI.Selection.ObjectType.LinkedElement, LinkStructuralFilter(),
            "Select structural elements in the link, then click Finish")
        for ref in refs:
            link = doc.GetElement(ref.ElementId)
            e = link.GetLinkDocument().GetElement(ref.LinkedElementId)
            picked.append((e, link.GetTotalTransform(),
                           u"link: {}".format(link.Name), None))
    return picked


try:
    picked = pick_elements()
    if not picked:
        forms.alert("No structural element selected.", exitscript=True)

    pour = select_from_buttons([FIRST_POUR, NEXT_POUR],
                               title=TITLE,
                               label="Is this the first pour? (first pour = nothing "
                                     "around yet: every face except the top)",
                               version=__version__)
    if not pour:
        forms.alert("Cancelled.", exitscript=True)
    deduct = pour == NEXT_POUR

    report = ADAReport(TITLE)
    report.line(u"Mode: <b>{}</b>".format(pour))
    rows = []
    markers = []          # list of triangle lists, one per element
    total_bottom = total_side = total_contact = 0.0

    for element, transform, location, host_id in picked:
        try:
            faces = formwork_faces(element, transform, deduct)
        except Exception as geo_err:
            report.warn(u"Element {}: could not read its geometry ({})".format(
                eid_value(element.Id), geo_err))
            continue
        if not faces:
            report.warn(u"Element {} ({}): no solid geometry - skipped".format(
                eid_value(element.Id), element.Category.Name))
            continue

        bottom = sum(n for k, _, n, _ in faces if k == "bottom") * SQFT_TO_M2
        side = sum(n for k, _, n, _ in faces if k == "side") * SQFT_TO_M2
        gross = sum(g for _, g, _, _ in faces) * SQFT_TO_M2
        contact = max(0.0, gross - bottom - side)
        total_bottom += bottom
        total_side += side
        total_contact += contact

        id_cell = report.link(host_id, title=str(eid_value(element.Id))) \
            if host_id is not None else str(eid_value(element.Id))
        rows.append([id_cell, element.Category.Name, location,
                     u"{:.2f}".format(bottom), u"{:.2f}".format(side),
                     u"<b>{:.2f}</b>".format(bottom + side),
                     u"{:.2f}".format(contact)])

        tris = []
        for _, _, _, kept in faces:
            tris.extend(marker_triangles(kept, transform))
        if tris:
            markers.append(tris)

    if rows:
        report.subheader("Formwork per element (m2)")
        report.table(["Element ID", "Category", "Location", "Bottom", "Sides",
                      "Total", "Contact (deducted)"], rows)

    report.subheader("Summary")
    report.line(u"Elements: <b>{}</b>".format(len(rows)))
    report.line(u"Bottom / soffit formwork: <b>{:.2f} m2</b>".format(total_bottom))
    report.line(u"Side formwork: <b>{:.2f} m2</b>".format(total_side))
    report.line(u"Total formwork: <b>{:.2f} m2</b>".format(total_bottom + total_side))
    report.line(u"Deducted (in contact with neighbouring elements): {:.2f} m2"
                .format(total_contact))

    if markers:
        try:
            fill_id = get_solid_fill_pattern_id()
            with revit.Transaction("QTO Formwork Marker"):
                for tris in markers:
                    create_marker(tris, fill_id)
            uidoc.RefreshActiveView()
            report.success(u"{} formwork marker(s) drawn (teal, {}% transparent)."
                           .format(len(markers), MARKER_TRANSPARENCY))
        except Exception as marker_err:
            report.error(u"Could not draw the formwork markers: {}".format(marker_err))

    report.flush()

except Exception as e:
    if 'cancel' not in str(e).lower():
        report = report if 'report' in globals() else ADAReport(TITLE)
        report.error(u"Error: {}".format(e))
        report.flush()
        import traceback
        print(traceback.format_exc())
