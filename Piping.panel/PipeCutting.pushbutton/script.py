# -*- coding: utf-8 -*-
"""Cut (split) selected host pipes by the plane of a face picked in a linked model."""

__title__ = "Cut Pipes\nby Linked Face"
__author__ = "ADA TOOLS"
__doc__ = ("Pick a planar face inside a linked model, then select one or more "
           "pipes in the host model. Every pipe whose centerline crosses the "
           "infinite plane defined by that face is split at the crossing point.")

from Autodesk.Revit.DB import (
    Options, ViewDetailLevel, GeometryInstance, Solid, PlanarFace,
    Transaction, Line, ElementId, RevitLinkInstance
)
from Autodesk.Revit.DB.Plumbing import Pipe, PlumbingUtils
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, script

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

TOL = 1e-9
END_MARGIN = 1e-4          # feet: keep the break point off the exact endpoints
PARALLEL_TOL = 1e-6        # dot-product tolerance for "pipe parallel to plane"
FACE_HIT_TOL = 1e-3        # feet: max distance from picked point to resolved face


class PipeFilter(ISelectionFilter):
    """Allow only host-model pipes to be picked."""
    def AllowElement(self, element):
        return isinstance(element, Pipe)

    def AllowReference(self, reference, position):
        return False


def collect_planar_faces(geom_elem, faces):
    """Recursively gather PlanarFaces from a GeometryElement (handles nested instances)."""
    if geom_elem is None:
        return
    for obj in geom_elem:
        if isinstance(obj, GeometryInstance):
            collect_planar_faces(obj.GetInstanceGeometry(), faces)
        elif isinstance(obj, Solid):
            if obj.Faces.Size:
                for f in obj.Faces:
                    if isinstance(f, PlanarFace):
                        faces.append(f)


def find_face_at_point(faces, pt):
    """Return (face, distance) for the planar face closest to pt (link coords)."""
    best, best_d = None, None
    for f in faces:
        res = f.Project(pt)
        if res is not None:
            d = res.Distance
            if best_d is None or d < best_d:
                best, best_d = f, d
    return best, best_d


def line_plane_point(line, p_origin, p_normal):
    """Intersection of a segment with an infinite plane (host coords), or None."""
    p0 = line.GetEndPoint(0)
    p1 = line.GetEndPoint(1)
    vec = p1 - p0
    length = vec.GetLength()
    if length < TOL:
        return None
    direction = vec.Normalize()
    denom = p_normal.DotProduct(direction)
    if abs(denom) < PARALLEL_TOL:                    # pipe parallel to plane
        return None
    s = p_normal.DotProduct(p_origin - p0) / denom   # distance along direction
    if s <= END_MARGIN or s >= length - END_MARGIN:  # crossing outside segment
        return None
    return p0 + direction.Multiply(s)


# ---------------------------------------------------------------------------
# 1) Pick a face in a linked model
# ---------------------------------------------------------------------------
try:
    face_ref = uidoc.Selection.PickObject(
        ObjectType.LinkedElement,
        "Select a planar face in the linked model to define the cut plane")
except OperationCanceledException:
    script.exit()

link_inst = doc.GetElement(face_ref.ElementId)
if not isinstance(link_inst, RevitLinkInstance):
    output.print_md("**Error:** the picked reference is not a linked element.")
    script.exit()

link_doc = link_inst.GetLinkDocument()
if link_doc is None:
    output.print_md("**Error:** the linked model is not loaded.")
    script.exit()

link_tf = link_inst.GetTotalTransform()
linked_elem = link_doc.GetElement(face_ref.LinkedElementId)

opt = Options()
opt.ComputeReferences = False
opt.IncludeNonVisibleObjects = False
opt.DetailLevel = ViewDetailLevel.Fine

faces = []
collect_planar_faces(linked_elem.get_Geometry(opt), faces)
if not faces:
    output.print_md("**Error:** no planar faces found on the picked element.")
    script.exit()

# Resolve which face was clicked (work in the linked model's coordinates)
pt_link = link_tf.Inverse.OfPoint(face_ref.GlobalPoint)
face, dist = find_face_at_point(faces, pt_link)
if face is None or dist is None or dist > FACE_HIT_TOL:
    output.print_md("**Error:** could not resolve a flat face at the picked "
                    "point. Pick a planar face (not a curved surface).")
    script.exit()

# Build the cut plane in host coordinates
plane_origin = link_tf.OfPoint(face.Origin)
plane_normal = link_tf.OfVector(face.FaceNormal).Normalize()

# ---------------------------------------------------------------------------
# 2) Pick pipes in the host model
# ---------------------------------------------------------------------------
try:
    pipe_refs = uidoc.Selection.PickObjects(
        ObjectType.Element, PipeFilter(),
        "Select the pipes to cut, then click Finish")
except OperationCanceledException:
    script.exit()

pipe_ids = [r.ElementId for r in pipe_refs]

# ---------------------------------------------------------------------------
# 3) Split each crossing pipe at the plane
# ---------------------------------------------------------------------------
cut = skipped = failed = 0
t = Transaction(doc, "Cut pipes by linked face")
t.Start()
try:
    for pid in pipe_ids:
        pipe = doc.GetElement(pid)
        loc = pipe.Location
        curve = loc.Curve if loc is not None else None
        if not isinstance(curve, Line):
            skipped += 1
            continue

        hit = line_plane_point(curve, plane_origin, plane_normal)
        if hit is None:
            skipped += 1
            continue

        # snap the break point exactly onto the pipe centerline
        hit = curve.Project(hit).XYZPoint
        try:
            new_id = PlumbingUtils.BreakCurve(doc, pid, hit)
            if new_id is not None and new_id != ElementId.InvalidElementId:
                cut += 1
            else:
                failed += 1
        except Exception:
            failed += 1
    t.Commit()
except Exception as ex:
    t.RollBack()
    output.print_md("**Aborted:** {}".format(ex))
    script.exit()

output.print_md(
    "**Done.** Cut **{0}** pipe(s) &nbsp;|&nbsp; skipped {1} "
    "(no crossing / not straight) &nbsp;|&nbsp; failed {2}."
    .format(cut, skipped, failed))