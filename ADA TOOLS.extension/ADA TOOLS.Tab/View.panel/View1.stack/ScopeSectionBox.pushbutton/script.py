# -*- coding: utf-8 -*-
__doc__ = """Match a scope box and a 3D view's section box.

  1. pick the scope box in the list,
  2. choose the direction:

  SECTION BOX -> SCOPE BOX: pick a 3D view with an active section box.
     The scope box is MOVED and ROTATED onto the section box (same
     centre, same orientation). The Revit API cannot resize a scope box,
     so the report gives the exact width / depth / height to set with
     the scope box handles (with its current size next to it).

  SCOPE BOX -> SECTION BOX: pick a 3D view. Its section box takes the
     exact extents and rotation of the scope box (fully automatic) and
     is switched on.

A pinned scope box is unpinned for the move and re-pinned. Each run is
one undoable transaction."""
__title__ = "Scope Box\n<> Section Box"
__version__ = "Version 1.0"
__author__ = "ADA"

import math

import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import (
    BoundingBoxXYZ, BuiltInCategory, ElementTransformUtils,
    FilteredElementCollector, Line, Options, Transaction, Transform, View3D,
    XYZ
)

from pyrevit import forms, revit, script

# Custom ADA GUI - themed list / buttons (lib/GUI) and the shared dark/gold
# themed report (lib/GUI/ReportTheme.py)
from GUI.forms import select_from_dict, select_from_buttons
from GUI.ReportTheme import ADAReport

doc = revit.doc

TITLE = __title__.replace("\n", " ")
FT_TO_M = 0.3048
TO_SCOPE = "Section box -> Scope box"
TO_SECTION = "Scope box -> Section box"


class Box(object):
    """Oriented box: centre (world), X / Y axes (horizontal), sizes along
    X, Y, Z (feet)."""

    def __init__(self, center, x_axis, y_axis, sx, sy, sz):
        self.center, self.x_axis, self.y_axis = center, x_axis, y_axis
        self.sx, self.sy, self.sz = sx, sy, sz

    @property
    def angle(self):
        return math.atan2(self.x_axis.Y, self.x_axis.X)

    def sizes_m(self):
        return [u"{:.3f}".format(v * FT_TO_M) for v in (self.sx, self.sy, self.sz)]


def scope_box_shape(scope_box):
    """Oriented box of a scope box, from its edge lines."""
    lines = [g for g in scope_box.get_Geometry(Options()) if isinstance(g, Line)]
    if not lines:
        raise Exception("the scope box has no readable geometry")
    points = [l.GetEndPoint(i) for l in lines for i in (0, 1)]
    zmin = min(p.Z for p in points)
    horizontal = [l for l in lines
                  if abs(l.GetEndPoint(0).Z - zmin) < 1e-6
                  and abs(l.GetEndPoint(1).Z - zmin) < 1e-6]
    d = horizontal[0].Direction
    x_axis = XYZ(d.X, d.Y, 0).Normalize()
    y_axis = XYZ.BasisZ.CrossProduct(x_axis)

    xs = [p.DotProduct(x_axis) for p in points]
    ys = [p.DotProduct(y_axis) for p in points]
    zs = [p.Z for p in points]
    center = (x_axis.Multiply((min(xs) + max(xs)) / 2.0)
              + y_axis.Multiply((min(ys) + max(ys)) / 2.0)
              + XYZ.BasisZ.Multiply((min(zs) + max(zs)) / 2.0))
    return Box(center, x_axis, y_axis,
               max(xs) - min(xs), max(ys) - min(ys), max(zs) - min(zs))


def section_box_shape(view):
    bb = view.GetSectionBox()
    t = bb.Transform
    mid = XYZ((bb.Min.X + bb.Max.X) / 2.0, (bb.Min.Y + bb.Max.Y) / 2.0,
              (bb.Min.Z + bb.Max.Z) / 2.0)
    x = XYZ(t.BasisX.X, t.BasisX.Y, 0).Normalize()
    return Box(t.OfPoint(mid), x, XYZ.BasisZ.CrossProduct(x),
               bb.Max.X - bb.Min.X, bb.Max.Y - bb.Min.Y, bb.Max.Z - bb.Min.Z)


def normalize_half_turn(angle):
    """Angle in (-90, 90] deg - a box looks the same after a half turn."""
    while angle <= -math.pi / 2:
        angle += math.pi
    while angle > math.pi / 2:
        angle -= math.pi
    return angle


def pick_scope_box():
    boxes = list(FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
                 .WhereElementIsNotElementType())
    if not boxes:
        forms.alert("There are no scope boxes in this project.", title=TITLE,
                    exitscript=True)
    chosen = select_from_dict({b.Name: b for b in boxes}, title=TITLE,
                              label="Scope box:", button_name="Next",
                              version=__version__, SelectMultiple=False)
    if not chosen:
        script.exit()
    return chosen[0]


def pick_view(with_section_box):
    views = [v for v in FilteredElementCollector(doc).OfClass(View3D)
             if not v.IsTemplate and (v.IsSectionBoxActive or not with_section_box)]
    if not views:
        forms.alert("There is no 3D view {}.".format(
            "with an active section box" if with_section_box else ""),
            title=TITLE, exitscript=True)
    label = ("3D view whose section box the scope box should follow:"
             if with_section_box else "3D view whose section box should "
                                      "take the scope box:")
    options = {}
    for v in views:
        key = v.Name + ("" if v.IsSectionBoxActive else "   (no section box yet)")
        options[key] = v
    chosen = select_from_dict(options, title=TITLE, label=label,
                              button_name="Apply", version=__version__,
                              SelectMultiple=False)
    if not chosen:
        script.exit()
    return chosen[0]


def scope_to_section(scope_box, view, report):
    box = scope_box_shape(scope_box)
    bb = BoundingBoxXYZ()
    t = Transform.CreateTranslation(box.center)
    t.BasisX = box.x_axis
    t.BasisY = box.y_axis
    t.BasisZ = XYZ.BasisZ
    bb.Transform = t
    bb.Min = XYZ(-box.sx / 2.0, -box.sy / 2.0, -box.sz / 2.0)
    bb.Max = XYZ(box.sx / 2.0, box.sy / 2.0, box.sz / 2.0)
    with Transaction(doc, "ADA - Section box = scope box") as tr:
        tr.Start()
        view.IsSectionBoxActive = True
        view.SetSectionBox(bb)
        tr.Commit()
    report.line(u"Section box of <b>{}</b> set to scope box <b>{}</b>.".format(
        view.Name, scope_box.Name))
    report.table(["", "X (m)", "Y (m)", "Z (m)"],
                 [["Section box = scope box"] + box.sizes_m()])
    report.success("Section box matches the scope box.")


def section_to_scope(scope_box, view, report):
    target = section_box_shape(view)
    current = scope_box_shape(scope_box)
    rotate = normalize_half_turn(target.angle - current.angle)
    move = target.center - current.center

    was_pinned = scope_box.Pinned
    with Transaction(doc, "ADA - Scope box onto section box") as tr:
        tr.Start()
        if was_pinned:
            scope_box.Pinned = False
        if abs(rotate) > 1e-9:
            axis = Line.CreateBound(current.center, current.center + XYZ.BasisZ)
            ElementTransformUtils.RotateElement(doc, scope_box.Id, axis, rotate)
        if move.GetLength() > 1e-9:
            ElementTransformUtils.MoveElement(doc, scope_box.Id, move)
        if was_pinned:
            scope_box.Pinned = True
        tr.Commit()

    after = scope_box_shape(scope_box)
    # express the scope box size along the section box axes
    if abs(after.x_axis.DotProduct(target.x_axis)) < 0.5:
        after.sx, after.sy = after.sy, after.sx
    report.line(u"Scope box <b>{}</b> moved and rotated onto the section box of "
                u"<b>{}</b> (rotation {:.2f} deg, move {:.3f} m).".format(
                    scope_box.Name, view.Name, math.degrees(rotate),
                    move.GetLength() * FT_TO_M))
    report.table(["Size along the section box", "X (m)", "Y (m)", "Z (m)"],
                 [["Target (section box)"] + target.sizes_m(),
                  ["Scope box now"] + after.sizes_m()])
    same = all(abs(a - b) < 1e-4 for a, b in
               ((target.sx, after.sx), (target.sy, after.sy), (target.sz, after.sz)))
    if same:
        report.success("The scope box now matches the section box exactly.")
    else:
        report.warn("Centre and rotation match. The Revit API cannot resize a "
                    "scope box: drag its handles to the target sizes above "
                    "(the box is centred, so drag both sides by half of the "
                    "difference).")


def main():
    scope_box = pick_scope_box()
    direction = select_from_buttons([TO_SCOPE, TO_SECTION], title=TITLE,
                                    label=u"'{}' - which one adapts?".format(
                                        scope_box.Name),
                                    version=__version__)
    if not direction:
        script.exit()
    view = pick_view(with_section_box=(direction == TO_SCOPE))

    report = ADAReport(TITLE)
    try:
        if direction == TO_SCOPE:
            section_to_scope(scope_box, view, report)
        else:
            scope_to_section(scope_box, view, report)
    except Exception as ex:
        report.error(u"Failed, nothing was changed: {}".format(ex))
    report.flush()


if __name__ == "__main__":
    main()
