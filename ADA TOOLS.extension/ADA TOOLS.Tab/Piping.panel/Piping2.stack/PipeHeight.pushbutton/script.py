# -*- coding: utf-8 -*-
__doc__ = """Align the centerline of a first picked pipe to the centerline Z of a second
picked pipe, evaluated at the exact 2D (X,Y) intersection of the two centerlines.

Both pipes may be sloped: the first pipe is translated purely along Z, so its
slope, length and direction are preserved. Only its elevation changes.

ADA TOOLS - Coordination
IronPython 2 / Revit API"""

__title__ = "Align Pipe Z\nat XY Cross"
__author__ = "ADA"

import math

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    XYZ,
    Line,
    Transaction,
    TransactionStatus,
    ElementTransformUtils,
    UnitUtils,
)
from Autodesk.Revit.DB.Plumbing import Pipe
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import revit, forms, script

# Shared ADA-Tools dark/gold themed report (see lib/GUI/ReportTheme.py)
from GUI.ReportTheme import ADAReport

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

# Set to True to also accept ducts, cable trays, conduits (any MEPCurve).
ALLOW_ANY_MEP_CURVE = False

# Feet. 1e-4 ft is about 0.03 mm.
ZERO_TOL = 1e-4

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def to_mm(value_ft):
    """Convert an internal (feet) value to millimeters, 2021+ and legacy safe."""
    try:
        from Autodesk.Revit.DB import UnitTypeId
        return UnitUtils.ConvertFromInternalUnits(value_ft, UnitTypeId.Millimeters)
    except Exception:
        from Autodesk.Revit.DB import DisplayUnitType
        return UnitUtils.ConvertFromInternalUnits(
            value_ft, DisplayUnitType.DUT_MILLIMETERS
        )


class MEPCurveFilter(ISelectionFilter):
    """Restrict picking to pipes (or any MEPCurve if enabled)."""

    def AllowElement(self, element):
        if ALLOW_ANY_MEP_CURVE:
            from Autodesk.Revit.DB import MEPCurve
            return isinstance(element, MEPCurve)
        return isinstance(element, Pipe)

    def AllowReference(self, reference, position):
        return False


def pick_pipe(prompt):
    """Pick one pipe and return the element, or None if the user escapes."""
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element, MEPCurveFilter(), prompt
        )
    except OperationCanceledException:
        return None
    if ref is None:
        return None
    return doc.GetElement(ref.ElementId)


def get_centerline(element):
    """Return the straight centerline of an MEPCurve, or None."""
    loc = element.Location
    if loc is None or not hasattr(loc, "Curve"):
        return None
    curve = loc.Curve
    if not isinstance(curve, Line):
        return None
    return curve


def plan_intersection(line_a, line_b):
    """Intersect two 3D lines projected on the XY plane.

    Returns a dict with the normalized parameters on each line, the XY
    intersection point and the Z of each centerline at that point.
    Returns a string error code instead when no intersection exists.
    """
    p0 = line_a.GetEndPoint(0)
    p1 = line_a.GetEndPoint(1)
    q0 = line_b.GetEndPoint(0)
    q1 = line_b.GetEndPoint(1)

    dax, day, daz = p1.X - p0.X, p1.Y - p0.Y, p1.Z - p0.Z
    dbx, dby, dbz = q1.X - q0.X, q1.Y - q0.Y, q1.Z - q0.Z

    len_a = math.sqrt(dax * dax + day * day)
    len_b = math.sqrt(dbx * dbx + dby * dby)

    if len_a < 1e-9 or len_b < 1e-9:
        return "vertical"

    # Parallelism test on unit plan directions, so it is scale independent.
    cross_unit = (dax / len_a) * (dby / len_b) - (day / len_a) * (dbx / len_b)
    if abs(cross_unit) < 1e-6:
        return "parallel"

    det = dax * dby - day * dbx
    ex = q0.X - p0.X
    ey = q0.Y - p0.Y

    t = (ex * dby - ey * dbx) / det   # parameter on line A, 0..1 inside segment
    s = (ex * day - ey * dax) / det   # parameter on line B

    z_a = p0.Z + t * daz
    z_b = q0.Z + s * dbz

    return {
        "t": t,
        "s": s,
        "point": XYZ(p0.X + t * dax, p0.Y + t * day, 0.0),
        "z_a": z_a,
        "z_b": z_b,
        "dz": z_b - z_a,
    }


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    label = "element" if ALLOW_ANY_MEP_CURVE else "pipe"

    pipe_a = pick_pipe("Pick the %s TO MOVE (its Z will change)" % label.upper())
    if pipe_a is None:
        return

    pipe_b = pick_pipe("Pick the REFERENCE %s (stays in place)" % label.upper())
    if pipe_b is None:
        return

    if pipe_a.Id.IntegerValue == pipe_b.Id.IntegerValue:
        forms.alert("You picked the same element twice.", exitscript=True)

    line_a = get_centerline(pipe_a)
    line_b = get_centerline(pipe_b)

    if line_a is None or line_b is None:
        forms.alert(
            "One of the picked elements has no straight centerline.\n"
            "Flexible or curved runs are not supported.",
            exitscript=True,
        )

    result = plan_intersection(line_a, line_b)

    if result == "vertical":
        forms.alert(
            "One of the picked elements is vertical (a riser).\n"
            "It has no direction in plan, so no XY intersection can be found.",
            exitscript=True,
        )
    if result == "parallel":
        forms.alert(
            "The two centerlines are parallel in plan.\n"
            "They never cross, so no intersection Z can be computed.",
            exitscript=True,
        )

    dz = result["dz"]

    # Warn when the crossing point falls outside one of the real segments.
    outside = []
    if not (-1e-6 <= result["t"] <= 1.0 + 1e-6):
        outside.append("the first")
    if not (-1e-6 <= result["s"] <= 1.0 + 1e-6):
        outside.append("the second")
    if outside:
        proceed = forms.alert(
            "The XY intersection lies outside %s picked segment.\n"
            "The centerlines were extended to find it.\n\n"
            "Required move: %.1f mm\n\nApply anyway?"
            % (" and ".join(outside), to_mm(dz)),
            yes=True,
            no=True,
        )
        if not proceed:
            return

    if abs(dz) < ZERO_TOL:
        forms.alert(
            "Both centerlines already share the same Z at the crossing point.\n"
            "Nothing to move.",
            exitscript=True,
        )

    trans = Transaction(doc, "Align pipe centerline Z at XY crossing")
    trans.Start()
    try:
        ElementTransformUtils.MoveElement(doc, pipe_a.Id, XYZ(0.0, 0.0, dz))
        trans.Commit()
    except Exception as err:
        if trans.GetStatus() == TransactionStatus.Started:
            trans.RollBack()
        forms.alert(
            "The move failed.\n\n%s\n\n"
            "This usually happens when the element is connected to fittings "
            "or equipment that constrain it. Disconnect it and retry."
            % str(err),
            exitscript=True,
        )

    pt = result["point"]
    report = ADAReport(__title__.replace(chr(10), " "))
    report.subheader("Pipe Centerline Aligned")
    report.line("Moved: {}".format(report.link(pipe_a.Id)))
    report.line("Reference: {}".format(report.link(pipe_b.Id)))
    report.line("Crossing point (X, Y): <b>{:.1f} , {:.1f} mm</b>".format(
        to_mm(pt.X), to_mm(pt.Y)))
    report.line("Z before: <b>{:.1f} mm</b> / Z target: <b>{:.1f} mm</b>".format(
        to_mm(result["z_a"]), to_mm(result["z_b"])))
    report.success("Vertical shift applied: <b>{:+.1f} mm</b>".format(to_mm(dz)))
    report.flush()


main()