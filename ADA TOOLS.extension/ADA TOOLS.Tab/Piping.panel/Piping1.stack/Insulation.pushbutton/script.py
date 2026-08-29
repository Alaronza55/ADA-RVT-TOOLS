# -*- coding: utf-8 -*-
__doc__ = """Align the bottom elevation of pipes to a reference pipe.

Workflow:
    1. Pick one REFERENCE pipe -> its bottom elevation is computed
       (centerline Z - outer radius - insulation thickness, if insulated).
    2. Pick one or more pipes to align.
    3. Each pipe is moved vertically so that its own bottom
       (insulation included, if it has any) matches the reference bottom.

Notes:
    - Insulation is hosted on the pipe, so it follows automatically.
    - Sloped pipes are aligned on their LOWEST point (flagged in the report)."""

__title__ = "Align Pipe\nBottoms"
__author__ = "ADA TOOLS"

from pyrevit import revit, DB, script

from Autodesk.Revit.DB import XYZ, ElementTransformUtils, BuiltInParameter
from Autodesk.Revit.DB.Plumbing import Pipe, PipeInsulation
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

# Shared ADA-Tools dark/gold themed report (see lib/GUI/ReportTheme.py)
from GUI.ReportTheme import ADAReport

doc = revit.doc
uidoc = revit.uidoc
out = script.get_output()
report = ADAReport(__title__.replace(chr(10), " "))

TOL = 1.0 / 304800.0  # ~0.001 mm in feet


# ---------------------------------------------------------------- selection
class PipeSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, Pipe)

    def AllowReference(self, reference, position):
        return False


# ---------------------------------------------------------------- helpers
def to_mm(value_ft):
    """Internal units (ft) -> mm, compatible with old and new API."""
    try:
        from Autodesk.Revit.DB import UnitUtils, UnitTypeId
        return DB.UnitUtils.ConvertFromInternalUnits(value_ft, UnitTypeId.Millimeters)
    except Exception:
        from Autodesk.Revit.DB import UnitUtils, DisplayUnitType
        return DB.UnitUtils.ConvertFromInternalUnits(
            value_ft, DisplayUnitType.DUT_MILLIMETERS)


def get_outer_diameter(pipe):
    """Outer diameter of the pipe in feet."""
    for bip in (BuiltInParameter.RBS_PIPE_OUTER_DIAMETER,
                BuiltInParameter.RBS_PIPE_DIAMETER_PARAM):
        try:
            p = pipe.get_Parameter(bip)
        except Exception:
            p = None
        if p and p.HasValue:
            val = p.AsDouble()
            if val > 0:
                return val

    for name in ("Outside Diameter", "Diametre exterieur", "Diameter"):
        p = pipe.LookupParameter(name)
        if p and p.HasValue and p.AsDouble() > 0:
            return p.AsDouble()

    return 0.0


def get_insulation_thickness(pipe):
    """Insulation thickness in feet (0.0 if the pipe is not insulated)."""
    ins_ids = PipeInsulation.GetInsulationIds(doc, pipe.Id)
    if ins_ids is None or ins_ids.Count == 0:
        return 0.0

    thicknesses = []
    for iid in ins_ids:
        ins = doc.GetElement(iid)
        if ins is None:
            continue
        try:
            thicknesses.append(ins.Thickness)
        except Exception:
            p = ins.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS)
            if p and p.HasValue:
                thicknesses.append(p.AsDouble())

    return max(thicknesses) if thicknesses else 0.0


def get_bottom_data(pipe):
    """Returns (bottom_z, radius, insulation_thickness, is_sloped) in feet."""
    curve = pipe.Location.Curve
    z0 = curve.GetEndPoint(0).Z
    z1 = curve.GetEndPoint(1).Z
    is_sloped = abs(z0 - z1) > TOL

    radius = get_outer_diameter(pipe) / 2.0
    insul = get_insulation_thickness(pipe)
    bottom = min(z0, z1) - radius - insul
    return bottom, radius, insul, is_sloped


# ---------------------------------------------------------------- 1. reference
try:
    ref_pick = uidoc.Selection.PickObject(
        ObjectType.Element, PipeSelectionFilter(),
        "Select the REFERENCE pipe (bottom elevation to match)")
except OperationCanceledException:
    script.exit()

ref_pipe = doc.GetElement(ref_pick.ElementId)
ref_bottom, ref_r, ref_ins, ref_sloped = get_bottom_data(ref_pipe)

report.subheader("Reference Pipe {}".format(report.link(ref_pipe.Id)))
report.line("Outer radius: <b>{:.1f} mm</b>".format(to_mm(ref_r)))
report.line("Insulation: <b>{}</b>".format(
    "{:.1f} mm".format(to_mm(ref_ins)) if ref_ins > 0 else "none"))
report.line("Bottom elevation (project): <b>{:.1f} mm</b>".format(to_mm(ref_bottom)))
if ref_sloped:
    report.warn("Reference pipe is sloped - lowest point used")

# ---------------------------------------------------------------- 2. targets
try:
    picks = uidoc.Selection.PickObjects(
        ObjectType.Element, PipeSelectionFilter(),
        "Select the pipes to ALIGN, then click Finish")
except OperationCanceledException:
    script.exit()

targets = []
for p in picks:
    if p.ElementId == ref_pipe.Id:
        continue
    el = doc.GetElement(p.ElementId)
    if isinstance(el, Pipe):
        targets.append(el)

if not targets:
    script.exit()

# ---------------------------------------------------------------- 3. align
rows = []
moved = 0
failed = 0

with revit.Transaction("Align pipe bottoms"):
    for pipe in targets:
        try:
            bottom, radius, insul, sloped = get_bottom_data(pipe)
            dz = ref_bottom - bottom

            if abs(dz) < TOL:
                rows.append([report.link(pipe.Id),
                             "{:.1f}".format(to_mm(insul)) if insul else "-",
                             "0.0", "already aligned"])
                continue

            ElementTransformUtils.MoveElement(doc, pipe.Id, XYZ(0, 0, dz))
            moved += 1
            rows.append([report.link(pipe.Id),
                         "{:.1f}".format(to_mm(insul)) if insul else "-",
                         "{:+.1f}".format(to_mm(dz)),
                         "sloped - lowest point" if sloped else "OK"])

        except Exception as ex:
            failed += 1
            rows.append([report.link(pipe.Id), "-", "-",
                         "FAILED: {}".format(ex)])

# ---------------------------------------------------------------- 4. report
report.subheader("Result - {} moved, {} failed".format(moved, failed))
report.table(["Pipe", "Insulation (mm)", "dZ (mm)", "Status"], rows)
report.flush()