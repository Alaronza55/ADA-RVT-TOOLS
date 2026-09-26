# -*- coding: utf-8 -*-
__doc__ = """Align viewports on a sheet so their grids (and the whole model)
line up exactly.

  1. select the viewports on the sheet (or select them before
     launching the tool),
  2. choose the reference viewport - it stays where it is,
  3. choose whether to align in both directions, horizontally only
     or vertically only.

Every other viewport is moved so that the same model point lands on
the same spot of the sheet as in the reference viewport - so grids,
walls and everything else sit exactly on top of each other, no matter
the crop size or the annotations around each view.

Views can only match if they have the same scale and the same
orientation (same view direction, same crop rotation). Viewports that
differ from the reference are skipped and listed in the report.
Pinned viewports are unpinned for the move and re-pinned. The whole
change is one undoable transaction. Requires Revit 2022 or newer."""
__title__ = "Align\nViewports"
__version__ = "Version 1.0"
__author__ = "ADA"

import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import Transaction, Viewport, ViewSheet, XYZ
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import forms, revit, script

# Custom ADA GUI - themed single-choice list (lib/GUI/SelectFromDict.py),
# button-choice popup (lib/GUI/SelectFromButtons.py) and the shared
# dark/gold themed report (lib/GUI/ReportTheme.py)
from GUI.forms import select_from_dict, select_from_buttons
from GUI.ReportTheme import ADAReport

doc = revit.doc
uidoc = revit.uidoc

TITLE = __title__.replace("\n", " ")
FEET_TO_MM = 304.8
TOL = 1e-7   # feet on sheet

ALIGN_BOTH = "Both directions"
ALIGN_X = "Horizontally only"
ALIGN_Y = "Vertically only"


class ViewportFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, Viewport)

    def AllowReference(self, reference, point):
        return False


# ---------------------------------------------------------------------------
# geometry
# ---------------------------------------------------------------------------

def model_to_sheet(viewport):
    """Transform model -> sheet for a viewport (Revit 2022+ API).
    Raises if the view cannot be mapped (perspective view...)."""
    view = doc.GetElement(viewport.ViewId)
    to_projection = view.GetModelToProjectionTransforms()
    if to_projection.Count == 0:
        raise Exception("view has no model projection")
    projection = to_projection[0].GetModelToProjectionTransform()
    return viewport.GetProjectionToSheetTransform().Multiply(projection)


def sheet_xy(vector):
    return XYZ(vector.X, vector.Y, 0)


def same_mapping(t_ref, t_other):
    """True if both transforms map model directions to the same sheet
    directions - same scale, view direction and rotation."""
    for axis in (XYZ.BasisX, XYZ.BasisY, XYZ.BasisZ):
        a = sheet_xy(t_ref.OfVector(axis))
        b = sheet_xy(t_other.OfVector(axis))
        if not a.IsAlmostEqualTo(b, 1e-6):
            return False
    return True


# ---------------------------------------------------------------------------
# user input
# ---------------------------------------------------------------------------

def get_viewports(sheet):
    preselected = [doc.GetElement(eid)
                   for eid in uidoc.Selection.GetElementIds()]
    viewports = [el for el in preselected if isinstance(el, Viewport)]
    if len(viewports) >= 2:
        return viewports

    try:
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, ViewportFilter(),
            "Select the viewports to align, then click Finish")
    except OperationCanceledException:
        script.exit()
    viewports = [doc.GetElement(r.ElementId) for r in refs]
    if len(viewports) < 2:
        forms.alert("Select at least two viewports.", title=TITLE,
                    exitscript=True)
    return viewports


def viewport_label(viewport):
    view = doc.GetElement(viewport.ViewId)
    return "{}  (1:{})".format(view.Name, view.Scale)


def pick_reference(viewports):
    options = {viewport_label(vp): vp for vp in viewports}
    chosen = select_from_dict(
        options,
        title=TITLE,
        label="Reference viewport (stays in place):",
        button_name="Next",
        version=__version__,
        SelectMultiple=False)
    if not chosen:
        script.exit()
    return chosen[0]


def pick_mode():
    mode = select_from_buttons(
        [ALIGN_BOTH, ALIGN_X, ALIGN_Y],
        title=TITLE,
        label="Align the other viewports to the reference:",
        version=__version__)
    if not mode:
        script.exit()
    return mode


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def fmt_mm(value):
    return "{:.2f}".format(value * FEET_TO_MM)


def main():
    sheet = doc.ActiveView
    if not isinstance(sheet, ViewSheet):
        forms.alert("Open a sheet first - the viewports to align must be on "
                    "the active sheet.", title=TITLE, exitscript=True)
    if not hasattr(Viewport, "GetProjectionToSheetTransform"):
        forms.alert("This tool needs Revit 2022 or newer.", title=TITLE,
                    exitscript=True)

    viewports = get_viewports(sheet)
    reference = pick_reference(viewports)
    mode = pick_mode()

    report = ADAReport(TITLE)
    report.line("Sheet: <b>{} - {}</b>".format(sheet.SheetNumber, sheet.Name))
    report.line("Reference: <b>{}</b> &nbsp;|&nbsp; Mode: <b>{}</b>".format(
        viewport_label(reference), mode))

    try:
        t_ref = model_to_sheet(reference)
    except Exception as ex:
        report.error("The reference view cannot be used: {}".format(ex))
        report.flush()
        return

    # Any model point works: when the mappings share scale and orientation,
    # matching one point matches them all.
    anchor = XYZ.Zero
    target = t_ref.OfPoint(anchor)

    moves = []      # (viewport, delta)
    rows = []
    for vp in viewports:
        if vp.Id == reference.Id:
            continue
        label = viewport_label(vp)
        try:
            t_vp = model_to_sheet(vp)
        except Exception as ex:
            rows.append([label, "-", "-", "Skipped - {}".format(ex)])
            continue
        if not same_mapping(t_ref, t_vp):
            rows.append([label, "-", "-",
                         "Skipped - different scale or orientation"])
            continue

        delta = sheet_xy(target - t_vp.OfPoint(anchor))
        if mode == ALIGN_X:
            delta = XYZ(delta.X, 0, 0)
        elif mode == ALIGN_Y:
            delta = XYZ(0, delta.Y, 0)

        if delta.GetLength() < TOL:
            rows.append([label, "0.00", "0.00", "Already aligned"])
            continue
        moves.append((vp, delta))
        rows.append([label, fmt_mm(delta.X), fmt_mm(delta.Y), "Moved"])

    if moves:
        try:
            with Transaction(doc, "ADA - Align Viewports") as t:
                t.Start()
                for vp, delta in moves:
                    was_pinned = vp.Pinned
                    if was_pinned:
                        vp.Pinned = False
                    vp.SetBoxCenter(vp.GetBoxCenter() + delta)
                    if was_pinned:
                        vp.Pinned = True
                t.Commit()
        except Exception as ex:
            report.error("Failed, nothing was changed: {}".format(ex))
            report.flush()
            return

    report.subheader("Viewports")
    report.table(["View", "Move X (mm)", "Move Y (mm)", "Result"], rows)

    skipped = [r for r in rows if r[3].startswith("Skipped")]
    if skipped:
        report.warn("{} viewport(s) skipped - grids can only match between "
                    "views with the same scale and orientation as the "
                    "reference.".format(len(skipped)))
    if moves:
        report.success("{} viewport(s) aligned to the reference.".format(
            len(moves)))
    elif not skipped:
        report.success("All viewports were already aligned - nothing was "
                       "moved.")
    report.flush()


if __name__ == "__main__":
    main()
