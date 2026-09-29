# -*- coding: utf-8 -*-
__doc__ = """Copy the crop region of one viewport to other viewports on the
same sheet.

  1. select the source viewport on the sheet (or select it before
     launching the tool) - its crop region is saved,
  2. pick one or more of the other viewports of the sheet from the
     list,
  3. the saved crop region is applied to all of them.

Rectangular crops (rotated or not) and custom-shaped crops are both
copied, at the same place in the model. The crop is turned on in the
target views if it was off. On Revit 2022+ each target viewport is
then moved back so the model stays where it was on the sheet.

Views are skipped (and listed in the report) when they look in a
different direction than the source (e.g. a section vs a plan), when
their crop is driven by a Scope Box, or when the source has a split
crop region. The whole change is one undoable transaction."""
__title__ = "Copy Crop\nRegion"
__version__ = "Version 1.0"
__author__ = "ADA"

import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import (
    BoundingBoxXYZ, BuiltInParameter, CurveLoop, ElementId, SubTransaction,
    Transaction, Transform, View3D, ViewSheet, ViewType, Viewport, XYZ
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import forms, revit, script

# Custom ADA GUI - themed list (lib/GUI/SelectFromDict.py) and the shared
# dark/gold themed report (lib/GUI/ReportTheme.py)
from GUI.forms import select_from_dict
from GUI.ReportTheme import ADAReport

doc = revit.doc
uidoc = revit.uidoc

TITLE = __title__.replace("\n", " ")

NO_CROP_TYPES = (ViewType.DraftingView, ViewType.Legend, ViewType.Rendering,
                 ViewType.DrawingSheet, ViewType.Report)


class ViewportFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, Viewport)

    def AllowReference(self, reference, point):
        return False


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def view_of(viewport):
    return doc.GetElement(viewport.ViewId)


def viewport_label(viewport):
    view = view_of(viewport)
    return "{}  ({}, 1:{})".format(view.Name, view.ViewType, view.Scale)


def croppable(view):
    if view.ViewType in NO_CROP_TYPES:
        return False
    if isinstance(view, View3D) and view.IsPerspective:
        return False
    return True


def scope_box_driven(view):
    param = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
    return (param is not None
            and param.AsElementId() != ElementId.InvalidElementId)


def model_to_sheet(viewport):
    """Model -> sheet transform of a viewport (Revit 2022+), or None."""
    try:
        view = view_of(viewport)
        projection = view.GetModelToProjectionTransforms()[0]\
            .GetModelToProjectionTransform()
        return viewport.GetProjectionToSheetTransform().Multiply(projection)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# saved crop
# ---------------------------------------------------------------------------

class SavedCrop(object):
    """Crop region of the source view, stored in model coordinates."""

    def __init__(self, view):
        self.direction = view.ViewDirection.Normalize()
        box = view.CropBox
        self.box_transform = box.Transform
        self.box_min = box.Min
        self.box_max = box.Max

        manager = view.GetCropRegionShapeManager()
        self.loops = list(manager.GetCropShape()) if manager.ShapeSet else []
        self.is_split = manager.Split if hasattr(manager, "Split") else False

    @property
    def is_shape(self):
        return bool(self.loops)

    def offset_to(self, view):
        """Vector along the view direction from the source view plane to
        the target view plane (plans at another level, parallel sections)."""
        delta = view.CropBox.Transform.Origin - self.box_transform.Origin
        return self.direction.Multiply(delta.DotProduct(self.direction))

    def apply(self, view):
        offset = self.offset_to(view)
        manager = view.GetCropRegionShapeManager()

        if self.is_shape:
            move = Transform.CreateTranslation(offset)
            loop = CurveLoop.CreateViaTransform(self.loops[0], move)
            manager.SetCropShape(loop)
            return

        # rectangular: same plan position / rotation as the source, but keep
        # the target's own depth (far clip / view range) along the direction
        if manager.ShapeSet:
            manager.RemoveCropRegionShape()
        old = view.CropBox
        new = BoundingBoxXYZ()
        transform = Transform(self.box_transform)
        transform.Origin = self.box_transform.Origin + offset
        new.Transform = transform
        new.Min = XYZ(self.box_min.X, self.box_min.Y, old.Min.Z)
        new.Max = XYZ(self.box_max.X, self.box_max.Y, old.Max.Z)
        view.CropBox = new


# ---------------------------------------------------------------------------
# user input
# ---------------------------------------------------------------------------

def pick_source():
    preselected = [doc.GetElement(eid)
                   for eid in uidoc.Selection.GetElementIds()]
    viewports = [el for el in preselected if isinstance(el, Viewport)]
    if len(viewports) == 1:
        return viewports[0]

    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element, ViewportFilter(),
            "Select the viewport whose crop region should be copied")
    except OperationCanceledException:
        script.exit()
    return doc.GetElement(ref.ElementId)


def pick_targets(sheet, source):
    options = {}
    for vp_id in sheet.GetAllViewports():
        if vp_id == source.Id:
            continue
        vp = doc.GetElement(vp_id)
        if croppable(view_of(vp)):
            options[viewport_label(vp)] = vp
    if not options:
        forms.alert("There are no other croppable viewports on this sheet.",
                    title=TITLE, exitscript=True)

    chosen = select_from_dict(
        options,
        title=TITLE,
        label="Apply the crop of '{}' to:".format(view_of(source).Name),
        button_name="Apply Crop",
        version=__version__,
        SelectMultiple=True)
    if not chosen:
        script.exit()
    return chosen


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    sheet = doc.ActiveView
    if not isinstance(sheet, ViewSheet):
        forms.alert("Open a sheet first - the viewports must be on the "
                    "active sheet.", title=TITLE, exitscript=True)

    source = pick_source()
    source_view = view_of(source)
    if not croppable(source_view):
        forms.alert("'{}' has no crop region ({}).".format(
                    source_view.Name, source_view.ViewType),
                    title=TITLE, exitscript=True)
    if not source_view.CropBoxActive:
        forms.alert("The crop region of '{}' is turned off - turn it on and "
                    "set it first.".format(source_view.Name),
                    title=TITLE, exitscript=True)

    saved = SavedCrop(source_view)
    if saved.is_split or len(saved.loops) > 1:
        forms.alert("'{}' has a split crop region - split crops cannot be "
                    "copied.".format(source_view.Name),
                    title=TITLE, exitscript=True)

    targets = pick_targets(sheet, source)

    report = ADAReport(TITLE)
    report.line("Sheet: <b>{} - {}</b>".format(sheet.SheetNumber, sheet.Name))
    report.line("Source: <b>{}</b> &nbsp;({} crop)".format(
        viewport_label(source),
        "custom-shaped" if saved.is_shape else "rectangular"))

    rows = []
    done = 0
    with Transaction(doc, "ADA - Copy Crop Region") as t:
        t.Start()
        for vp in targets:
            view = view_of(vp)
            label = viewport_label(vp)

            if not view.ViewDirection.Normalize().IsAlmostEqualTo(
                    saved.direction, 1e-6):
                rows.append([label, "Skipped - looks in a different "
                                    "direction than the source"])
                continue
            if scope_box_driven(view):
                rows.append([label, "Skipped - crop is driven by a Scope Box"])
                continue

            anchor_before = None
            to_sheet = model_to_sheet(vp)
            if to_sheet is not None:
                anchor_before = to_sheet.OfPoint(XYZ.Zero)

            sub = SubTransaction(doc)
            sub.Start()
            try:
                was_pinned = vp.Pinned
                if was_pinned:
                    vp.Pinned = False
                view.CropBoxActive = True
                saved.apply(view)
                doc.Regenerate()

                # keep the model where it was on the sheet
                kept = ""
                if anchor_before is not None:
                    to_sheet = model_to_sheet(vp)
                    if to_sheet is not None:
                        drift = anchor_before - to_sheet.OfPoint(XYZ.Zero)
                        drift = XYZ(drift.X, drift.Y, 0)
                        if drift.GetLength() > 1e-9:
                            vp.SetBoxCenter(vp.GetBoxCenter() + drift)
                        kept = " (model kept in place on the sheet)"
                if was_pinned:
                    vp.Pinned = True
                sub.Commit()
                rows.append([label, "Crop applied" + kept])
                done += 1
            except Exception as ex:
                sub.RollBack()
                rows.append([label, "Failed - {}".format(ex)])
        t.Commit()

    report.subheader("Target viewports")
    report.table(["View", "Result"], rows)

    problems = len(rows) - done
    if problems:
        report.warn("{} viewport(s) not changed - see the table.".format(
            problems))
    if done:
        report.success("Crop region applied to {} viewport(s).".format(done))
    report.flush()


if __name__ == "__main__":
    main()
