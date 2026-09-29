# -*- coding: utf-8 -*-
__doc__ = """Copy the crop region of one viewport to other viewports on the
same sheet.

  1. tick the reference viewport in the list of the active sheet's
     viewports - its crop region is saved,
  2. tick one or more of the other viewports of the sheet,
  3. the saved crop region is applied to all of them.

Rectangular crops (rotated or not) and custom-shaped crops are both
copied, at the same place in the model. The crop is turned on in the
target views if it was off. If a target's crop is driven by a Scope
Box, its Scope Box is set to None so the copied crop overrides it.
On Revit 2022+ each target viewport is then moved back so the model
stays where it was on the sheet.

Views are skipped (and listed in the report) when they look in a
different direction than the reference (e.g. a section vs a plan).
A reference with a split crop region cannot be copied. The whole
change is one undoable transaction."""
__title__ = "Copy Crop\nRegion"
__version__ = "Version 1.0"
__author__ = "ADA"

import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import (
    BoundingBoxXYZ, BuiltInParameter, CurveLoop, ElementId, SubTransaction,
    Transaction, Transform, View3D, ViewSheet, ViewType, XYZ
)

from pyrevit import forms, revit, script

# Custom ADA GUI - themed list (lib/GUI/SelectFromDict.py) and the shared
# dark/gold themed report (lib/GUI/ReportTheme.py)
from GUI.forms import select_from_dict
from GUI.ReportTheme import ADAReport

doc = revit.doc

TITLE = __title__.replace("\n", " ")

NO_CROP_TYPES = (ViewType.DraftingView, ViewType.Legend, ViewType.Rendering,
                 ViewType.DrawingSheet, ViewType.Report)


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


def clear_scope_box(view):
    """Set the view's Scope Box to None. Returns the removed scope box's
    name, or None if the view had no scope box."""
    param = view.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
    if param is None or param.AsElementId() == ElementId.InvalidElementId:
        return None
    scope_box = doc.GetElement(param.AsElementId())
    name = scope_box.Name if scope_box is not None else "?"
    if not param.Set(ElementId.InvalidElementId):
        raise Exception("could not remove the Scope Box '{}'".format(name))
    return name


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

def croppable_viewports(sheet):
    viewports = [doc.GetElement(vp_id) for vp_id in sheet.GetAllViewports()]
    return [vp for vp in viewports if croppable(view_of(vp))]


def pick_source(viewports):
    options = {}
    for vp in viewports:
        label = viewport_label(vp)
        if not view_of(vp).CropBoxActive:
            label += "  - crop off"
        options[label] = vp

    chosen = select_from_dict(
        options,
        title=TITLE,
        label="Tick the reference viewport (its crop region is copied):",
        button_name="Next",
        version=__version__,
        SelectMultiple=False)
    if not chosen:
        script.exit()
    return chosen[0]


def pick_targets(viewports, source):
    options = {viewport_label(vp): vp for vp in viewports
               if vp.Id != source.Id}
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

    viewports = croppable_viewports(sheet)
    if len(viewports) < 2:
        forms.alert("This sheet needs at least two viewports with a crop "
                    "region (plans, sections, elevations...).",
                    title=TITLE, exitscript=True)

    source = pick_source(viewports)
    source_view = view_of(source)
    if not source_view.CropBoxActive:
        forms.alert("The crop region of '{}' is turned off - turn it on and "
                    "set it first.".format(source_view.Name),
                    title=TITLE, exitscript=True)

    saved = SavedCrop(source_view)
    if saved.is_split or len(saved.loops) > 1:
        forms.alert("'{}' has a split crop region - split crops cannot be "
                    "copied.".format(source_view.Name),
                    title=TITLE, exitscript=True)

    targets = pick_targets(viewports, source)

    report = ADAReport(TITLE)
    report.line("Sheet: <b>{} - {}</b>".format(sheet.SheetNumber, sheet.Name))
    report.line("Reference: <b>{}</b> &nbsp;({} crop)".format(
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
                                    "direction than the reference"])
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
                removed_scope_box = clear_scope_box(view)
                view.CropBoxActive = True
                saved.apply(view)
                doc.Regenerate()

                notes = ""
                if removed_scope_box:
                    notes += " (Scope Box '{}' set to None)".format(
                        removed_scope_box)

                # keep the model where it was on the sheet
                kept = notes
                if anchor_before is not None:
                    to_sheet = model_to_sheet(vp)
                    if to_sheet is not None:
                        drift = anchor_before - to_sheet.OfPoint(XYZ.Zero)
                        drift = XYZ(drift.X, drift.Y, 0)
                        if drift.GetLength() > 1e-9:
                            vp.SetBoxCenter(vp.GetBoxCenter() + drift)
                        kept += " (model kept in place on the sheet)"
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
