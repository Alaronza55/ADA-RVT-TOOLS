# -*- coding: utf-8 -*-
__doc__ = """Pick a Revit link, then pick a point of this model - Survey
Point, Project Base Point or Internal Origin. The link is moved so
that ITS Internal Origin sits exactly on the chosen point (X, Y and
elevation).

The link is only moved - its current rotation is kept.

Nothing inside the link file is modified - only the link instance is
placed. If it is pinned it is unpinned for the move and re-pinned
afterwards. The whole change is one undoable transaction."""
__title__ = "Align Link\nInternal Origin"
__version__ = "Version 1.0"
__author__ = "ADA"

import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import (
    BasePoint, BuiltInCategory, ElementTransformUtils,
    FilteredElementCollector, RevitLinkInstance, Transaction, XYZ
)

from pyrevit import forms, revit, script

# Custom ADA GUI - themed single-choice list (lib/GUI/SelectFromDict.py),
# button-choice popup (lib/GUI/SelectFromButtons.py) and the shared
# dark/gold themed report (lib/GUI/ReportTheme.py)
from GUI.forms import select_from_dict, select_from_buttons
from GUI.ReportTheme import ADAReport

doc = revit.doc

TITLE = __title__.replace("\n", " ")
FEET_TO_METERS = 0.3048
TOL_LENGTH = 1e-6   # feet

TARGET_SURVEY = "Survey Point"
TARGET_PBP = "Project Base Point"
TARGET_ORIGIN = "Internal Origin"


def eid_value(eid):
    """ElementId numeric value, 2024+ safe."""
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue


def first_of_category(category):
    return (FilteredElementCollector(doc)
            .OfCategory(category)
            .WhereElementIsNotElementType()
            .FirstElement())


def target_point(target):
    """The chosen point of this model, in its internal coordinates."""
    if target == TARGET_ORIGIN:
        return XYZ.Zero
    if target == TARGET_PBP:
        try:
            point = BasePoint.GetProjectBasePoint(doc)
        except AttributeError:
            point = first_of_category(BuiltInCategory.OST_ProjectBasePoint)
    else:
        try:
            point = BasePoint.GetSurveyPoint(doc)
        except AttributeError:
            point = first_of_category(BuiltInCategory.OST_SharedBasePoint)
    if point is None:
        forms.alert("Could not find the {} in this model.".format(target),
                    title=TITLE, exitscript=True)
    return point.Position


def pick_link():
    links = list(FilteredElementCollector(doc)
                 .OfClass(RevitLinkInstance)
                 .WhereElementIsNotElementType())
    if not links:
        forms.alert("There are no Revit links in this model.",
                    title=TITLE, exitscript=True)

    options = {}
    for link in links:
        name = link.Name
        if link.GetLinkDocument() is None:
            name += "  (not loaded)"
        options["{}  [{}]".format(name, eid_value(link.Id))] = link

    chosen = select_from_dict(
        options,
        title=TITLE,
        label="Select the link to align:",
        button_name="Next",
        version=__version__,
        SelectMultiple=False)
    if not chosen:
        script.exit()
    return chosen[0]


def pick_target():
    target = select_from_buttons(
        [TARGET_SURVEY, TARGET_PBP, TARGET_ORIGIN],
        title=TITLE,
        label="Align the link's Internal Origin to this model's:",
        version=__version__)
    if not target:
        script.exit()
    return target


def fmt_m(value):
    return "{:.4f}".format(value * FEET_TO_METERS)


def fmt_xyz_m(xyz):
    return [fmt_m(xyz.X), fmt_m(xyz.Y), fmt_m(xyz.Z)]


def main():
    link = pick_link()
    target = pick_target()
    goal_origin = target_point(target)

    # the link's Internal Origin, in this model's internal coordinates
    before = link.GetTotalTransform().Origin
    move_by = goal_origin - before

    report = ADAReport(TITLE)
    report.line("Link: <b>{}</b> &nbsp;{}".format(
        link.Name, report.link(link.Id, title="Select link")))
    report.line("Target: this model's <b>{}</b>".format(target))

    if move_by.GetLength() < TOL_LENGTH:
        report.success("The link's Internal Origin is already on the {} - "
                       "nothing was moved.".format(target))
        report.flush()
        return

    was_pinned = link.Pinned
    try:
        with Transaction(doc, "ADA - Align Link Internal Origin") as t:
            t.Start()
            if was_pinned:
                link.Pinned = False
            ElementTransformUtils.MoveElement(doc, link.Id, move_by)
            if was_pinned:
                link.Pinned = True
            t.Commit()
    except Exception as ex:
        report.error("Failed, nothing was changed: {}".format(ex))
        report.flush()
        return

    after = link.GetTotalTransform().Origin

    report.subheader("Link Internal Origin (this model's internal coordinates, m)")
    report.table(
        ["", "X", "Y", "Z"],
        [["Before"] + fmt_xyz_m(before),
         ["After"] + fmt_xyz_m(after),
         ["Target: " + target] + fmt_xyz_m(goal_origin)])

    report.subheader("Move applied to the link (m)")
    report.table(["X", "Y", "Z", "Distance"],
                 [fmt_xyz_m(move_by) + [fmt_m(move_by.GetLength())]])

    if was_pinned:
        report.line("The link was pinned - it was unpinned for the move and "
                    "re-pinned.")

    gap = (after - goal_origin).GetLength()
    if gap > TOL_LENGTH:
        report.warn("The link's Internal Origin is still {} m from the {} - "
                    "check the link placement.".format(fmt_m(gap), target))
    else:
        report.success("Link aligned - its Internal Origin is now on this "
                       "model's {}.".format(target))
    report.flush()


if __name__ == "__main__":
    main()
