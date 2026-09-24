# -*- coding: utf-8 -*-
__doc__ = """Rotate a Revit link around a point of this model.

  1. pick the link,
  2. pick the rotation centre - this model's Survey Point, Project
     Base Point or Internal Origin,
  3. pick the direction - clockwise or counter-clockwise (seen from
     above, in plan),
  4. type the angle in degrees ("." or "," accepted as decimal
     separator).

The link rotates around a vertical axis through the chosen point.
Nothing inside the link file is modified - only the link instance is
placed. If it is pinned it is unpinned for the rotation and re-pinned
afterwards. The whole change is one undoable transaction."""
__title__ = "Rotate\nLink"
__version__ = "Version 1.0"
__author__ = "ADA"

import math
import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import (
    BasePoint, BuiltInCategory, ElementTransformUtils,
    FilteredElementCollector, Line, RevitLinkInstance, Transaction, XYZ
)

from pyrevit import forms, revit, script

# Custom ADA GUI - themed single-choice list (lib/GUI/SelectFromDict.py),
# button-choice popup (lib/GUI/SelectFromButtons.py), text-input popup
# (lib/GUI/AskForInputs.py) and the shared dark/gold themed report
# (lib/GUI/ReportTheme.py)
from GUI.forms import select_from_dict, select_from_buttons, ask_for_inputs
from GUI.ReportTheme import ADAReport

doc = revit.doc

TITLE = __title__.replace("\n", " ")
FEET_TO_METERS = 0.3048

PIVOT_SURVEY = "Survey Point"
PIVOT_PBP = "Project Base Point"
PIVOT_ORIGIN = "Internal Origin"

CLOCKWISE = "Clockwise"
COUNTER_CLOCKWISE = "Counter-clockwise"


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


def pivot_point(pivot):
    """The chosen point of this model, in its internal coordinates."""
    if pivot == PIVOT_ORIGIN:
        return XYZ.Zero
    if pivot == PIVOT_PBP:
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
        forms.alert("Could not find the {} in this model.".format(pivot),
                    title=TITLE, exitscript=True)
    return point.Position


def instance_rotation(transform):
    """Rotation of a link instance around Z, counter-clockwise from +X."""
    return math.atan2(transform.BasisX.Y, transform.BasisX.X)


def parse_degrees(text):
    """Float from user text; accepts ',' or '.' as decimal separator.
    Returns None if it is not a number."""
    cleaned = (text or "").strip().replace(" ", "").replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# user input
# ---------------------------------------------------------------------------

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
        label="Select the link to rotate:",
        button_name="Next",
        version=__version__,
        SelectMultiple=False)
    if not chosen:
        script.exit()
    return chosen[0]


def pick_pivot():
    pivot = select_from_buttons(
        [PIVOT_SURVEY, PIVOT_PBP, PIVOT_ORIGIN],
        title=TITLE,
        label="Rotate the link around this model's:",
        version=__version__)
    if not pivot:
        script.exit()
    return pivot


def pick_direction():
    direction = select_from_buttons(
        [CLOCKWISE, COUNTER_CLOCKWISE],
        title=TITLE,
        label="Rotation direction (seen in plan, from above):",
        version=__version__)
    if not direction:
        script.exit()
    return direction


def ask_degrees(direction):
    """Ask the angle until a positive number is typed or the user cancels."""
    default = "90"
    while True:
        values = ask_for_inputs(
            [("Angle (degrees)", default)],
            title=TITLE,
            label="Rotate {}, by how many degrees?".format(direction.lower()),
            button_name="Rotate",
            version=__version__)
        if values is None:
            script.exit()

        degrees = parse_degrees(values[0])
        if degrees is not None and degrees > 0:
            return degrees

        forms.alert("Type a positive number of degrees (e.g. 90 or 12,5).\n"
                    "The direction was already chosen in the previous step.",
                    title=TITLE)
        default = values[0]


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def fmt_xyz_m(xyz):
    return ["{:.4f}".format(v * FEET_TO_METERS) for v in (xyz.X, xyz.Y, xyz.Z)]


def fmt_deg(angle_rad):
    return u"{:.4f}°".format(math.degrees(angle_rad))


def main():
    link = pick_link()
    pivot = pick_pivot()
    direction = pick_direction()
    degrees = ask_degrees(direction)

    center = pivot_point(pivot)
    # Revit rotates counter-clockwise (seen from above) for positive angles
    angle = math.radians(degrees)
    if direction == CLOCKWISE:
        angle = -angle

    before = link.GetTotalTransform()

    report = ADAReport(TITLE)
    report.line("Link: <b>{}</b> &nbsp;{}".format(
        link.Name, report.link(link.Id, title="Select link")))
    report.line(u"Rotation: <b>{:g}° {}</b> around this model's "
                "<b>{}</b>".format(degrees, direction.lower(), pivot))

    was_pinned = link.Pinned
    try:
        with Transaction(doc, "ADA - Rotate Link") as t:
            t.Start()
            if was_pinned:
                link.Pinned = False
            axis = Line.CreateBound(center, center + XYZ.BasisZ)
            ElementTransformUtils.RotateElement(doc, link.Id, axis, angle)
            if was_pinned:
                link.Pinned = True
            t.Commit()
    except Exception as ex:
        report.error("Failed, nothing was changed: {}".format(ex))
        report.flush()
        return

    after = link.GetTotalTransform()

    report.subheader("Link Internal Origin (this model's internal "
                     "coordinates, m)")
    report.table(
        ["", "X", "Y", "Z", "Link rotation (CCW from X)"],
        [["Before"] + fmt_xyz_m(before.Origin)
         + [fmt_deg(instance_rotation(before))],
         ["After"] + fmt_xyz_m(after.Origin)
         + [fmt_deg(instance_rotation(after))],
         ["Rotation centre: " + pivot] + fmt_xyz_m(center) + [""]])

    if was_pinned:
        report.line("The link was pinned - it was unpinned for the rotation "
                    "and re-pinned.")
    report.success(u"Link rotated {:g}° {}.".format(
        degrees, direction.lower()))
    report.flush()


if __name__ == "__main__":
    main()
