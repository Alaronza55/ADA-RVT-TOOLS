# -*- coding: utf-8 -*-
__doc__ = """Move the Project Base Point to the coordinates you type, in
meters - North/South, East/West and Elevation, the same three values
shown in the Project Base Point's properties.

The window opens pre-filled with the current values, so you only need
to change the ones that differ. Both "." and "," are accepted as the
decimal separator.

Only the base point moves - the model geometry and the Survey Point
stay where they are, and the Angle to True North is not changed. If
the base point is pinned it is unpinned for the move and re-pinned
afterwards. The whole change is one undoable transaction."""
__title__ = "Set Project\nBase Point"
__version__ = "Version 1.0"
__author__ = "ADA"

import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import (
    BasePoint, BuiltInCategory, BuiltInParameter, FilteredElementCollector,
    Transaction
)

from pyrevit import forms, revit, script

# Custom ADA GUI - themed multi-field input popup (see
# lib/GUI/AskForInputs.py) and the shared dark/gold themed report
# (see lib/GUI/ReportTheme.py)
from GUI.forms import ask_for_inputs
from GUI.ReportTheme import ADAReport

doc = revit.doc

TITLE = __title__.replace("\n", " ")
FEET_TO_METERS = 0.3048

# (label, parameter) in the order they are asked
COORD_FIELDS = [
    ("North/South (m)", BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM),
    ("East/West (m)",   BuiltInParameter.BASEPOINT_EASTWEST_PARAM),
    ("Elevation (m)",   BuiltInParameter.BASEPOINT_ELEVATION_PARAM),
]


def get_project_base_point():
    try:
        return BasePoint.GetProjectBasePoint(doc)
    except AttributeError:
        # Revit < 2020
        return (FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                .WhereElementIsNotElementType()
                .FirstElement())


def read_coords_m(base_point):
    """[N/S, E/W, Elevation] of the base point, in meters."""
    return [base_point.get_Parameter(bip).AsDouble() * FEET_TO_METERS
            for _, bip in COORD_FIELDS]


def format_m(value):
    return "{:.4f}".format(value)


def parse_meters(text):
    """Float from user text; accepts ',' or '.' as decimal separator and
    ignores spaces. Returns None if it is not a number."""
    cleaned = (text or "").strip().replace(" ", "").replace(u" ", "")
    if "," in cleaned and "." in cleaned:
        cleaned = cleaned.replace(",", "")      # ',' used as thousands sep
    else:
        cleaned = cleaned.replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return None


def ask_new_coords(current_m):
    """Show the themed input window until the user enters valid numbers
    or cancels. Returns [N/S, E/W, Elevation] in meters."""
    defaults = [format_m(v) for v in current_m]
    while True:
        fields = [(label, defaults[i])
                  for i, (label, _) in enumerate(COORD_FIELDS)]
        values = ask_for_inputs(
            fields,
            title=TITLE,
            label="Project Base Point coordinates, in meters:",
            button_name="Set Base Point",
            version=__version__)
        if values is None:
            script.exit()

        parsed = [parse_meters(v) for v in values]
        bad = [COORD_FIELDS[i][0] for i, v in enumerate(parsed) if v is None]
        if not bad:
            return parsed

        forms.alert("Not a valid number:\n  " + "\n  ".join(bad),
                    title=TITLE)
        defaults = values   # re-open with what the user typed


def set_coords(base_point, new_m):
    was_pinned = base_point.Pinned
    with Transaction(doc, "ADA - Set Project Base Point") as t:
        t.Start()
        if was_pinned:
            base_point.Pinned = False
        for (label, bip), value_m in zip(COORD_FIELDS, new_m):
            param = base_point.get_Parameter(bip)
            if param is None or param.IsReadOnly:
                raise Exception("Parameter '{}' is read-only.".format(label))
            param.Set(value_m / FEET_TO_METERS)
        if was_pinned:
            base_point.Pinned = True
        t.Commit()
    return was_pinned


def main():
    base_point = get_project_base_point()
    if base_point is None:
        forms.alert("No Project Base Point found in this model.",
                    title=TITLE, exitscript=True)

    before_m = read_coords_m(base_point)
    new_m = ask_new_coords(before_m)

    report = ADAReport(TITLE)
    report.line("Project: <b>{}</b>".format(doc.Title))

    try:
        was_pinned = set_coords(base_point, new_m)
    except Exception as ex:
        report.error("Failed, nothing was changed: {}".format(ex))
        report.flush()
        return

    after_m = read_coords_m(base_point)

    report.subheader("Project Base Point")
    report.table(
        ["Coordinate", "Before", "After"],
        [[label, format_m(b), format_m(a)]
         for (label, _), b, a in zip(COORD_FIELDS, before_m, after_m)])

    mismatch = [label for (label, _), want, got
                in zip(COORD_FIELDS, new_m, after_m)
                if abs(want - got) > 1e-6]
    if mismatch:
        report.warn("Revit did not apply the requested value for: "
                    "<b>{}</b>".format(", ".join(mismatch)))
    if was_pinned:
        report.line("The base point was pinned - it was unpinned for the "
                    "move and re-pinned.")
    if not mismatch:
        report.success("Project Base Point moved to the new coordinates.")
    report.flush()


if __name__ == "__main__":
    main()
