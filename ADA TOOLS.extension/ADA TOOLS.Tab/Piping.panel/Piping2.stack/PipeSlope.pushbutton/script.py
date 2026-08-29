# -*- coding: utf-8 -*-
"""Apply a slope percentage and a slope direction to selected pipes."""

__title__ = "Slope\nPipes"
__version__ = "Version 1.0"
__author__ = "ADA TOOLS"
__doc__ = ("Select one or more pipes, choose the direction the pipes should "
           "fall towards, enter a slope percentage and apply it.")

import math

from pyrevit import revit, DB, UI, forms, script

# Shared ADA-Tools dark/gold themed report and picker (see lib/GUI/ReportTheme.py,
# lib/GUI/SelectFromDict.py)
from GUI.ReportTheme import ADAReport
from GUI.forms import select_from_dict

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()
logger = script.get_logger()

TOL = 1.0e-6

# ---------------------------------------------------------------------------
# Options
# ---------------------------------------------------------------------------

OPT_TO_END = "Down towards the pipe END point"
OPT_TO_START = "Down towards the pipe START point"
OPT_NORTH = "Down towards project North (+Y)"
OPT_SOUTH = "Down towards project South (-Y)"
OPT_EAST = "Down towards project East (+X)"
OPT_WEST = "Down towards project West (-X)"
OPT_POINT = "Down towards a point I pick in the view"

DIRECTION_OPTIONS = [OPT_TO_END, OPT_TO_START, OPT_NORTH, OPT_SOUTH,
                     OPT_EAST, OPT_WEST, OPT_POINT]

CARDINAL = {
    OPT_NORTH: (0.0, 1.0),
    OPT_SOUTH: (0.0, -1.0),
    OPT_EAST: (1.0, 0.0),
    OPT_WEST: (-1.0, 0.0),
}

PIVOT_HIGH = "Keep the HIGH end at its current elevation (pipe drops)"
PIVOT_LOW = "Keep the LOW end at its current elevation (pipe rises)"
PIVOT_MID = "Keep the MID point at its current elevation (pipe pivots)"

PIVOT_OPTIONS = [PIVOT_HIGH, PIVOT_LOW, PIVOT_MID]


# ---------------------------------------------------------------------------
# Selection
# ---------------------------------------------------------------------------

class PipeSelectionFilter(UI.Selection.ISelectionFilter):
    """Only allow rigid pipes to be picked."""

    def AllowElement(self, element):
        return isinstance(element, DB.Plumbing.Pipe)

    def AllowReference(self, reference, point):
        return False


def collect_pipes():
    """Use the current selection if it holds pipes, otherwise ask the user."""
    selection = revit.get_selection()
    pipes = [el for el in selection.elements
             if isinstance(el, DB.Plumbing.Pipe)]
    if pipes:
        return pipes

    try:
        refs = uidoc.Selection.PickObjects(
            UI.Selection.ObjectType.Element,
            PipeSelectionFilter(),
            "Select the pipes to slope, then click Finish"
        )
    except Exception:
        return []
    return [doc.GetElement(ref.ElementId) for ref in refs]


# ---------------------------------------------------------------------------
# Geometry helpers
# ---------------------------------------------------------------------------

def to_mm(value):
    """Internal feet to millimetres, Revit 2021+ with a legacy fallback."""
    try:
        return DB.UnitUtils.ConvertFromInternalUnits(
            value, DB.UnitTypeId.Millimeters)
    except Exception:
        return DB.UnitUtils.ConvertFromInternalUnits(
            value, DB.DisplayUnitType.DUT_MILLIMETERS)


def resolve_low_end(p0, p1, direction, target):
    """Return (index of the endpoint that must go down, horizontal length).

    Returns (None, length) when the slope direction cannot be resolved,
    for example a vertical pipe or a pipe perpendicular to the direction.
    """
    dx = p1.X - p0.X
    dy = p1.Y - p0.Y
    hlen = math.sqrt(dx * dx + dy * dy)
    if hlen < TOL:
        return None, hlen

    if direction == OPT_TO_END:
        return 1, hlen
    if direction == OPT_TO_START:
        return 0, hlen

    if direction == OPT_POINT:
        d0 = math.sqrt((target.X - p0.X) ** 2 + (target.Y - p0.Y) ** 2)
        d1 = math.sqrt((target.X - p1.X) ** 2 + (target.Y - p1.Y) ** 2)
        if abs(d0 - d1) < TOL:
            return None, hlen
        return (1 if d1 < d0 else 0), hlen

    vx, vy = CARDINAL[direction]
    dot = dx * vx + dy * vy
    if abs(dot) < TOL:
        return None, hlen
    return (1 if dot > 0 else 0), hlen


def apply_slope(pipe, direction, slope, pivot, target):
    """Rebuild the pipe location curve with the requested slope.

    Returns (True, height_difference_in_feet) or (False, reason).
    """
    location = pipe.Location
    if not isinstance(location, DB.LocationCurve):
        return False, "no location curve"

    curve = location.Curve
    if not isinstance(curve, DB.Line):
        return False, "not a straight pipe"

    p0 = curve.GetEndPoint(0)
    p1 = curve.GetEndPoint(1)

    low_index, hlen = resolve_low_end(p0, p1, direction, target)
    if low_index is None:
        return False, "vertical pipe or perpendicular to slope direction"

    dz = hlen * slope

    points = [p0, p1]
    low = points[low_index]
    high = points[1 - low_index]

    if pivot == PIVOT_HIGH:
        new_high_z = high.Z
        new_low_z = high.Z - dz
    elif pivot == PIVOT_LOW:
        new_low_z = low.Z
        new_high_z = low.Z + dz
    else:
        mid_z = 0.5 * (p0.Z + p1.Z)
        new_high_z = mid_z + dz / 2.0
        new_low_z = mid_z - dz / 2.0

    new_points = [None, None]
    new_points[low_index] = DB.XYZ(low.X, low.Y, new_low_z)
    new_points[1 - low_index] = DB.XYZ(high.X, high.Y, new_high_z)

    if new_points[0].DistanceTo(new_points[1]) < \
            doc.Application.ShortCurveTolerance:
        return False, "resulting curve would be too short"

    location.Curve = DB.Line.CreateBound(new_points[0], new_points[1])
    return True, dz


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

pipes = collect_pipes()
if not pipes:
    forms.alert("No pipe selected. Nothing to do.", exitscript=True)

direction = select_from_dict(
    DIRECTION_OPTIONS,
    title=__title__,
    label="Slope direction - which way should the pipes fall?",
    button_name="Continue",
    version=__version__,
    SelectMultiple=False
)
if direction:
    direction = direction[0]
if not direction:
    script.exit()

target_point = None
if direction == OPT_POINT:
    try:
        target_point = uidoc.Selection.PickPoint(
            "Pick the point the pipes should fall towards")
    except Exception:
        script.exit()

raw_value = forms.ask_for_string(
    default="1.5",
    prompt="Slope in percent (for example 1.5 for 1.5%)",
    title="Slope value"
)
if raw_value is None:
    script.exit()

try:
    cleaned = raw_value.replace("%", "").replace(",", ".").strip()
    slope_value = abs(float(cleaned)) / 100.0
except ValueError:
    forms.alert("'{0}' is not a valid number.".format(raw_value),
                exitscript=True)

pivot = select_from_dict(
    PIVOT_OPTIONS,
    title=__title__,
    label="Which end keeps its current elevation?",
    button_name="Apply",
    version=__version__,
    SelectMultiple=False
)
if pivot:
    pivot = pivot[0]
if not pivot:
    script.exit()

results = []
tgroup = DB.TransactionGroup(doc, "Apply pipe slope")
tgroup.Start()

for pipe in pipes:
    trans = DB.Transaction(doc, "Slope pipe {0}".format(pipe.Id))
    trans.Start()
    try:
        success, info = apply_slope(pipe, direction, slope_value,
                                    pivot, target_point)
        if success:
            trans.Commit()
            results.append([output.linkify(pipe.Id), "Applied",
                            "{0:.1f} mm height difference".format(to_mm(info))])
        else:
            trans.RollBack()
            results.append([output.linkify(pipe.Id), "Skipped", info])
    except Exception as err:
        trans.RollBack()
        logger.debug("Pipe {0} failed: {1}".format(pipe.Id, err))
        results.append([output.linkify(pipe.Id), "Failed", str(err)])

tgroup.Assimilate()

applied = len([row for row in results if row[1] == "Applied"])
report = ADAReport(__title__.replace(chr(10), " "))
report.subheader("Pipe Slope: {0}% - {1}".format(round(slope_value * 100.0, 4), direction))
report.success("{0} of {1} pipes updated.".format(applied, len(pipes)))
report.table(["Pipe", "Result", "Details"], results)
report.flush()