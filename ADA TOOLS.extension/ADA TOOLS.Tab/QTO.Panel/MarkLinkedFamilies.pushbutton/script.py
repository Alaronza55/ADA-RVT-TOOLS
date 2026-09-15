# -*- coding: utf-8 -*-
__doc__ = """Find the family instances of a selected linked model that are
VISIBLE IN THE ACTIVE VIEW and whose FAMILY NAME contains a user-typed
piece of text, then mark each one with a 2D circle (detail lines) drawn
on the active view.

Workflow:
  1. Pick which loaded link to search in.
  2. Type the text to look for in family names (case-insensitive
     "contains" match) and the circle diameter in millimeters - both
     asked in a single window.

Visibility: on Revit 2024+ the script asks Revit itself which linked
elements the active view shows, via the FilteredElementCollector
overload that takes (host doc, view id, link instance id) - so view
range, crop region, hidden categories and view filters are all
respected exactly. On older Revit versions, where that overload does
not exist, it falls back to a geometric test: the element's bounding
box (taken through the link's placement transform into this model's
coordinates) is checked against the plan view's own view range (read
from GetViewRange - NOT from the crop box, whose Z extents are
meaningless for plan views) and against the crop region when active;
for sections/elevations the oriented crop box is used directly.

For each match the script takes its center point - the center of its
bounding box, falling back to its location point - projects it onto the
view's plane, and draws a circle of the requested diameter there as two
semicircular detail arcs (Revit does not accept a closed curve as a
single detail line).

The circles are real detail lines owned by the current model, so they
print, they can be selected/deleted like any other detail line, and
each one is listed with a clickable link in the report."""
__title__ = "Mark Linked\nFamilies"
__version__ = "Version 1.0"
__author__ = "ADA"

import math

from pyrevit import revit, DB
from pyrevit import forms

# Custom ADA GUI - themed list picker and text-input popup (see
# lib/GUI/SelectFromDict.py and lib/GUI/AskForInputs.py) and the shared
# dark/gold themed report (see lib/GUI/ReportTheme.py)
from GUI.forms import select_from_dict, ask_for_inputs
from GUI.ReportTheme import ADAReport

doc = revit.doc
uidoc = revit.uidoc

TITLE = __title__.replace("\n", " ")

MM_PER_FOOT = 304.8
CIRCLE_COLOR = DB.Color(255, 140, 0)  # orange, same accent as the QTO markers
CIRCLE_LINE_WEIGHT = 5
TOLERANCE = 0.01  # feet, slack for the geometric visibility tests

PLAN_VIEW_TYPES = (
    DB.ViewType.FloorPlan,
    DB.ViewType.CeilingPlan,
    DB.ViewType.EngineeringPlan,
    DB.ViewType.AreaPlan,
)

# View types whose plane detail curves can be drawn on (and that show
# model elements, so "visible in the view" means something)
DETAIL_CURVE_VIEW_TYPES = PLAN_VIEW_TYPES + (
    DB.ViewType.Elevation,
    DB.ViewType.Section,
    DB.ViewType.Detail,
)


def pick_link_instance():
    """Let the user choose one loaded RevitLinkInstance. Titles are
    disambiguated with the instance id when the same link is placed
    more than once."""
    instances = [
        inst for inst in DB.FilteredElementCollector(doc)
        .OfClass(DB.RevitLinkInstance)
        .ToElements()
        if inst.GetLinkDocument() is not None
    ]
    if not instances:
        forms.alert("No loaded Revit links found in this model.", exitscript=True)

    titles = [inst.GetLinkDocument().Title for inst in instances]
    name_map = {}
    for inst, title in zip(instances, titles):
        if titles.count(title) > 1:
            title = "{} [{}]".format(title, inst.Id.IntegerValue)
        name_map[title] = inst

    chosen = select_from_dict(
        name_map,
        title=TITLE,
        label="Select Linked Model:",
        button_name="Select",
        version=__version__,
        SelectMultiple=False,
    )
    if not chosen:
        forms.alert("Cancelled.", exitscript=True)
    return chosen[0]


def ask_search_and_diameter():
    """One themed window asking both the family-name text and the
    circle diameter; loops until the diameter parses as a positive
    number. Returns (search_text, diameter_mm)."""
    diameter_default = "500"
    while True:
        values = ask_for_inputs(
            [("Family name contains:", ""),
             ("Circle diameter (mm):", diameter_default)],
            title=TITLE,
            label="Mark matching families with a circle:",
            button_name="Mark",
            version=__version__,
        )
        if values is None:
            forms.alert("Cancelled.", exitscript=True)

        search_text = values[0].strip()
        raw_diameter = values[1].strip()
        diameter_default = raw_diameter  # keep what they typed on retry

        if not search_text:
            forms.alert("Please type the text to search for in family names.")
            continue
        try:
            diameter_mm = float(raw_diameter.replace(",", "."))
        except ValueError:
            forms.alert("'{}' is not a number.".format(raw_diameter))
            continue
        if diameter_mm <= 0:
            forms.alert("Diameter must be greater than zero.")
            continue
        return search_text, diameter_mm


def element_center(element):
    """Center of the element's bounding box (in its own document's
    coordinates), falling back to its location point."""
    bbox = element.get_BoundingBox(None)
    if bbox is not None:
        return DB.XYZ(
            (bbox.Min.X + bbox.Max.X) / 2.0,
            (bbox.Min.Y + bbox.Max.Y) / 2.0,
            (bbox.Min.Z + bbox.Max.Z) / 2.0,
        )
    location = element.Location
    if isinstance(location, DB.LocationPoint):
        return location.Point
    return None


# --- Visibility: primary path (Revit 2024+) -------------------------------

def collect_visible_via_revit(view, link_instance):
    """Ask Revit which linked family instances the view actually shows,
    via the FilteredElementCollector(doc, viewId, linkInstanceId)
    overload added in Revit 2024. Returns a list, or None when the
    overload does not exist (older Revit)."""
    try:
        collector = DB.FilteredElementCollector(doc, view.Id, link_instance.Id)
    except (TypeError, Exception):
        return None
    try:
        return list(collector
                    .OfClass(DB.FamilyInstance)
                    .WhereElementIsNotElementType())
    except Exception:
        return None


# --- Visibility: geometric fallback (Revit < 2024) ------------------------

def plan_view_z_range(view):
    """(z_bottom, z_top) shown by a plan view, in host model
    coordinates, read from the view range: top clip plane down to the
    view depth plane. Either side is None when unlimited/unresolved."""
    vr = view.GetViewRange()

    def plane_z(plane):
        try:
            level_id = vr.GetLevelId(plane)
        except Exception:
            return None
        if level_id == DB.ElementId.InvalidElementId:
            return None
        unlimited = getattr(DB.PlanViewRange, 'Unlimited', None)
        if unlimited is not None and level_id == unlimited:
            return None
        current = getattr(DB.PlanViewRange, 'Current', None)
        if current is not None and level_id == current:
            level = view.GenLevel
        else:
            level = doc.GetElement(level_id)
        if not isinstance(level, DB.Level):
            return None
        return level.ProjectElevation + vr.GetOffset(plane)

    return plane_z(DB.PlanViewPlane.ViewDepthPlane), \
        plane_z(DB.PlanViewPlane.TopClipPlane)


def host_corners(element, link_transform):
    """The element's bounding-box corners taken into host coordinates
    through the link's placement transform (or just its center when it
    has no bounding box). None when neither exists."""
    bbox = element.get_BoundingBox(None)
    if bbox is None:
        center = element_center(element)
        if center is None:
            return None
        corners = [center]
    else:
        corners = [
            DB.XYZ(x, y, z)
            for x in (bbox.Min.X, bbox.Max.X)
            for y in (bbox.Min.Y, bbox.Max.Y)
            for z in (bbox.Min.Z, bbox.Max.Z)
        ]
    return [link_transform.OfPoint(p) for p in corners]


def crop_overlap(points_host, view, axes):
    """Overlap test between the points' extents and the view's crop
    box, in the crop box's own oriented space, along the given axes
    ('x', 'y', 'z')."""
    crop = view.CropBox
    to_local = crop.Transform.Inverse
    local = [to_local.OfPoint(p) for p in points_host]
    getters = {'x': lambda p: p.X, 'y': lambda p: p.Y, 'z': lambda p: p.Z}
    for axis in axes:
        get = getters[axis]
        values = [get(p) for p in local]
        if max(values) < get(crop.Min) - TOLERANCE:
            return False
        if min(values) > get(crop.Max) + TOLERANCE:
            return False
    return True


def is_visible_geometric(element, link_transform, view, plan_z):
    """Fallback visibility test for Revit versions without the linked
    visible-in-view collector. Plan views: world-Z against the view
    range, XY against the crop region when active. Other views: the
    oriented crop box (whose extents ARE meaningful there)."""
    corners = host_corners(element, link_transform)
    if corners is None:
        return False

    if plan_z is not None:
        z_bottom, z_top = plan_z
        zs = [p.Z for p in corners]
        if z_top is not None and min(zs) > z_top + TOLERANCE:
            return False
        if z_bottom is not None and max(zs) < z_bottom - TOLERANCE:
            return False
        if view.CropBoxActive and not crop_overlap(corners, view, 'xy'):
            return False
        return True

    axes = 'xyz' if view.CropBoxActive else 'z'
    return crop_overlap(corners, view, axes)


# --- Drawing ---------------------------------------------------------------

def project_onto_view_plane(point, view):
    """Drop the point onto the view's plane - detail curves must lie
    exactly in it."""
    normal = view.ViewDirection
    depth = point.Subtract(view.Origin).DotProduct(normal)
    return point.Subtract(normal.Multiply(depth))


def draw_circle(view, center, radius):
    """Draw a circle as two semicircular detail arcs in the view's own
    right/up axes; returns the created DetailCurve elements."""
    x_axis = view.RightDirection
    y_axis = view.UpDirection
    curves = []
    for start, end in ((0.0, math.pi), (math.pi, 2.0 * math.pi)):
        arc = DB.Arc.Create(center, radius, start, end, x_axis, y_axis)
        curves.append(doc.Create.NewDetailCurve(view, arc))
    return curves


try:
    view = doc.ActiveView
    if view.ViewType not in DETAIL_CURVE_VIEW_TYPES:
        forms.alert(
            "The active view ({}) cannot host detail lines.\n"
            "Open a plan, section, elevation or detail view "
            "and run the tool again.".format(view.ViewType),
            exitscript=True)

    link_instance = pick_link_instance()
    link_doc = link_instance.GetLinkDocument()
    search_text, diameter_mm = ask_search_and_diameter()

    radius_ft = (diameter_mm / 2.0) / MM_PER_FOOT
    transform = link_instance.GetTotalTransform()
    search_lower = search_text.lower()

    def family_name_of(fi):
        try:
            return fi.Symbol.Family.Name
        except Exception:
            return None

    report = ADAReport(TITLE)
    report.line("Link: <b>{}</b>".format(link_doc.Title))
    report.line("View: <b>{}</b>".format(view.Name))
    report.line("Family name contains: <b>{}</b>".format(search_text))
    report.line("Circle diameter: <b>{:.0f} mm</b>".format(diameter_mm))

    if link_instance.IsHidden(view):
        report.warn("The selected link is hidden in the active view - "
                    "nothing to mark.")
        report.flush()
    else:
        # How many instances in the WHOLE link match the text (for the
        # report), regardless of visibility
        name_matches = 0
        for fi in DB.FilteredElementCollector(link_doc)\
                .OfClass(DB.FamilyInstance)\
                .WhereElementIsNotElementType():
            name = family_name_of(fi)
            if name and search_lower in name.lower():
                name_matches += 1

        # Visible instances: let Revit decide (2024+), else geometric test
        visible = collect_visible_via_revit(view, link_instance)
        if visible is not None:
            visibility_method = "Revit's own visible-in-view collector"
        else:
            visibility_method = "geometric (view range / crop region)"
            plan_z = plan_view_z_range(view) \
                if view.ViewType in PLAN_VIEW_TYPES else None
            visible = [
                fi for fi in DB.FilteredElementCollector(link_doc)
                .OfClass(DB.FamilyInstance)
                .WhereElementIsNotElementType()
                if is_visible_geometric(fi, transform, view, plan_z)
            ]

        matches = []
        for fi in visible:
            name = family_name_of(fi)
            if name and search_lower in name.lower():
                matches.append((fi, name))

        report.line("Visibility check: {}".format(visibility_method))

        if not matches:
            if name_matches:
                report.warn(
                    "{} instance(s) in the link match the text, but none "
                    "is visible in the active view.".format(name_matches))
            else:
                report.warn("No family instance in the link matches that text.")
            report.flush()
        else:
            table_rows = []
            skipped = 0
            drawn = 0

            ogs = DB.OverrideGraphicSettings()
            ogs.SetProjectionLineColor(CIRCLE_COLOR)
            ogs.SetProjectionLineWeight(CIRCLE_LINE_WEIGHT)

            with revit.Transaction("Mark Linked Families"):
                for fi, family_name in matches:
                    center_link = element_center(fi)
                    if center_link is None:
                        skipped += 1
                        report.warn(
                            "Element ID <b>{}</b> ({}): no bounding box or "
                            "location point, skipped.".format(
                                fi.Id.IntegerValue, family_name))
                        continue

                    center_host = transform.OfPoint(center_link)
                    center_view = project_onto_view_plane(center_host, view)

                    try:
                        curves = draw_circle(view, center_view, radius_ft)
                    except Exception as draw_err:
                        skipped += 1
                        report.warn(
                            "Element ID <b>{}</b> ({}): could not draw circle "
                            "({}).".format(fi.Id.IntegerValue, family_name, draw_err))
                        continue

                    for curve in curves:
                        view.SetElementOverrides(curve.Id, ogs)

                    drawn += 1
                    table_rows.append([
                        str(fi.Id.IntegerValue),
                        family_name,
                        fi.Name,
                        report.link(curves[0].Id, title="Show circle"),
                    ])

            uidoc.RefreshActiveView()

            if table_rows:
                report.subheader("Marked Elements")
                report.table(
                    ["Link Element ID", "Family", "Type", "Circle"],
                    table_rows)

            report.subheader("Summary")
            report.line("Instances matching the text in the whole link: "
                        "<b>{}</b>".format(name_matches))
            report.line("Visible in the active view: <b>{}</b>".format(len(matches)))
            if skipped:
                report.warn("{} instance(s) skipped - see warnings above.".format(skipped))
            report.success(
                "{} circle(s) of {:.0f} mm drawn on view '{}'.".format(
                    drawn, diameter_mm, view.Name))
            report.flush()

except Exception as e:
    if 'cancel' not in str(e).lower():
        report = report if 'report' in globals() else ADAReport(TITLE)
        report.error("Error: {}".format(e))
        report.flush()
        import traceback
        print(traceback.format_exc())
