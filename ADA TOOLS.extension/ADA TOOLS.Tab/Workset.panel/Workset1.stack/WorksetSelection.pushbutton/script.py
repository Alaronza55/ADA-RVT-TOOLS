# -*- coding: utf-8 -*-
__doc__ = """List all user worksets, pick one or more, choose whether to search
the entire model or just the active view, then select every element
belonging to the chosen workset(s). Requires a workshared model."""

__title__ = "Select by\nWorkset"
__version__ = "Version 1.0"
__author__ = "ADA Tools"

from pyrevit import revit, DB, forms, script
from System.Collections.Generic import List

# Shared ADA-Tools dark/gold themed report and pickers (see lib/GUI/ReportTheme.py,
# lib/GUI/SelectFromDict.py, lib/GUI/SelectFromButtons.py)
from GUI.ReportTheme import ADAReport
from GUI.forms import select_from_dict, select_from_buttons

doc = revit.doc
uidoc = revit.uidoc

# --- 1. Model must be workshared -------------------------------------------
if not doc.IsWorkshared:
    forms.alert("This model is not workshared - there are no user worksets.",
                exitscript=True)

# --- 2. Collect user worksets ----------------------------------------------
worksets = DB.FilteredWorksetCollector(doc) \
             .OfKind(DB.WorksetKind.UserWorkset) \
             .ToWorksets()

if not worksets:
    forms.alert("No user worksets found in this model.", exitscript=True)

ws_map = {}
for ws in worksets:
    label = ws.Name if ws.IsOpen else "{} (closed)".format(ws.Name)
    ws_map[label] = ws

# --- 3. Ask the user which workset(s) --------------------------------------
selected = select_from_dict(
    sorted(ws_map.keys()),
    title=__title__,
    label="Select workset(s):",
    button_name="Select Elements",
    version=__version__,
    SelectMultiple=True
)

if not selected:
    script.exit()

if isinstance(selected, str):
    selected = [selected]

# --- 4. Ask for the search scope -------------------------------------------
scope = select_from_buttons(
    ["Entire Model", "Active View Only"],
    title=__title__,
    label="Search scope:",
    version=__version__
)

if not scope:
    script.exit()

# --- 5. Build the workset filter -------------------------------------------
filters = List[DB.ElementFilter]()
for label in selected:
    filters.Add(DB.ElementWorksetFilter(ws_map[label].Id))

ws_filter = filters[0] if filters.Count == 1 else DB.LogicalOrFilter(filters)

# --- 6. Collect and select --------------------------------------------------
if scope == "Active View Only":
    collector = DB.FilteredElementCollector(doc, doc.ActiveView.Id)
else:
    collector = DB.FilteredElementCollector(doc)

element_ids = collector.WherePasses(ws_filter) \
                       .WhereElementIsNotElementType() \
                       .ToElementIds()

if element_ids.Count == 0:
    forms.alert("No elements found in: {}".format(", ".join(selected)))
    script.exit()

uidoc.Selection.SetElementIds(element_ids)

report = ADAReport(__title__.replace(chr(10), " "))
report.success("<b>{}</b> element(s) selected.".format(element_ids.Count))
report.line("Scope: <b>{}</b>".format(scope))
report.table(["Workset"], [[label] for label in selected])
report.flush()

forms.alert("{} element(s) selected in:\n\n{}".format(
    element_ids.Count, "\n".join(selected)))

