# -*- coding: utf-8 -*-
__doc__ = """Select one or more elements, either in the current model or inside a
linked model, and print the Element ID of each selected element."""
__title__ = "List\nElement IDs"
__version__ = "Version 1.0"
__author__ = "ADA"

from pyrevit import revit, UI
from pyrevit import forms

# Shared ADA-Tools dark/gold themed report (see lib/GUI/ReportTheme.py) and
# small button-choice popup (see lib/GUI/SelectFromButtons.py)
from GUI.ReportTheme import ADAReport
from GUI.forms import select_from_buttons

doc = revit.doc
uidoc = revit.uidoc

try:
    # Ask the user where to pick elements from
    source = select_from_buttons(
        ["Current Model", "Linked Model"],
        title=__title__,
        label="Select elements in:",
        version=__version__
    )

    if not source:
        forms.alert("Cancelled.", exitscript=True)

    # Each entry is (display_id, label)
    picked_ids = []

    if source == "Current Model":
        # Reuse existing selection if there is one, otherwise prompt
        selection = uidoc.Selection
        selected_ids = selection.GetElementIds()

        if not selected_ids or selected_ids.Count == 0:
            refs = uidoc.Selection.PickObjects(
                UI.Selection.ObjectType.Element,
                "Select elements to list their IDs"
            )
            selected_ids = [ref.ElementId for ref in refs]

        for elem_id in selected_ids:
            picked_ids.append((elem_id.IntegerValue, None))

    else:  # Linked Model
        refs = uidoc.Selection.PickObjects(
            UI.Selection.ObjectType.LinkedElement,
            "Select elements in the linked model to list their IDs"
        )

        for ref in refs:
            link_instance = doc.GetElement(ref.ElementId)
            link_name = link_instance.Name if link_instance else "Unknown Link"
            picked_ids.append((ref.LinkedElementId.IntegerValue, link_name))

    if not picked_ids:
        forms.alert("No elements selected.", exitscript=True)

    report = ADAReport(__title__.replace(chr(10), " "))

    table_rows = [
        [str(display_id), link_name if link_name else "current model"]
        for display_id, link_name in picked_ids
    ]
    report.table(["Element ID", "Location"], table_rows)

    report.subheader("Summary")
    report.line("Total elements selected: <b>{}</b>".format(len(picked_ids)))
    report.flush()

except Exception as e:
    report = ADAReport(__title__.replace(chr(10), " "))
    report.error("Error: {}".format(e))
    report.flush()
    import traceback
    print(traceback.format_exc())
