# -*- coding: utf-8 -*-
"""
Select one or more elements, either in the current model or inside a
linked model, and print the Element ID of each selected element.
"""
__title__ = "List\nElement IDs"
__author__ = "ADA"

from pyrevit import revit, UI
from pyrevit import forms

doc = revit.doc
uidoc = revit.uidoc

try:
    # Ask the user where to pick elements from
    source = forms.CommandSwitchWindow.show(
        ["Current Model", "Linked Model"],
        message="Select elements in:"
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

    print("=" * 70)
    print("SELECTED ELEMENT IDs")
    print("=" * 70)

    entries = []
    for display_id, link_name in picked_ids:
        if link_name:
            entries.append("{} (in link: {})".format(display_id, link_name))
        else:
            entries.append(str(display_id))

    print("; ".join(entries))

    print("-" * 70)
    print("Total elements selected: {}".format(len(picked_ids)))
    print("=" * 70)

except Exception as e:
    print("Error: {}".format(e))
    import traceback
    traceback.print_exc()
