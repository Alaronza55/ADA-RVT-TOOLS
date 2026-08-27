# -*- coding: utf-8 -*-
__doc__ = """Search worksets by name and open or close them.

Lists every user workset in the project in a searchable picker (type
to filter by name), lets you select one or more, then asks whether
to OPEN or CLOSE the selected workset(s).

This is a session-level state (whether the workset's elements are
loaded into this Revit session), not a per-view visibility toggle -
see "Workset Visibility" for that. Requires a workshared model.
"""
__title__ = "Open/Close\nWorkset"
__author__ = "ADA"

from pyrevit import revit, DB, forms, script

doc = revit.doc
output = script.get_output()

if not doc.IsWorkshared:
    forms.alert("This model is not workshared - there are no user worksets.",
                exitscript=True)

worksets = list(DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset))

if not worksets:
    forms.alert("No user worksets found in this model.", exitscript=True)

workset_by_name = {w.Name: w for w in worksets}

selected_names = forms.SelectFromList.show(
    sorted(workset_by_name.keys()),
    title="Search Worksets",
    button_name="Select",
    multiselect=True
)

if not selected_names:
    script.exit()

selected_worksets = [workset_by_name[name] for name in selected_names]

action = forms.CommandSwitchWindow.show(
    ["Open", "Close"],
    message="Open or close the selected workset(s)?"
)

if not action:
    script.exit()

open_it = (action == "Open")

results = []
workset_table = doc.GetWorksetTable()

with revit.Transaction("{} Worksets".format(action)):
    for w in selected_worksets:
        try:
            workset_table.SetWorksetOpen(w.Id, open_it)
            results.append((w.Name, "OK"))
        except Exception as ex:
            results.append((w.Name, "FAILED: {}".format(ex)))

output.print_md("### {} - {} workset(s)".format(action, len(selected_worksets)))
output.print_table(
    table_data=[[name, status] for name, status in results],
    columns=["Workset", "Result"]
)
