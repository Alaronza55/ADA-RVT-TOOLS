# -*- coding: utf-8 -*-
__doc__ = """Delete all the Phase Filters of the project.

You get a confirmation window listing how many phase filters will be
deleted. Each filter is deleted on its own: if Revit refuses one (a
project may need to keep at least one phase filter), it is kept and
listed in the report with the reason, and the others are still
deleted. Views that used a deleted filter fall back to Revit's
default. The whole change is one undoable transaction."""
__title__ = "Delete Phase\nFilters"
__version__ = "Version 1.0"
__author__ = "ADA"

import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import (
    FilteredElementCollector, PhaseFilter, SubTransaction, Transaction
)

from pyrevit import forms, revit, script

# Custom ADA GUI - button-choice popup (lib/GUI/SelectFromButtons.py) and
# the shared dark/gold themed report (lib/GUI/ReportTheme.py)
from GUI.forms import select_from_buttons
from GUI.ReportTheme import ADAReport

doc = revit.doc

TITLE = __title__.replace("\n", " ")
CONFIRM = "Delete all phase filters"
CANCEL = "Cancel"


def main():
    filters = list(FilteredElementCollector(doc).OfClass(PhaseFilter))
    if not filters:
        forms.alert("This project has no phase filters.", title=TITLE,
                    exitscript=True)

    choice = select_from_buttons(
        [CONFIRM, CANCEL],
        title=TITLE,
        label="Delete the {} phase filter(s) of this project?".format(
            len(filters)),
        version=__version__)
    if choice != CONFIRM:
        script.exit()

    deleted, kept = [], []
    with Transaction(doc, "ADA - Delete Phase Filters") as t:
        t.Start()
        for phase_filter in filters:
            name = phase_filter.Name
            sub = SubTransaction(doc)
            sub.Start()
            try:
                doc.Delete(phase_filter.Id)
                sub.Commit()
                deleted.append(name)
            except Exception as ex:
                sub.RollBack()
                kept.append((name, str(ex)))
        t.Commit()

    report = ADAReport(TITLE)
    rows = [[name, "Deleted"] for name in deleted]
    rows += [[name, "Kept - {}".format(err)] for name, err in kept]
    report.subheader("Phase Filters")
    report.table(["Phase filter", "Result"], rows)
    if kept:
        report.warn("{} phase filter(s) could not be deleted - see the "
                    "table.".format(len(kept)))
    if deleted:
        report.success("{} phase filter(s) deleted.".format(len(deleted)))
    report.flush()


if __name__ == "__main__":
    main()
