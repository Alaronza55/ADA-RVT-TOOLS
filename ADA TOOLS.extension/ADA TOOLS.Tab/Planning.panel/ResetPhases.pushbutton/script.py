# -*- coding: utf-8 -*-
__doc__ = """Reset the project's phasing to a single phase, "PHASE0 - Genesis".

A Revit project always needs at least one phase (and the API cannot
create new ones), so the tool KEEPS one phase - the one already named
"PHASE0 - Genesis" if there is one, otherwise the earliest phase -
and:

  1. sets every element of the model to Phase Created = that phase
     and Phase Demolished = None,
  2. sets every view (and view template) to that phase,
  3. deletes all the other phases,
  4. renames the kept phase to "PHASE0 - Genesis".

You get a confirmation window first. Elements that cannot be changed
(e.g. inside a group, or owned by someone else in a workshared model)
are counted in the report. Rooms / Spaces / Areas belong to a phase
and cannot be moved to another one - those placed in a phase that is
deleted may be deleted with it (they are counted in the confirmation).
The whole change is one undoable transaction."""
__title__ = "Reset\nPhases"
__version__ = "Version 1.0"
__author__ = "ADA"

import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import (
    BuiltInParameter, ElementId, FailureProcessingResult, FailureSeverity,
    FilteredElementCollector, IFailuresPreprocessor, SpatialElement,
    SubTransaction, Transaction, View
)

from pyrevit import forms, revit, script

# Custom ADA GUI - button-choice popup (lib/GUI/SelectFromButtons.py) and
# the shared dark/gold themed report (lib/GUI/ReportTheme.py)
from GUI.forms import select_from_buttons
from GUI.ReportTheme import ADAReport

doc = revit.doc

TITLE = __title__.replace("\n", " ")
GENESIS_NAME = "PHASE0 - Genesis"

CONFIRM = "Reset all phases"
CANCEL = "Cancel"


class SwallowWarnings(IFailuresPreprocessor):
    """Delete warnings (not errors) so the reset does not stop on every
    'element created before its host' style message."""

    def PreprocessFailures(self, failures_accessor):
        for failure in failures_accessor.GetFailureMessages():
            if failure.GetSeverity() == FailureSeverity.Warning:
                failures_accessor.DeleteWarning(failure)
        return FailureProcessingResult.Continue


def get_phases():
    """Project phases, earliest first."""
    return [phase for phase in doc.Phases]


def set_param(element, bip, value):
    """Set a parameter if it exists and is editable. Returns True if the
    value was written, False if the element could not be changed, None if
    it has no such parameter / already has the value."""
    param = element.get_Parameter(bip)
    if param is None:
        return None
    if param.AsElementId() == value:
        return None
    if param.IsReadOnly:
        return False
    try:
        return bool(param.Set(value))
    except Exception:
        return False


def room_phase_id(room):
    param = room.get_Parameter(BuiltInParameter.ROOM_PHASE)
    return param.AsElementId() if param is not None else ElementId.InvalidElementId


def main():
    phases = get_phases()
    if not phases:
        forms.alert("This model has no phases.", title=TITLE, exitscript=True)

    keep = next((p for p in phases if p.Name == GENESIS_NAME), phases[0])
    to_delete = [p for p in phases if p.Id != keep.Id]

    spatial = list(FilteredElementCollector(doc).OfClass(SpatialElement))
    spatial_at_risk = [s for s in spatial
                       if room_phase_id(s) not in (keep.Id,
                                                   ElementId.InvalidElementId)]

    label = ("Keep '{}' as '{}', delete {} other phase(s) and set ALL elements "
             "and views to it.".format(keep.Name, GENESIS_NAME, len(to_delete)))
    if spatial_at_risk:
        label += ("\n{} Room(s)/Space(s)/Area(s) are in other phases and may "
                  "be deleted.".format(len(spatial_at_risk)))
    choice = select_from_buttons(
        [CONFIRM, CANCEL], title=TITLE, label=label, version=__version__)
    if choice != CONFIRM:
        script.exit()

    report = ADAReport(TITLE)
    report.line("Kept phase: <b>{}</b> &rarr; renamed <b>{}</b>".format(
        keep.Name, GENESIS_NAME))

    elements_changed = elements_failed = 0
    views_changed = views_failed = 0
    deleted, not_deleted = [], []
    demolished_cleared = 0

    t = Transaction(doc, "ADA - Reset Phases")
    options = t.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(SwallowWarnings())
    t.SetFailureHandlingOptions(options)
    t.Start()
    try:
        # 1. every element -> created in the kept phase, never demolished
        for el in FilteredElementCollector(doc).WhereElementIsNotElementType():
            created = set_param(el, BuiltInParameter.PHASE_CREATED, keep.Id)
            demolished = set_param(el, BuiltInParameter.PHASE_DEMOLISHED,
                                   ElementId.InvalidElementId)
            if created is False or demolished is False:
                elements_failed += 1
            elif created or demolished:
                elements_changed += 1
            if demolished:
                demolished_cleared += 1

        # 2. every view / view template -> kept phase
        for view in FilteredElementCollector(doc).OfClass(View):
            result = set_param(view, BuiltInParameter.VIEW_PHASE, keep.Id)
            if result is False:
                views_failed += 1
            elif result:
                views_changed += 1

        # 3. delete the other phases, one at a time
        for phase in to_delete:
            name = phase.Name
            sub = SubTransaction(doc)
            sub.Start()
            try:
                doc.Delete(phase.Id)
                sub.Commit()
                deleted.append(name)
            except Exception as ex:
                sub.RollBack()
                not_deleted.append((name, str(ex)))

        # 4. rename the kept phase (after the deletes, so the name is free)
        if keep.Name != GENESIS_NAME:
            keep.Name = GENESIS_NAME

        t.Commit()
    except Exception as ex:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        report.error("Failed, nothing was changed: {}".format(ex))
        report.flush()
        return

    report.subheader("Phases")
    rows = [[GENESIS_NAME, "Kept"]]
    rows += [[name, "Deleted"] for name in deleted]
    rows += [[name, "Could not be deleted - {}".format(err)]
             for name, err in not_deleted]
    report.table(["Phase", "Result"], rows)

    report.subheader("Elements & Views")
    report.table(
        ["", "Changed", "Could not be changed"],
        [["Elements (Phase Created / Demolished)",
          str(elements_changed), str(elements_failed)],
         ["Views & view templates (Phase)",
          str(views_changed), str(views_failed)]])
    if demolished_cleared:
        report.line("Phase Demolished cleared on <b>{}</b> element(s).".format(
            demolished_cleared))

    spatial_left = [s for s in FilteredElementCollector(doc)
                    .OfClass(SpatialElement)]
    lost = len(spatial) - len(spatial_left)
    if lost:
        report.warn("{} Room(s)/Space(s)/Area(s) were deleted with their "
                    "phase.".format(lost))
    if elements_failed:
        report.warn("{} element(s) could not be changed - typically members "
                    "of groups or elements owned by another user."
                    .format(elements_failed))
    if not_deleted:
        report.warn("{} phase(s) could not be deleted - see the table."
                    .format(len(not_deleted)))
    else:
        report.success("The project now has a single phase: {}.".format(
            GENESIS_NAME))
    report.flush()


if __name__ == "__main__":
    main()
