# -*- coding: utf-8 -*-
__doc__ = """Save the camera of the active 3D view and create one 3D view per
phase - for a 4D sequence of the planning.

  1. open the 3D view whose camera you want (orbit / zoom / section box),
  2. click the button, pick the phases (all by default with Select
     All) and the name prefix of the views.

Each view is a duplicate of the active view - same camera, perspective,
section box and visibility settings - with:
  - View Phase   = the phase it refers to,
  - Phase Filter = "Phasing 4D" (created if it does not exist: previous
    phases shown by category, elements of the view's phase shown
    overridden, demolished / temporary elements hidden).

The views are named "<prefix> - <phase>". Run it again after moving the
camera: views that already exist get the new camera / section box
instead of being duplicated. The new views are detached from any view
template (a template would force the same phase on all of them).
The whole change is one undoable transaction."""
__title__ = "4D Views\nper Phase"
__version__ = "Version 1.0"
__author__ = "ADA"

import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import (
    BuiltInParameter, ElementId, ElementOnPhaseStatus, FilteredElementCollector,
    PhaseFilter, PhaseStatusPresentation, Transaction, View3D,
    ViewDuplicateOption
)

from pyrevit import forms, revit, script

# Custom ADA GUI - themed list / inputs (lib/GUI), the shared dark/gold
# themed report (lib/GUI/ReportTheme.py)
from GUI.forms import select_from_dict, ask_for_inputs
from GUI.ReportTheme import ADAReport
from Snippets._planning import FORBIDDEN_CHARS

doc = revit.doc
source = doc.ActiveView

TITLE = __title__.replace("\n", " ")
FILTER_NAME = "Phasing 4D"

config = script.get_config()


def clean_view_name(text):
    return u" ".join(u"".join(ch for ch in text if ch not in FORBIDDEN_CHARS).split())


def get_or_create_phase_filter():
    """('Phasing 4D' filter, created?)"""
    for pf in FilteredElementCollector(doc).OfClass(PhaseFilter):
        if pf.Name == FILTER_NAME:
            return pf, False
    pf = PhaseFilter.Create(doc, FILTER_NAME)
    # the API spells it "ShowOverriden"
    overridden = (getattr(PhaseStatusPresentation, "ShowOverriden", None)
                  or getattr(PhaseStatusPresentation, "ShowOverridden"))
    settings = [(ElementOnPhaseStatus.Existing, PhaseStatusPresentation.ShowByCategory),
                (ElementOnPhaseStatus.New, overridden),
                (ElementOnPhaseStatus.Demolished, PhaseStatusPresentation.DontShow),
                (ElementOnPhaseStatus.Temporary, PhaseStatusPresentation.DontShow)]
    for status, presentation in settings:
        pf.SetPhaseStatusPresentation(status, presentation)
    return pf, True


def set_unique_name(view, name):
    """Name the view; add ' (2)', ' (3)'... if another view uses the name."""
    candidate, n = name, 1
    while True:
        try:
            view.Name = candidate
            return candidate
        except Exception:
            n += 1
            if n > 50:
                raise
            candidate = u"{} ({})".format(name, n)


def set_param(view, bip, value_id):
    param = view.get_Parameter(bip)
    if param is None or param.IsReadOnly:
        return False
    return bool(param.Set(value_id))


def copy_camera(target):
    """Camera + section box of the source view onto an existing view."""
    if target.IsLocked:
        return False
    target.SetOrientation(source.GetOrientation())
    target.IsSectionBoxActive = source.IsSectionBoxActive
    if source.IsSectionBoxActive:
        target.SetSectionBox(source.GetSectionBox())
    return True


def main():
    if not isinstance(source, View3D) or source.IsTemplate:
        forms.alert("Open the 3D view whose camera you want to use first.",
                    title=TITLE, exitscript=True)

    phases = list(doc.Phases)
    options = {u"{:02d}. {}".format(i + 1, p.Name): p for i, p in enumerate(phases)}
    chosen = select_from_dict(options, title=TITLE,
                              label="Phases to create a 3D view for:",
                              button_name="Next", version=__version__,
                              SelectMultiple=True)
    if not chosen:
        script.exit()
    chosen_ids = set(p.Id.ToString() for p in chosen)
    chosen = [p for p in phases if p.Id.ToString() in chosen_ids]   # project order

    values = ask_for_inputs([("View name prefix", config.get_option("last_prefix", "4D"))],
                            title=TITLE, label="Views are named '<prefix> - <phase>':",
                            button_name="Create views", version=__version__)
    if values is None:
        script.exit()
    prefix = clean_view_name(values[0]) or "4D"
    config.last_prefix = prefix
    script.save_config()

    existing = {v.Name: v for v in FilteredElementCollector(doc).OfClass(View3D)
                if not v.IsTemplate}

    rows = []
    with Transaction(doc, "ADA - 4D views per phase") as t:
        t.Start()
        phase_filter, filter_created = get_or_create_phase_filter()
        for phase in chosen:
            name = clean_view_name(u"{} - {}".format(prefix, phase.Name))
            view = existing.get(name)
            if view is not None and view.Id != source.Id:
                status = "Camera updated" if copy_camera(view) else \
                    "Exists - locked view, camera not changed"
            else:
                view = doc.GetElement(source.Duplicate(ViewDuplicateOption.Duplicate))
                name = set_unique_name(view, name)
                view.ViewTemplateId = ElementId.InvalidElementId
                status = "Created"
            ok_phase = set_param(view, BuiltInParameter.VIEW_PHASE, phase.Id)
            ok_filter = set_param(view, BuiltInParameter.VIEW_PHASE_FILTER,
                                  phase_filter.Id)
            if not (ok_phase and ok_filter):
                status += " - phase / filter not set (controlled by a view template?)"
            rows.append([name, phase.Name, status])
        t.Commit()

    report = ADAReport(TITLE)
    report.line(u"Camera of: <b>{}</b> &nbsp;|&nbsp; phase filter: <b>{}</b>{}".format(
        source.Name, FILTER_NAME, u" (created)" if filter_created else u""))
    report.table(["View", "Phase", "Result"], rows)
    created = sum(1 for r in rows if r[2].startswith("Created"))
    updated = sum(1 for r in rows if r[2].startswith("Camera updated"))
    report.success(u"{} view(s) created, {} updated.".format(created, updated))
    report.flush()


if __name__ == "__main__":
    main()
