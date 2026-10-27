# -*- coding: utf-8 -*-
__doc__ = """Copy the viewpoint of a 3D view to other 3D views.

  1. pick the reference 3D view in the list,
  2. tick the 3D views that should take the same viewpoint.

The camera (eye position, view direction and up direction) of the
reference view is applied to every selected view. Nothing else changes
(section box, crop, visibility, graphics...). Locked views are unlocked,
updated and locked again. A perspective view stays perspective and an
isometric view stays isometric (Revit cannot switch between the two), so
mixing them gives the same direction but a different framing.

One undoable transaction."""
__title__ = "Match\nViewpoint"
__version__ = "Version 1.0"
__author__ = "ADA"

import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import FilteredElementCollector, Transaction, View3D

from pyrevit import forms, revit, script

# Custom ADA GUI - themed list (lib/GUI) and the shared dark/gold themed
# report (lib/GUI/ReportTheme.py)
from GUI.forms import select_from_dict
from GUI.ReportTheme import ADAReport

doc = revit.doc

TITLE = __title__.replace("\n", " ")


def view_label(view):
    return u"{}{}{}".format(view.Name,
                            u"  (perspective)" if view.IsPerspective else u"",
                            u"  (locked)" if view.IsLocked else u"")


def pick_reference(views):
    options = {view_label(v): v for v in views}
    chosen = select_from_dict(options, title=TITLE,
                              label="Reference 3D view (viewpoint to copy):",
                              button_name="Next", version=__version__,
                              SelectMultiple=False)
    if not chosen:
        script.exit()
    return chosen[0]


def pick_targets(views, reference):
    options = {view_label(v): v for v in views if v.Id != reference.Id}
    if not options:
        forms.alert("There is no other 3D view.", title=TITLE, exitscript=True)
    chosen = select_from_dict(options, title=TITLE,
                              label=u"Apply the viewpoint of '{}' to:".format(reference.Name),
                              button_name="Apply", version=__version__,
                              SelectMultiple=True)
    if not chosen:
        script.exit()
    return chosen


def apply_viewpoint(reference, targets, report):
    orientation = reference.GetOrientation()
    done, failed = [], []
    with Transaction(doc, "ADA - Match 3D viewpoint") as tr:
        tr.Start()
        for view in targets:
            try:
                relock = view.IsLocked
                if relock:
                    view.Unlock()
                view.SetOrientation(orientation)
                if relock:
                    view.SaveOrientationAndLock()
                note = u""
                if view.IsPerspective != reference.IsPerspective:
                    note = u"perspective / isometric differ: same direction, different framing"
                elif relock:
                    note = u"re-locked"
                done.append([view.Name, note])
            except Exception as ex:
                failed.append([view.Name, u"{}".format(ex)])
        tr.Commit()

    if done:
        report.subheader("Updated views")
        report.table(["View", "Note"], done)
    if failed:
        report.subheader("Failed")
        report.table(["View", "Reason"], failed)
    if done:
        report.success(u"{} view(s) now share the viewpoint of '{}'.".format(
            len(done), reference.Name))
    else:
        report.error("No view was updated.")


views = sorted([v for v in FilteredElementCollector(doc).OfClass(View3D)
                if not v.IsTemplate],
               key=lambda v: v.Name.lower())
if len(views) < 2:
    forms.alert("This project needs at least two 3D views.", title=TITLE,
                exitscript=True)

reference = pick_reference(views)
targets = pick_targets(views, reference)

report = ADAReport(TITLE)
report.line(u"Reference view: <b>{}</b>{}".format(
    reference.Name, u" (perspective)" if reference.IsPerspective else u""))
try:
    apply_viewpoint(reference, targets, report)
except Exception as ex:
    report.error(u"Failed, nothing was changed: {}".format(ex))
report.flush()
