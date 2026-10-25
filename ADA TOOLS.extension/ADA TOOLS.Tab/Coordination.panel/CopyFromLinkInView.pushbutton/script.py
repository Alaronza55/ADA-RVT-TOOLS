# -*- coding: utf-8 -*-
__doc__ = """Copy the elements of a Revit link that are VISIBLE IN THE ACTIVE
VIEW into the active project, at the same place.

  1. pick the link,
  2. pick the categories to copy (the list shows how many elements of
     each category are visible),
  3. the elements are copied with Revit's own copy between documents
     (like copy / paste), through the link's placement - hosted
     elements keep their hosts when both are copied.

"Visible in the active view" is decided by Revit itself: crop region,
view range, section box, hidden elements / categories, filters...
Only model elements are copied (no annotations, views, levels, grids).
Types with the same name as in the project use the project's type.
Elements that cannot be copied are skipped and counted in the report.
The whole copy is one undoable step. Requires Revit 2024 or newer."""
__title__ = "Copy from Link\n(Active View)"
__version__ = "Version 1.0"
__author__ = "ADA"

import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import (
    BasePoint, CategoryType, CopyPasteOptions, DuplicateTypeAction,
    ElementId, ElementTransformUtils, FailureProcessingResult,
    FailureSeverity, FilteredElementCollector, IDuplicateTypeNamesHandler,
    IFailuresPreprocessor, ImportInstance, RevitLinkInstance, Transaction,
    TransactionGroup, TransactionStatus, View, ViewSheet, ViewSchedule
)
from System.Collections.Generic import List

from pyrevit import forms, revit, script

# Custom ADA GUI - themed list (lib/GUI/SelectFromDict.py) and the shared
# dark/gold themed report (lib/GUI/ReportTheme.py)
from GUI.forms import select_from_dict
from GUI.ReportTheme import ADAReport

doc = revit.doc
view = doc.ActiveView

TITLE = __title__.replace("\n", " ")
CHUNK = 50          # elements per copy call when a category has to be split


class UseDestinationTypes(IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DuplicateTypeAction.UseDestinationTypes


class SwallowWarnings(IFailuresPreprocessor):
    """Warnings (duplicate instances, overlaps...) must not stop the copy."""

    def PreprocessFailures(self, failures_accessor):
        for failure in failures_accessor.GetFailureMessages():
            if failure.GetSeverity() == FailureSeverity.Warning:
                failures_accessor.DeleteWarning(failure)
        return FailureProcessingResult.Continue


def eid_value(eid):
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue


def pick_link():
    links = [l for l in FilteredElementCollector(doc).OfClass(RevitLinkInstance)
             if l.GetLinkDocument() is not None]
    if not links:
        forms.alert("There is no loaded Revit link in this model.",
                    title=TITLE, exitscript=True)
    options = {"{}  [{}]".format(l.Name, eid_value(l.Id)): l for l in links}
    chosen = select_from_dict(options, title=TITLE,
                              label="Link to copy from (elements visible in "
                                    "'{}'):".format(view.Name),
                              button_name="Next", version=__version__,
                              SelectMultiple=False)
    if not chosen:
        script.exit()
    return chosen[0]


def visible_link_elements(link):
    """{category name: [ElementId]} of the link's MODEL elements visible in
    the active view (Revit 2024+ collector)."""
    try:
        collector = FilteredElementCollector(doc, view.Id, link.Id)
    except Exception:
        forms.alert("This tool needs Revit 2024 or newer (collector of link "
                    "elements visible in a view).", title=TITLE,
                    exitscript=True)
    by_category = {}
    for el in collector.WhereElementIsNotElementType():
        cat = el.Category
        if cat is None or cat.CategoryType != CategoryType.Model:
            continue
        if isinstance(el, (RevitLinkInstance, ImportInstance, BasePoint, View)):
            continue
        by_category.setdefault(cat.Name, []).append(el.Id)
    return by_category


def copy_ids(link_doc, ids, transform, options):
    """Copy ids in ONE call (so hosted elements travel with their hosts); if
    Revit refuses the batch, split it - into chunks, then halves, down to
    one element. Returns (new element ids, failed source ids)."""
    if not ids:
        return [], []
    t = Transaction(doc, "ADA - copy from link")
    fho = t.GetFailureHandlingOptions()
    fho.SetFailuresPreprocessor(SwallowWarnings())
    t.SetFailureHandlingOptions(fho)
    t.Start()
    try:
        new_ids = ElementTransformUtils.CopyElements(
            link_doc, List[ElementId](ids), doc, transform, options)
        if t.Commit() == TransactionStatus.Committed:
            return list(new_ids), []
    except Exception:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
    if len(ids) == 1:
        return [], list(ids)
    size = CHUNK if len(ids) > CHUNK else max(1, len(ids) // 2)
    copied, failed = [], []
    for i in range(0, len(ids), size):
        c, f = copy_ids(link_doc, ids[i:i + size], transform, options)
        copied += c
        failed += f
    return copied, failed


def count_by_category(document, ids):
    counts = {}
    for eid in ids:
        el = document.GetElement(eid)
        if el is not None and el.Category is not None:
            counts[el.Category.Name] = counts.get(el.Category.Name, 0) + 1
    return counts


def main():
    if isinstance(view, (ViewSheet, ViewSchedule)) or not view.CanBePrinted:
        forms.alert("Open a model view (plan, section, 3D...) first.",
                    title=TITLE, exitscript=True)

    link = pick_link()
    link_doc = link.GetLinkDocument()
    by_category = visible_link_elements(link)
    if not by_category:
        forms.alert("No model element of '{}' is visible in this view."
                    .format(link.Name), title=TITLE, exitscript=True)

    options = {"{}  ({})".format(name, len(ids)): name
               for name, ids in by_category.items()}
    chosen = select_from_dict(options, title=TITLE,
                              label="Categories to copy (visible elements):",
                              button_name="Copy", version=__version__,
                              SelectMultiple=True)
    if not chosen:
        script.exit()

    paste = CopyPasteOptions()
    paste.SetDuplicateTypeNamesHandler(UseDestinationTypes())
    transform = link.GetTotalTransform()

    ids = [eid for name in chosen for eid in by_category[name]]
    group = TransactionGroup(doc, "ADA - Copy from Link (active view)")
    group.Start()
    try:
        new_ids, failed_ids = copy_ids(link_doc, ids, transform, paste)
        group.Assimilate()
    except Exception as ex:
        group.RollBack()
        report = ADAReport(TITLE)
        report.error("Failed, nothing was copied: {}".format(ex))
        report.flush()
        return

    copied_by_cat = count_by_category(doc, new_ids)
    failed_by_cat = count_by_category(link_doc, failed_ids)
    rows = [[name, str(len(by_category[name])),
             str(copied_by_cat.get(name, 0)), str(failed_by_cat.get(name, 0))]
            for name in sorted(chosen)]
    # elements Revit adds on its own (e.g. sketches) appear under their category
    total_copied = len(new_ids)
    total_failed = len(failed_ids)

    report = ADAReport(TITLE)
    report.line(u"Link: <b>{}</b> &nbsp;|&nbsp; view: <b>{}</b>".format(
        link.Name, view.Name))
    report.table(["Category", "Visible", "Copied", "Failed"], rows)
    if total_failed:
        report.warn("{} element(s) could not be copied (e.g. hosted elements "
                    "whose host was not copied, or elements Revit cannot copy "
                    "between documents).".format(total_failed))
    report.success("{} element(s) copied into the project.".format(total_copied))
    report.flush()


if __name__ == "__main__":
    main()
