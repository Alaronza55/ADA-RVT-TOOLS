# -*- coding: utf-8 -*-
__doc__ = """Fill a parameter of the model elements from a column of the
Excel planning (.xlsx / .xlsm) - e.g. the contractor doing each phase.

  1. pick the Excel file and the sheet,
  2. type the range of the phase names (e.g. B7:B52), the range of the
     column holding the values (same size, e.g. E7:E52 - any column,
     also one computed with formulas) and the parameter name,
  3. confirm.

The parameter is created if it does not exist yet: a text shared
parameter, instance, on every model category, in the PHASING group of
the properties. Then every element whose Phase Created has the same
name as a planning row gets that row's value. A phase listed on
several rows gets all its distinct values, separated by ", ".
Elements of phases that are not in the planning are left untouched.

Phase names are matched the same way as "Import Planning" creates them
(line breaks -> " - ", forbidden characters removed, case and spaces
ignored). Formulas give the value last calculated by Excel. The last
file, sheet, ranges and parameter name are remembered, and the Excel
file can stay open in Excel while it is read."""
__title__ = "Fill from\nPlanning"
__version__ = "Version 1.0"
__author__ = "ADA"

import os

import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import (
    BuiltInParameter, FilteredElementCollector, Transaction
)

from pyrevit import forms, revit, script

# Custom ADA GUI - themed buttons / inputs (lib/GUI), the shared dark/gold
# themed report (lib/GUI/ReportTheme.py) and the Planning helpers
# (lib/Snippets/_planning.py)
from GUI.forms import select_from_buttons, ask_for_inputs
from GUI.ReportTheme import ADAReport
from Snippets._planning import (bound_definition, clean_name, eid_value,
                                ensure_shared_text_parameter, name_key,
                                pick_sheet, pick_workbook)

doc = revit.doc

TITLE = __title__.replace("\n", " ")
CANCEL = "Cancel"

config = script.get_config()


def clean_value(text):
    """Cell text on one line (line breaks / repeated spaces -> one space)."""
    return u" ".join((text or u"").split())


def ask_inputs(workbook, sheet):
    """Ask ranges + parameter name until valid. Returns
    (rows [(phase, value)], phase_ref, value_ref, param_name)."""
    phase_ref = config.get_option("last_phase_range", "A2:A20")
    value_ref = config.get_option("last_value_range", "B2:B20")
    param_name = config.get_option("last_param_name", "Contractor")
    while True:
        values = ask_for_inputs(
            [("Phase names range", phase_ref),
             ("Values range (same size)", value_ref),
             ("Parameter name", param_name)],
            title=TITLE,
            label="Planning on sheet '{}' (e.g. B7:B52 and E7:E52):".format(
                sheet),
            button_name="Read Excel",
            version=__version__)
        if values is None:
            script.exit()
        phase_ref, value_ref = values[0].strip(), values[1].strip()
        param_name = clean_name(values[2])

        if not param_name:
            forms.alert("Type a parameter name.", title=TITLE)
            continue
        try:
            phases = workbook.read_range(sheet, phase_ref)
            cells = workbook.read_range(sheet, value_ref)
        except Exception as ex:
            forms.alert(str(ex), title=TITLE)
            continue
        if len(cells) != len(phases):
            forms.alert("The values range has {} cell(s) but the phase range "
                        "has {} - they must be the same size (one value per "
                        "phase row).".format(len(cells), len(phases)),
                        title=TITLE)
            continue

        rows = [(clean_name(p), clean_value(v)) for p, v in zip(phases, cells)
                if clean_name(p)]
        if not rows:
            forms.alert("No phase names found in {}.".format(phase_ref),
                        title=TITLE)
            continue
        return rows, phase_ref, value_ref, param_name


def merge_rows(rows):
    """[(phase, value)] in first-seen order, one entry per phase; distinct
    non-empty values of a phase listed several times are joined."""
    order, values, names = [], {}, {}
    for phase, value in rows:
        key = name_key(phase)
        if key not in values:
            order.append(key)
            values[key] = []
            names[key] = phase
        if value and value not in values[key]:
            values[key].append(value)
    return [(names[k], u", ".join(values[k])) for k in order]


def fill(param_name, value_by_phase_id):
    """Set `param_name` on every element whose Phase Created is a key of
    value_by_phase_id. Returns ({phase id: filled count}, failed count)."""
    filled, failed = {}, 0
    element_ids = list(FilteredElementCollector(doc)
                       .WhereElementIsNotElementType().ToElementIds())
    for el_id in element_ids:
        el = doc.GetElement(el_id)
        if el is None or not el.IsValidObject:
            continue
        created = el.get_Parameter(BuiltInParameter.PHASE_CREATED)
        if created is None:
            continue
        phase_id = eid_value(created.AsElementId())
        if phase_id not in value_by_phase_id:
            continue
        param = el.LookupParameter(param_name)
        if param is None:
            continue
        value = value_by_phase_id[phase_id]
        try:
            if param.AsString() == value:
                filled[phase_id] = filled.get(phase_id, 0) + 1
                continue
            if param.IsReadOnly or not param.Set(value):
                failed += 1
            else:
                filled[phase_id] = filled.get(phase_id, 0) + 1
        except Exception:
            failed += 1
    return filled, failed


def main():
    workbook = pick_workbook(config, TITLE)
    sheet = pick_sheet(workbook, TITLE, __version__)
    rows, phase_ref, value_ref, param_name = ask_inputs(workbook, sheet)
    planning = merge_rows(rows)

    config.last_file = workbook.path
    config.last_phase_range = phase_ref
    config.last_value_range = value_ref
    config.last_param_name = param_name
    script.save_config()

    by_key = {name_key(p.Name): p for p in doc.Phases}
    matched = [(phase, value, by_key[name_key(phase)])
               for phase, value in planning if name_key(phase) in by_key]
    if not matched:
        forms.alert("None of the {} planning phase(s) exists in this project "
                    "- run 'Import Planning' first.".format(len(planning)),
                    title=TITLE, exitscript=True)

    exists = bound_definition(doc, param_name) is not None
    action = u"{} '{}'".format("Fill" if exists else "Create & fill",
                               param_name)
    choice = select_from_buttons(
        [action, CANCEL],
        title=TITLE,
        label=(u"{} of {} planning phase(s) exist in the project.\nTheir "
               u"elements get the value of {} in '{}'.").format(
                   len(matched), len(planning), value_ref, param_name),
        version=__version__)
    if choice != action:
        script.exit()

    report = ADAReport(TITLE)
    report.line(u"Excel: <b>{}</b> &nbsp;|&nbsp; sheet <b>{}</b> &nbsp;|&nbsp; "
                u"phases {} &nbsp;|&nbsp; values {}".format(
                    os.path.basename(workbook.path), sheet, phase_ref,
                    value_ref))

    try:
        with Transaction(doc, u"ADA - Fill from Planning") as t:
            t.Start()
            created = ensure_shared_text_parameter(
                doc, param_name,
                u"Filled from the planning (ADA Fill from Planning)")
            filled, failed = fill(param_name, {eid_value(p.Id): v
                                               for _, v, p in matched})
            t.Commit()
    except Exception as ex:
        report.error(u"Failed, nothing was changed: {}".format(ex))
        report.flush()
        return

    if created:
        report.line(u"Shared parameter <b>{}</b> created (text, instance, all "
                    u"model categories, Phasing group).".format(param_name))

    table = []
    total = 0
    for phase, value in planning:
        p = by_key.get(name_key(phase))
        if p is None:
            table.append([phase, value or "-", "-", "Phase not in project"])
            continue
        count = filled.get(eid_value(p.Id), 0)
        total += count
        table.append([phase, value or "-", str(count),
                      "OK" if count else "No elements in this phase"])
    report.subheader(u"Planning -> '{}'".format(param_name))
    report.table(["Phase", "Value", "Elements", "Status"], table)

    if failed:
        report.warn(u"{} element(s) could not be changed - read-only "
                    u"parameter, element in a group or owned by another user, "
                    u"or '{}' is not a text parameter.".format(
                        failed, param_name))
    missing = len(planning) - len(matched)
    if missing:
        report.warn(u"{} planning phase(s) do not exist in the project - see "
                    u"'Import Planning'.".format(missing))
    report.success(u"'{}' set on {} element(s).".format(param_name, total))
    report.flush()


if __name__ == "__main__":
    main()
