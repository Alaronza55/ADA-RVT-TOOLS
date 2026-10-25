# -*- coding: utf-8 -*-
__doc__ = """Import a planning from Excel (.xlsx / .xlsm): phase names and,
optionally, the contractor doing each phase.

  1. pick the Excel file and the sheet,
  2. type the range of the phase names (e.g. A2:A20) and, optionally,
     the range of the contractors (same size, e.g. B2:B20),
  3. check the list of phases read from Excel (you can reorder it),
  4. choose whether to create / fill the shared parameter
     "Contractor" (Phasing group, on every model category).

PHASES: the Revit API cannot create (or rename) phases. Phases of the
planning that already exist in the project (same name) are used as
they are. The missing ones can be CREATED AUTOMATICALLY: the tool
opens Manage > Phases and a helper drives the dialog (Insert After +
types each name, in planning order, after the selected phase). Do not
touch the mouse / keyboard until the ADA message appears, then check
the list and click OK (Cancel discards everything). Run Import
Planning again to fill the Contractor parameter for the new phases -
the last file, sheet and ranges are remembered.

Phase names are cleaned: line breaks become " - " and characters
Revit does not accept in names ( \\ : { } [ ] | ; < > ? ` ~ ) are
removed. SHIFT+CLICK the button to only record the Phasing dialog's
structure (diagnostic log, if the automation does not work).

CONTRACTOR: every element whose Phase Created is one of the planning
phases gets that phase's contractor in its "Contractor" parameter.
Elements of phases that are not in the planning are left untouched.
The Excel file can stay open in Excel while it is read."""
__title__ = "Import\nPlanning"
__version__ = "Version 1.0"
__author__ = "ADA"

import os

import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')

from Autodesk.Revit.DB import (
    BuiltInParameter, CategoryType, ExternalDefinitionCreationOptions,
    FilteredElementCollector, Transaction
)
from Autodesk.Revit.UI import PostableCommand, RevitCommandId
from System.Diagnostics import Process, ProcessStartInfo, ProcessWindowStyle
from System.Windows.Forms import Clipboard

from pyrevit import forms, revit, script

# Custom ADA GUI - themed list / buttons / inputs / order list
# (lib/GUI) and the shared dark/gold themed report (lib/GUI/ReportTheme.py)
from GUI.forms import (select_from_dict, select_from_buttons, ask_for_inputs,
                       order_list)
from GUI.ReportTheme import ADAReport
from Snippets._xlsx import XlsxWorkbook

doc = revit.doc
app = doc.Application
uiapp = __revit__

TITLE = __title__.replace("\n", " ")

PARAM_NAME = "Contractor"
PARAM_GROUP = "ADA Planning"
SHARED_PARAM_FILE = os.path.join(os.environ.get("APPDATA", ""), "ADA Tools",
                                 "ADA_SharedParameters.txt")

PARAM_CREATE = "Create & fill 'Contractor'"
PARAM_FILL = "Fill existing 'Contractor'"
PARAM_SKIP = "Phases only"

ADA_DATA_DIR = os.path.join(os.environ.get("APPDATA", ""), "ADA Tools")
AUTOMATION_PS1 = os.path.normpath(os.path.join(
    os.path.dirname(__file__), "..", "..", "lib", "Automation",
    "create_phases.ps1"))
NAMES_FILE = os.path.join(ADA_DATA_DIR, "phases_to_create.txt")
LOG_FILE = os.path.join(ADA_DATA_DIR, "phases_automation_log.txt")

MISSING_AUTO = "Create them automatically"
MISSING_MANUAL = "Copy names & open Phases"
MISSING_CLOSE = "Close"

FORBIDDEN_CHARS = u'\\:{}[]|;<>?`~'

config = script.get_config()


def clean_phase_name(text):
    """One-line phase name Revit accepts: line breaks -> ' - ' (leading
    list dashes dropped), forbidden characters removed, spaces squeezed."""
    lines = [line.strip().lstrip(u"-•*").strip()
             for line in text.splitlines()]
    name = u" - ".join(line for line in lines if line)
    name = u"".join(ch for ch in name if ch not in FORBIDDEN_CHARS)
    return u" ".join(name.split())


def run_phasing_automation(mode, names=None):
    """Start the UI-automation helper (separate PowerShell process), then
    open the Phases dialog for it to drive."""
    if not os.path.isfile(AUTOMATION_PS1):
        forms.alert("Automation helper not found:\n{}".format(AUTOMATION_PS1),
                    title=TITLE, exitscript=True)
    if not os.path.isdir(ADA_DATA_DIR):
        os.makedirs(ADA_DATA_DIR)
    if names is not None:
        import io
        with io.open(NAMES_FILE, "w", encoding="utf-8") as f:
            f.write(u"\n".join(names))

    args = ('-NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden '
            '-File "{}" -RevitPid {} -NamesFile "{}" -LogFile "{}" -Mode {}'
            .format(AUTOMATION_PS1, Process.GetCurrentProcess().Id,
                    NAMES_FILE, LOG_FILE, mode))
    info = ProcessStartInfo("powershell.exe", args)
    info.UseShellExecute = False
    info.CreateNoWindow = True
    info.WindowStyle = ProcessWindowStyle.Hidden
    Process.Start(info)
    uiapp.PostCommand(
        RevitCommandId.LookupPostableCommandId(PostableCommand.Phases))


# ---------------------------------------------------------------------------
# user input
# ---------------------------------------------------------------------------

def pick_workbook():
    path = forms.pick_file(files_filter="Excel workbook (*.xlsx;*.xlsm)|"
                                        "*.xlsx;*.xlsm",
                           init_dir=os.path.dirname(
                               config.get_option("last_file", "")),
                           title="Select the planning (Excel)")
    if not path:
        script.exit()
    try:
        return XlsxWorkbook(path)
    except Exception as ex:
        forms.alert("Could not read the Excel file:\n{}\n\nOnly .xlsx / .xlsm "
                    "files are supported (save .xls files as .xlsx)."
                    .format(ex), title=TITLE, exitscript=True)


def pick_sheet(workbook):
    names = workbook.sheet_names
    if not names:
        forms.alert("The workbook has no sheets.", title=TITLE, exitscript=True)
    if len(names) == 1:
        return names[0]
    chosen = select_from_dict(
        {name: name for name in names},
        title=TITLE,
        label="Sheet containing the planning:",
        button_name="Next",
        version=__version__,
        SelectMultiple=False)
    if not chosen:
        script.exit()
    return chosen[0]


def ask_ranges(workbook, sheet):
    """Ask the ranges until they are valid. Returns
    [(phase_name, contractor), ...] in Excel order (blank names dropped)
    and the two range strings."""
    phase_ref = config.get_option("last_phase_range", "A2:A20")
    contractor_ref = config.get_option("last_contractor_range", "B2:B20")
    while True:
        values = ask_for_inputs(
            [("Phase names range", phase_ref),
             ("Contractors range (empty = none)", contractor_ref)],
            title=TITLE,
            label="Ranges on sheet '{}' (e.g. A2:A20):".format(sheet),
            button_name="Read Excel",
            version=__version__)
        if values is None:
            script.exit()
        phase_ref, contractor_ref = values[0].strip(), values[1].strip()

        try:
            phases = workbook.read_range(sheet, phase_ref)
            contractors = (workbook.read_range(sheet, contractor_ref)
                           if contractor_ref else [""] * len(phases))
        except Exception as ex:
            forms.alert(str(ex), title=TITLE)
            continue

        if len(contractors) != len(phases):
            forms.alert("The contractors range has {} cell(s) but the phase "
                        "range has {} - they must be the same size."
                        .format(len(contractors), len(phases)), title=TITLE)
            continue
        rows = [(clean_phase_name(p), c) for p, c in zip(phases, contractors)
                if clean_phase_name(p)]
        if not rows:
            forms.alert("No phase names found in {}.".format(phase_ref),
                        title=TITLE)
            continue
        return rows, phase_ref, contractor_ref


def merge_duplicates(rows):
    """One entry per phase name, in first-seen order. A phase listed
    several times gets all its contractors joined."""
    order, contractors = [], {}
    for phase, contractor in rows:
        if phase not in contractors:
            order.append(phase)
            contractors[phase] = []
        if contractor and contractor not in contractors[phase]:
            contractors[phase].append(contractor)
    return [(p, ", ".join(contractors[p])) for p in order]


# ---------------------------------------------------------------------------
# shared parameter
# ---------------------------------------------------------------------------

def bound_definition(name):
    iterator = doc.ParameterBindings.ForwardIterator()
    while iterator.MoveNext():
        if iterator.Key.Name == name:
            return iterator.Key
    return None


def external_definition():
    """The 'Contractor' definition from the ADA shared parameter file
    (created if needed). The user's shared parameter file setting is
    restored afterwards."""
    folder = os.path.dirname(SHARED_PARAM_FILE)
    if not os.path.isdir(folder):
        os.makedirs(folder)
    if not os.path.isfile(SHARED_PARAM_FILE):
        open(SHARED_PARAM_FILE, "w").close()

    original = app.SharedParametersFilename
    try:
        app.SharedParametersFilename = SHARED_PARAM_FILE
        def_file = app.OpenSharedParameterFile()
        group = (def_file.Groups.get_Item(PARAM_GROUP)
                 or def_file.Groups.Create(PARAM_GROUP))
        definition = group.Definitions.get_Item(PARAM_NAME)
        if definition is None:
            try:
                from Autodesk.Revit.DB import SpecTypeId   # Revit 2022+
                options = ExternalDefinitionCreationOptions(
                    PARAM_NAME, SpecTypeId.String.Text)
            except Exception:
                from Autodesk.Revit.DB import ParameterType
                options = ExternalDefinitionCreationOptions(
                    PARAM_NAME, ParameterType.Text)
            options.Description = "Contractor doing the work (ADA Import Planning)"
            definition = group.Definitions.Create(options)
        return definition
    finally:
        try:
            app.SharedParametersFilename = original or ""
        except Exception:
            pass


def bind_to_model_categories(definition):
    categories = app.Create.NewCategorySet()
    for category in doc.Settings.Categories:
        if (category.CategoryType == CategoryType.Model
                and category.AllowsBoundParameters):
            categories.Insert(category)
    binding = app.Create.NewInstanceBinding(categories)
    try:
        from Autodesk.Revit.DB import GroupTypeId   # Revit 2022+
        inserted = doc.ParameterBindings.Insert(definition, binding,
                                                GroupTypeId.Phasing)
    except Exception:
        from Autodesk.Revit.DB import BuiltInParameterGroup
        inserted = doc.ParameterBindings.Insert(
            definition, binding, BuiltInParameterGroup.PG_PHASING)
    if not inserted:
        raise Exception("Revit refused to add the '{}' parameter"
                        .format(PARAM_NAME))


def eid_value(eid):
    """ElementId numeric value, 2024+ safe."""
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue


def fill_contractors(contractor_by_phase_id):
    """Set Contractor on every element created in a planning phase
    ({phase id value: contractor}). Returns (filled, failed)."""
    filled = failed = 0
    element_ids = list(FilteredElementCollector(doc)
                       .WhereElementIsNotElementType().ToElementIds())
    for el_id in element_ids:
        el = doc.GetElement(el_id)
        if el is None or not el.IsValidObject:
            continue
        created = el.get_Parameter(BuiltInParameter.PHASE_CREATED)
        if created is None:
            continue
        contractor = contractor_by_phase_id.get(
            eid_value(created.AsElementId()))
        if contractor is None:
            continue
        param = el.LookupParameter(PARAM_NAME)
        if param is None:
            continue
        if param.AsString() == contractor:
            continue
        try:
            if param.IsReadOnly or not param.Set(contractor):
                failed += 1
            else:
                filled += 1
        except Exception:
            failed += 1
    return filled, failed


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    if __shiftclick__:
        run_phasing_automation("probe")
        return

    workbook = pick_workbook()
    sheet = pick_sheet(workbook)
    rows, phase_ref, contractor_ref = ask_ranges(workbook, sheet)
    planning = merge_duplicates(rows)
    has_contractors = any(c for _, c in planning)

    project_phases = list(doc.Phases)
    by_name = {p.Name.strip().lower(): p for p in project_phases}

    def match(name):
        return by_name.get(name.strip().lower())

    items = []
    for phase, contractor in planning:
        status = "in project" if match(phase) else "MISSING"
        label = phase
        if contractor:
            label += "   |   " + contractor
        items.append(("{}   ({})".format(label, status), (phase, contractor)))

    ordered = order_list(
        items,
        title=TITLE,
        label="Planning read from Excel - check / reorder:",
        button_name="Next",
        top_text="FIRST PHASE",
        bottom_text="LAST PHASE",
        version=__version__)
    if not ordered:
        script.exit()
    planning = ordered

    # remember for the next run (e.g. after adding missing phases)
    config.last_file = workbook.path
    config.last_phase_range = phase_ref
    config.last_contractor_range = contractor_ref
    script.save_config()

    param_mode = PARAM_SKIP
    if has_contractors:
        existing = bound_definition(PARAM_NAME)
        first = PARAM_FILL if existing else PARAM_CREATE
        param_mode = select_from_buttons(
            [first, PARAM_SKIP],
            title=TITLE,
            label="Shared parameter '{}' (Phasing group, all model "
                  "categories):".format(PARAM_NAME),
            version=__version__)
        if not param_mode:
            script.exit()

    matched = [(phase, contractor, match(phase))
               for phase, contractor in planning if match(phase)]
    missing = [phase for phase, _ in planning if not match(phase)]

    report = ADAReport(TITLE)
    report.line("Excel: <b>{}</b> &nbsp;|&nbsp; sheet <b>{}</b> &nbsp;|&nbsp; "
                "phases {} {}".format(
                    os.path.basename(workbook.path), sheet, phase_ref,
                    "| contractors {}".format(contractor_ref)
                    if contractor_ref else ""))

    # order check (the API cannot reorder phases either)
    positions = [project_phases.index(p) for _, _, p in matched]
    if positions != sorted(positions):
        report.warn("The phases exist in the project in a different order "
                    "than in the planning (Revit cannot reorder phases).")

    filled = failed = 0
    if param_mode in (PARAM_CREATE, PARAM_FILL) and matched:
        try:
            with Transaction(doc, "ADA - Import Planning") as t:
                t.Start()
                if bound_definition(PARAM_NAME) is None:
                    bind_to_model_categories(external_definition())
                    report.line("Shared parameter <b>{}</b> created (Phasing "
                                "group, all model categories).".format(
                                    PARAM_NAME))
                filled, failed = fill_contractors(
                    {eid_value(p.Id): c for _, c, p in matched if c})
                t.Commit()
        except Exception as ex:
            report.error("Could not create / fill '{}': {}".format(
                PARAM_NAME, ex))

    report.subheader("Planning")
    table = []
    for phase, contractor in planning:
        status = "In project" if match(phase) else "MISSING"
        table.append([phase, contractor or "-", status])
    report.table(["Phase", "Contractor", "Status"], table)

    if param_mode != PARAM_SKIP:
        report.line("'{}' filled on <b>{}</b> element(s){}.".format(
            PARAM_NAME, filled,
            ", {} could not be changed".format(failed) if failed else ""))

    if not missing:
        report.success("All {} planning phase(s) exist in the project."
                       .format(len(planning)))
        report.flush()
        return

    report.warn("{} phase(s) are missing.".format(len(missing)))
    report.flush()

    choice = select_from_buttons(
        [MISSING_AUTO, MISSING_MANUAL, MISSING_CLOSE],
        title=TITLE,
        label=("{} phase(s) are missing. Create them automatically in the "
               "Phases dialog?\nDo not touch mouse / keyboard until the ADA "
               "message appears, then check and click OK.").format(
                   len(missing)),
        version=__version__)
    if choice == MISSING_AUTO:
        run_phasing_automation("create", missing)
    elif choice == MISSING_MANUAL:
        Clipboard.SetText(u"\r\n".join(missing))
        uiapp.PostCommand(
            RevitCommandId.LookupPostableCommandId(PostableCommand.Phases))


if __name__ == "__main__":
    main()
