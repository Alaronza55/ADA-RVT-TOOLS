# -*- coding: utf-8 -*-
__doc__ = """Import the phases of a planning from Excel (.xlsx / .xlsm).

  1. pick the Excel file and the sheet,
  2. type the range of the phase names (e.g. B7:B52),
  3. check the list of phases read from Excel (you can reorder it).

The Revit API cannot create (or rename) phases. Phases of the planning
that already exist in the project (same name) are kept as they are.
The missing ones can be CREATED AUTOMATICALLY: the tool opens
Manage > Phases and a helper drives the dialog (After + types each
name, in planning order, after the selected phase). Do not touch the
mouse / keyboard until the ADA message appears, then check the list
and click OK (Cancel discards everything). The last file, sheet and
range are remembered.

NUMBERING: phases are named "PHASE 0 - <Excel text>", "PHASE 1 - ...",
... in the order of the planning (as confirmed in the list). Matching
with existing phases ignores the "PHASE n - " part, so reordering or
older un-numbered phases do not create duplicates.

PERIOD -> PHASE DESCRIPTION: for each phase row, the coloured (or
filled) cells of its Gantt bar are found to the right of the names;
for each one the tool goes UP its column to the first date (the date
header of the planning). Columns with no date above them (status,
contractor...) are ignored. The period from the first to the last day
is written in the phase's Description, e.g.
"20/10/2026 to 24/10/2026 (5 days)" - with the planned days and the
number of separate periods when the bar is interrupted. New phases get
it typed in the Phasing dialog; existing phases get it through the API
when Revit allows it.

Phase names are cleaned: line breaks become " - " and characters
Revit does not accept in names ( \\ : { } [ ] | ; < > ? ` ~ ) are
removed. SHIFT+CLICK the button to only record the Phasing dialog's
structure (diagnostic log, if the automation does not work).

To fill a parameter from another column of the planning (contractor,
duration...), use "Fill from Planning". The Excel file can stay open
in Excel while it is read."""
__title__ = "Import\nPlanning"
__version__ = "Version 1.1"
__author__ = "ADA"

import io
import os

import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
try:
    clr.AddReference('System.Diagnostics.Process')   # Revit 2025+ (.NET 8)
except Exception:
    clr.AddReference('System')                       # Revit <= 2024 (.NET Framework)

from Autodesk.Revit.UI import PostableCommand, RevitCommandId
from System.Diagnostics import Process, ProcessStartInfo, ProcessWindowStyle
from System.Windows.Forms import Clipboard

from pyrevit import forms, revit, script

# Custom ADA GUI - themed buttons / inputs / order list (lib/GUI), the shared
# dark/gold themed report (lib/GUI/ReportTheme.py) and the Planning helpers
# (lib/Snippets/_planning.py)
from GUI.forms import select_from_buttons, ask_for_inputs, order_list
from GUI.ReportTheme import ADAReport
from Snippets._planning import (ADA_DATA_DIR, format_period, gantt_periods,
                                name_key, numbered_phase_name, phase_rows,
                                pick_sheet, pick_workbook)

doc = revit.doc
uiapp = __revit__

TITLE = __title__.replace("\n", " ")

AUTOMATION_PS1 = os.path.normpath(os.path.join(
    os.path.dirname(__file__), "..", "..", "lib", "Automation",
    "create_phases.ps1"))
NAMES_FILE = os.path.join(ADA_DATA_DIR, "phases_to_create.txt")
LOG_FILE = os.path.join(ADA_DATA_DIR, "phases_automation_log.txt")

MISSING_CANCEL = "Cancel"

config = script.get_config()


def run_phasing_automation(mode, names=None):
    """Start the UI-automation helper (separate PowerShell process), then
    open the Phases dialog for it to drive."""
    if not os.path.isfile(AUTOMATION_PS1):
        forms.alert("Automation helper not found:\n{}".format(AUTOMATION_PS1),
                    title=TITLE, exitscript=True)
    if not os.path.isdir(ADA_DATA_DIR):
        os.makedirs(ADA_DATA_DIR)
    if names is not None:
        with io.open(NAMES_FILE, "w", encoding="utf-8") as f:
            f.write(u"\n".join(names))

    args = ('-NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden '
            '-File "{}" -RevitPid {} -NamesFile "{}" -LogFile "{}" -Mode {}'
            .format(AUTOMATION_PS1, Process.GetCurrentProcess().Id,
                    NAMES_FILE, LOG_FILE, mode))
    with io.open(LOG_FILE, "w", encoding="utf-8") as f:
        f.write(u"Import Planning: starting helper ({}, {} name(s))\n"
                u"powershell.exe {}\n".format(
                    mode, len(names or []), args))

    info = ProcessStartInfo("powershell.exe", args)
    info.UseShellExecute = False
    info.CreateNoWindow = True
    info.WindowStyle = ProcessWindowStyle.Hidden
    info.RedirectStandardError = True
    helper = Process.Start(info)

    # a helper that dies at once (blocked PowerShell, missing file...) must
    # not leave the user looking at an empty Phases dialog
    if helper.WaitForExit(1500):
        error = helper.StandardError.ReadToEnd()
        forms.alert("The automation helper stopped immediately (exit code "
                    "{}).\n\n{}\n\nLog: {}".format(
                        helper.ExitCode, error[:800], LOG_FILE),
                    title=TITLE, exitscript=True)

    uiapp.PostCommand(
        RevitCommandId.LookupPostableCommandId(PostableCommand.Phases))


def ask_phase_range(workbook, sheet):
    """Ask the range until it gives phase names. Returns ([(name, excel row)]
    unique, in Excel order; first column right of the range; range string)."""
    phase_ref = config.get_option("last_phase_range", "A2:A20")
    while True:
        values = ask_for_inputs(
            [("Phase names range", phase_ref)],
            title=TITLE,
            label="Range on sheet '{}' (e.g. B7:B52):".format(sheet),
            button_name="Read Excel",
            version=__version__)
        if values is None:
            script.exit()
        phase_ref = values[0].strip()
        try:
            rows, first_col = phase_rows(workbook, sheet, phase_ref)
        except Exception as ex:
            forms.alert(str(ex), title=TITLE)
            continue
        if rows:
            return rows, first_col, phase_ref
        forms.alert("No phase names found in {}.".format(phase_ref),
                    title=TITLE)


def description_param(phase):
    """The phase's Description parameter, or None."""
    from Autodesk.Revit.DB import BuiltInParameter
    for bip_name in ("PHASE_DESCRIPTION", "ALL_MODEL_DESCRIPTION"):
        bip = getattr(BuiltInParameter, bip_name, None)
        if bip is not None:
            param = phase.get_Parameter(bip)
            if param is not None:
                return param
    return phase.LookupParameter("Description")


def set_existing_descriptions(pairs):
    """[(phase, description)] -> (set, unchanged, failed names). Writes the
    description of phases that already exist, through the API."""
    from Autodesk.Revit.DB import Transaction
    done = same = 0
    failed = []
    with Transaction(doc, "ADA - Phase descriptions") as t:
        t.Start()
        for phase, text in pairs:
            param = description_param(phase)
            try:
                if param is not None and param.AsString() == text:
                    same += 1
                elif param is None or param.IsReadOnly or not param.Set(text):
                    failed.append(phase.Name)
                else:
                    done += 1
            except Exception:
                failed.append(phase.Name)
        t.Commit()
    return done, same, failed


def main():
    if __shiftclick__:
        run_phasing_automation("probe")
        return

    workbook = pick_workbook(config, TITLE)
    sheet = pick_sheet(workbook, TITLE, __version__)
    rows, first_col, phase_ref = ask_phase_range(workbook, sheet)

    # period of each phase from its Gantt bar -> phase description
    periods = gantt_periods(workbook, sheet, [r for _, r in rows], first_col)
    description = {name: format_period(periods.get(r)) for name, r in rows}
    planning = [name for name, _ in rows]

    project_phases = list(doc.Phases)
    by_key = {name_key(p.Name): p for p in project_phases}

    def match(name):
        return by_key.get(name_key(name))

    items = []
    for name in planning:
        label = name
        if description[name]:
            label += u"   |   " + description[name]
        label += u"   ({})".format("in project" if match(name) else "MISSING")
        items.append((label, name))
    ordered = order_list(
        items,
        title=TITLE,
        label="Phases read from Excel - check / reorder:",
        button_name="Next",
        top_text="FIRST PHASE  (PHASE 0)",
        bottom_text="LAST PHASE  (numbered in this order)",
        version=__version__)
    if not ordered:
        script.exit()
    planning = ordered

    # remember for the next run
    config.last_file = workbook.path
    config.last_phase_range = phase_ref
    script.save_config()

    # Revit phase names: "PHASE 0 - ...", "PHASE 1 - ..." in planning order
    numbered = {name: numbered_phase_name(i, name)
                for i, name in enumerate(planning)}
    missing = [name for name in planning if not match(name)]
    renumber = [(match(n).Name, numbered[n]) for n in planning
                if match(n) and match(n).Name != numbered[n]]

    report = ADAReport(TITLE)
    report.line(u"Excel: <b>{}</b> &nbsp;|&nbsp; sheet <b>{}</b> &nbsp;|&nbsp; "
                u"phases {}".format(os.path.basename(workbook.path), sheet,
                                    phase_ref))

    # order check (the API cannot reorder phases either)
    positions = [project_phases.index(match(n)) for n in planning if match(n)]
    if positions != sorted(positions):
        report.warn("The phases exist in the project in a different order "
                    "than in the planning (Revit cannot reorder phases).")

    # descriptions of the phases that already exist (through the API)
    existing = [(match(n), description[n]) for n in planning
                if match(n) and description[n]]
    if existing:
        try:
            done, same, failed = set_existing_descriptions(existing)
            if done or same:
                report.line(u"Description set on <b>{}</b> existing phase(s)"
                            u"{}.".format(done, u" ({} already up to date)"
                                          .format(same) if same else u""))
            if failed:
                report.warn(u"Revit refused the description of {} existing "
                            u"phase(s) - type it in Manage > Phases.".format(
                                len(failed)))
        except Exception as ex:
            report.warn(u"Could not set the descriptions of the existing "
                        u"phases: {}".format(ex))

    no_period = [n for n in planning if not description[n]]
    report.subheader("Phases")
    report.table(["Phase", "Period (description)", "Status"],
                 [[match(name).Name if match(name) else numbered[name],
                   description[name] or "-",
                   "In project" if match(name) else "MISSING"]
                  for name in planning])
    if renumber:
        report.warn(u"{} existing phase(s) do not have their planning number "
                    u"(e.g. '{}' should be '{}') - Revit cannot rename phases "
                    u"from a script: rename them in Manage > Phases if needed."
                    .format(len(renumber), renumber[0][0], renumber[0][1]))
    if len(no_period) == len(planning):
        report.warn("No period found: no coloured / filled cell with a date "
                    "above it (going up the column) on the phase rows.")
    elif no_period:
        report.warn("{} phase(s) have no bar in the planning - no period."
                    .format(len(no_period)))

    if not missing:
        report.success("All {} planning phase(s) exist in the project."
                       .format(len(planning)))
        report.flush()
        return

    report.warn("{} phase(s) are missing.".format(len(missing)))
    report.flush()

    Clipboard.SetText(u"\r\n".join(numbered[n] for n in missing))
    create_label = "Create {} phases automatically".format(len(missing))
    choice = select_from_buttons(
        [create_label, MISSING_CANCEL],
        title=TITLE,
        label=("{} phase(s) are missing (names copied to the clipboard).\n"
               "The Phases dialog will open and fill itself - do not touch "
               "mouse / keyboard until the ADA message appears, then check "
               "and click OK.").format(len(missing)),
        version=__version__)
    if choice == create_label:
        # one line per phase: name<TAB>description (typed in the dialog)
        run_phasing_automation("create", [u"{}\t{}".format(numbered[n],
                                                           description[n])
                                          for n in missing])


if __name__ == "__main__":
    main()
