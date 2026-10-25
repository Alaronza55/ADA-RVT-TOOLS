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
from Snippets._planning import (ADA_DATA_DIR, clean_name, name_key,
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
    """Ask the range until it gives phase names. Returns (unique cleaned
    names in Excel order, range string)."""
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
            cells = workbook.read_range(sheet, phase_ref)
        except Exception as ex:
            forms.alert(str(ex), title=TITLE)
            continue

        names, seen = [], set()
        for cell in cells:
            name = clean_name(cell)
            if name and name_key(name) not in seen:
                seen.add(name_key(name))
                names.append(name)
        if names:
            return names, phase_ref
        forms.alert("No phase names found in {}.".format(phase_ref),
                    title=TITLE)


def main():
    if __shiftclick__:
        run_phasing_automation("probe")
        return

    workbook = pick_workbook(config, TITLE)
    sheet = pick_sheet(workbook, TITLE, __version__)
    planning, phase_ref = ask_phase_range(workbook, sheet)

    project_phases = list(doc.Phases)
    by_key = {name_key(p.Name): p for p in project_phases}

    def match(name):
        return by_key.get(name_key(name))

    items = [(u"{}   ({})".format(name, "in project" if match(name) else
                                  "MISSING"), name)
             for name in planning]
    ordered = order_list(
        items,
        title=TITLE,
        label="Phases read from Excel - check / reorder:",
        button_name="Next",
        top_text="FIRST PHASE",
        bottom_text="LAST PHASE",
        version=__version__)
    if not ordered:
        script.exit()
    planning = ordered

    # remember for the next run
    config.last_file = workbook.path
    config.last_phase_range = phase_ref
    script.save_config()

    missing = [name for name in planning if not match(name)]

    report = ADAReport(TITLE)
    report.line(u"Excel: <b>{}</b> &nbsp;|&nbsp; sheet <b>{}</b> &nbsp;|&nbsp; "
                u"phases {}".format(os.path.basename(workbook.path), sheet,
                                    phase_ref))

    # order check (the API cannot reorder phases either)
    positions = [project_phases.index(match(n)) for n in planning if match(n)]
    if positions != sorted(positions):
        report.warn("The phases exist in the project in a different order "
                    "than in the planning (Revit cannot reorder phases).")

    report.subheader("Phases")
    report.table(["Phase", "Status"],
                 [[name, "In project" if match(name) else "MISSING"]
                  for name in planning])

    if not missing:
        report.success("All {} planning phase(s) exist in the project."
                       .format(len(planning)))
        report.flush()
        return

    report.warn("{} phase(s) are missing.".format(len(missing)))
    report.flush()

    Clipboard.SetText(u"\r\n".join(missing))
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
        run_phasing_automation("create", missing)


if __name__ == "__main__":
    main()
