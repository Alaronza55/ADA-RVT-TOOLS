# -*- coding: utf-8 -*-
__doc__ = """Export a Gantt chart of the Revit phases as a PNG image.

One row per phase, in the project's phase order. The dates are read from
each phase's Description, as written by Import Planning:
  '20/10/2026 to 24/10/2026 (5 days)'  or  '20/10/2026 (1 day)'
(first date = start, second date = end). Phases without a date in their
Description are listed without a bar.

  1. choose whether to draw a red line at today's date,
  2. choose a logo (put top right; the last one is remembered) or none,
  3. choose where to save the PNG - it opens when done."""
__title__ = "Export\nGantt"
__version__ = "Version 1.0"
__author__ = "ADA"

import os
from datetime import date

from System.Drawing.Imaging import ImageFormat

from pyrevit import forms, revit, script

# Custom ADA GUI - button popup (lib/GUI) and the shared dark/gold themed
# report (lib/GUI/ReportTheme.py)
from GUI.forms import select_from_buttons
from GUI.ReportTheme import ADAReport
from Snippets._gantt import ask_logo, draw_gantt, load_logo, revit_phase_rows
from Snippets._planning import FORBIDDEN_CHARS

doc = revit.doc

TITLE = __title__.replace("\n", " ")
WIDTH = 2000            # px
NO_LINE = "No line"
TODAY = "Red line at today"

rows = revit_phase_rows(doc)
dated = [r for r in rows if r["start"]]
if not dated:
    forms.alert("No phase has dates in its Description.\n\nRun Import Planning "
                "with a period range first (dates like '20/10/2026 to "
                "24/10/2026').", title=TITLE, exitscript=True)

choice = select_from_buttons([NO_LINE, TODAY], title=TITLE,
                             label="Red line on the Gantt chart:",
                             version=__version__)
if not choice:
    script.exit()

marker = current = None
if choice == TODAY:
    today = date.today()
    marker = (today, today.strftime("%d/%m/%Y"))
    for i, r in enumerate(rows):
        if r["start"] and r["start"] <= today <= r["end"]:
            current = i
            break

logo_path = ask_logo(TITLE, __version__)

project = doc.Title or "Project"
default_name = u"Gantt - {}".format(
    u"".join(ch for ch in project if ch not in FORBIDDEN_CHARS))
path = forms.save_file(file_ext="png", default_name=default_name)
if not path:
    script.exit()

report = ADAReport(TITLE)
try:
    logo = None
    try:
        logo = load_logo(logo_path)
    except Exception as ex:
        report.warn(u"Could not read the logo ({}): chart made without it.".format(ex))
    bmp = draw_gantt(rows, WIDTH, marker=marker, current=current,
                     title=u"{} - Planning".format(project), logo=logo)
    bmp.Save(path, ImageFormat.Png)
    bmp.Dispose()
    t0 = min(r["start"] for r in dated)
    t1 = max(r["end"] for r in dated)
    report.line(u"<b>{}</b> phase(s), {} with dates, from {} to {}.".format(
        len(rows), len(dated), t0.strftime("%d/%m/%Y"), t1.strftime("%d/%m/%Y")))
    missing = [r["name"] for r in rows if not r["start"]]
    if missing:
        report.warn(u"No date in the Description of: {}".format(u", ".join(missing)))
    report.success(u"Gantt chart saved: {}".format(path))
    os.startfile(path)
except Exception as ex:
    report.error(u"Could not export the Gantt chart: {}".format(ex))
report.flush()
