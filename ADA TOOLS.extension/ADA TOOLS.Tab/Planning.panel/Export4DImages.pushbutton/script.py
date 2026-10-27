# -*- coding: utf-8 -*-
__doc__ = """Export one image per view with the Gantt chart below it and a red
line showing where the view stands in the planning (4D sequence).

  1. tick the views to export (e.g. the views made by "4D Views per
     Phase"),
  2. choose where the red line goes in the view's phase: at its start or
     at its end,
  3. choose a logo for the Gantt chart (the last one is remembered) or
     none,
  4. pick the output folder.

For each view: the view is exported as an image (1920 px wide), and the
Gantt chart of the Revit phases is drawn under it - phases finished
before the red line in grey, the view's phase in red, the following
phases in light gold. The view's phase is its View Phase parameter; the
dates come from the phase Descriptions written by Import Planning
('20/10/2026 to 24/10/2026 (5 days)'). Files are named after the views."""
__title__ = "4D\nImages"
__version__ = "Version 1.0"
__author__ = "ADA"

import os
import shutil
import tempfile

import clr

clr.AddReference('RevitAPI')
clr.AddReference('System.Drawing')

from Autodesk.Revit.DB import (
    BuiltInParameter, ElementId, ExportRange, FilteredElementCollector,
    FitDirectionType, ImageExportOptions, ImageFileType, ImageResolution, View,
    ViewType, ZoomFitType
)
from System.Collections.Generic import List
from System.Drawing import (
    Bitmap, Color, Font, FontStyle, Graphics, GraphicsUnit, Image, RectangleF,
    SolidBrush, StringAlignment, StringFormat
)
from System.Drawing.Drawing2D import InterpolationMode
from System.Drawing.Imaging import ImageFormat

from pyrevit import forms, revit, script

# Custom ADA GUI - themed list / buttons (lib/GUI) and the shared dark/gold
# themed report (lib/GUI/ReportTheme.py)
from GUI.forms import select_from_dict, select_from_buttons
from GUI.ReportTheme import ADAReport
from Snippets._gantt import (ONE_DAY, ask_logo, draw_gantt, load_logo,
                             revit_phase_rows)
from Snippets._planning import FORBIDDEN_CHARS

doc = revit.doc

TITLE = __title__.replace("\n", " ")
IMAGE_WIDTH = 1920      # px
AT_END = "End of the phase"
AT_START = "Start of the phase"
BANNER = Color.FromArgb(30, 30, 30)
GOLD = Color.FromArgb(212, 160, 23)
EXPORTABLE = (ViewType.ThreeD, ViewType.FloorPlan, ViewType.CeilingPlan,
              ViewType.Section, ViewType.Elevation, ViewType.EngineeringPlan,
              ViewType.AreaPlan, ViewType.Detail)


def view_phase_id(view):
    param = view.get_Parameter(BuiltInParameter.VIEW_PHASE)
    return param.AsElementId() if param is not None else ElementId.InvalidElementId


def clean_file_name(text):
    return u" ".join(u"".join(ch for ch in text if ch not in FORBIDDEN_CHARS
                              and ch not in u"/\\:*?\"<>|").split())


def pick_views():
    views = [v for v in FilteredElementCollector(doc).OfClass(View)
             if not v.IsTemplate and v.ViewType in EXPORTABLE and v.CanBePrinted
             and view_phase_id(v) != ElementId.InvalidElementId]
    if not views:
        forms.alert("There is no view with a phase to export.", title=TITLE,
                    exitscript=True)
    options = {}
    for v in sorted(views, key=lambda v: v.Name.lower()):
        phase = doc.GetElement(view_phase_id(v))
        options[u"{}   [{}]".format(v.Name, phase.Name if phase else u"?")] = v
    chosen = select_from_dict(options, title=TITLE, label="Views to export:",
                              button_name="Next", version=__version__,
                              SelectMultiple=True)
    if not chosen:
        script.exit()
    return chosen


def export_view_image(view, folder):
    """Export the view as a PNG in `folder` and return it as a Bitmap."""
    for f in os.listdir(folder):
        os.remove(os.path.join(folder, f))
    opts = ImageExportOptions()
    opts.ExportRange = ExportRange.SetOfViews
    opts.SetViewsAndSheets(List[ElementId]([view.Id]))
    opts.FilePath = os.path.join(folder, "view")
    opts.HLRandWFViewsFileType = ImageFileType.PNG
    opts.ShadowViewsFileType = ImageFileType.PNG
    opts.ImageResolution = ImageResolution.DPI_150
    opts.ZoomType = ZoomFitType.FitToPage
    opts.FitDirection = FitDirectionType.Horizontal
    opts.PixelSize = IMAGE_WIDTH
    doc.ExportImage(opts)
    files = [os.path.join(folder, f) for f in os.listdir(folder)
             if f.lower().endswith(".png")]
    if not files:
        raise Exception("Revit did not export an image")
    img = Image.FromFile(files[0])
    try:
        return Bitmap(img)          # copy, so the file is released
    finally:
        img.Dispose()


def compose(view_bmp, gantt_bmp, banner_text):
    """Banner + view image (scaled to the Gantt width) + Gantt chart."""
    width = gantt_bmp.Width
    view_h = int(round(view_bmp.Height * float(width) / view_bmp.Width))
    banner_h = 48
    out = Bitmap(width, banner_h + view_h + gantt_bmp.Height)
    g = Graphics.FromImage(out)
    g.Clear(Color.White)
    g.InterpolationMode = InterpolationMode.HighQualityBicubic
    g.FillRectangle(SolidBrush(BANNER), 0, 0, width, banner_h)
    sf = StringFormat()
    sf.LineAlignment = StringAlignment.Center
    g.DrawString(banner_text, Font("Segoe UI", 22.0, FontStyle.Bold, GraphicsUnit.Pixel),
                 SolidBrush(GOLD), RectangleF(14.0, 0.0, width - 28.0, float(banner_h)), sf)
    g.DrawImage(view_bmp, 0, banner_h, width, view_h)
    g.DrawImage(gantt_bmp, 0, banner_h + view_h, width, gantt_bmp.Height)
    g.Dispose()
    return out


rows = revit_phase_rows(doc)
if not any(r["start"] for r in rows):
    forms.alert("No phase has dates in its Description.\n\nRun Import Planning "
                "with a period range first (dates like '20/10/2026 to "
                "24/10/2026').", title=TITLE, exitscript=True)

views = pick_views()
position = select_from_buttons([AT_END, AT_START], title=TITLE,
                               label="Red line in the view's phase:",
                               version=__version__)
if not position:
    script.exit()
logo_path = ask_logo(TITLE, __version__)
folder = forms.pick_folder(title="Folder for the 4D images")
if not folder:
    script.exit()

row_h = 18 if len(rows) <= 30 else 14
index_of = {r["phase"].Id: i for i, r in enumerate(rows)}
report = ADAReport(TITLE)
done, failed = [], []
temp = tempfile.mkdtemp(prefix="ada_4d_")
logo = None
try:
    logo = load_logo(logo_path)
except Exception as ex:
    report.warn(u"Could not read the logo ({}): images made without it.".format(ex))
gantt_title = u"{} - Planning".format(doc.Title or u"Project")
try:
    for view in views:
        try:
            i = index_of.get(view_phase_id(view))
            row = rows[i] if i is not None else None
            marker, note = None, u""
            if row is None or not row["start"]:
                note = u"no date for the view's phase: no red line"
            elif position == AT_END:
                marker = (row["end"] + ONE_DAY, row["end"].strftime("%d/%m/%Y"))
            else:
                marker = (row["start"], row["start"].strftime("%d/%m/%Y"))
            phase_name = row["name"] if row else u"?"

            view_bmp = export_view_image(view, temp)
            gantt_bmp = draw_gantt(rows, IMAGE_WIDTH, marker=marker, current=i,
                                   row_h=row_h, title=gantt_title, logo=logo)
            banner = u"{}   |   {}{}".format(
                view.Name, phase_name, u"   |   " + marker[1] if marker else u"")
            out = compose(view_bmp, gantt_bmp, banner)
            path = os.path.join(folder, clean_file_name(view.Name) + u".png")
            out.Save(path, ImageFormat.Png)
            for b in (view_bmp, gantt_bmp, out):
                b.Dispose()
            done.append([view.Name, phase_name, marker[1] if marker else u"-", note])
        except Exception as ex:
            failed.append([view.Name, u"{}".format(ex)])
finally:
    shutil.rmtree(temp, ignore_errors=True)

if done:
    report.subheader("Exported images")
    report.table(["View", "Phase", "Red line", "Note"], done)
if failed:
    report.subheader("Failed")
    report.table(["View", "Reason"], failed)
if done:
    report.success(u"{} image(s) saved in {}".format(len(done), folder))
    os.startfile(folder)
else:
    report.error("No image was exported.")
report.flush()
