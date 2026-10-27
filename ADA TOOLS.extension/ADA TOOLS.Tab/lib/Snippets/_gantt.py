# -*- coding: utf-8 -*-
"""Gantt chart of the Revit phases, drawn with System.Drawing.

The dates come from the phase Descriptions, as written by Import Planning
('20/10/2026 to 24/10/2026 (5 days)' or '20/10/2026 (1 day)'): the first
date found is the start, the second the end (same day if only one)."""
import re
from datetime import date, timedelta

import clr

clr.AddReference('System.Drawing')
from System.Drawing import (
    Bitmap, Color, Font, FontStyle, Graphics, GraphicsUnit, Pen, RectangleF,
    SolidBrush, StringAlignment, StringFormat, StringFormatFlags,
    StringTrimming
)
from System.Drawing.Drawing2D import InterpolationMode, SmoothingMode
from System.Drawing.Text import TextRenderingHint

_DATE = re.compile(r"(\d{1,2})/(\d{1,2})/(\d{4})")
ONE_DAY = timedelta(days=1)

WHITE = Color.White
TEXT = Color.FromArgb(40, 40, 40)
GRID = Color.FromArgb(225, 225, 225)
ZEBRA = Color.FromArgb(246, 246, 246)
HEADER = Color.FromArgb(235, 235, 235)
HEADER_ALT = Color.FromArgb(222, 222, 222)
BAR = Color.FromArgb(212, 160, 23)          # ADA gold
BAR_DONE = Color.FromArgb(150, 150, 150)    # phases finished before the red line
BAR_TODO = Color.FromArgb(232, 205, 130)    # phases after the red line
BAR_CURRENT = Color.FromArgb(205, 40, 40)
CURRENT_ROW = Color.FromArgb(253, 228, 228)
RED = Color.FromArgb(220, 0, 0)


# ---------------------------------------------------------------------------
# phases -> rows
# ---------------------------------------------------------------------------

def parse_period(text):
    """(start, end) dates found in a description, or (None, None)."""
    found = []
    for d, m, y in _DATE.findall(text or u""):
        try:
            found.append(date(int(y), int(m), int(d)))
        except ValueError:
            continue
    if not found:
        return None, None
    return found[0], (found[1] if len(found) > 1 else found[0])


def phase_description(phase):
    from Autodesk.Revit.DB import BuiltInParameter
    for bip_name in ("PHASE_DESCRIPTION", "ALL_MODEL_DESCRIPTION"):
        bip = getattr(BuiltInParameter, bip_name, None)
        param = phase.get_Parameter(bip) if bip is not None else None
        if param is not None:
            return param.AsString() or u""
    param = phase.LookupParameter("Description")
    return (param.AsString() or u"") if param is not None else u""


def revit_phase_rows(doc):
    """[{'phase', 'name', 'start', 'end'}] in the project's phase order
    (start / end None when the description holds no date)."""
    rows = []
    for phase in doc.Phases:
        start, end = parse_period(phase_description(phase))
        if start and end and end < start:
            start, end = end, start
        rows.append({"phase": phase, "name": phase.Name, "start": start, "end": end})
    return rows


# ---------------------------------------------------------------------------
# drawing
# ---------------------------------------------------------------------------

def _string_format(align=StringAlignment.Near):
    sf = StringFormat(StringFormatFlags.NoWrap)
    sf.Alignment = align
    sf.LineAlignment = StringAlignment.Center
    sf.Trimming = StringTrimming.EllipsisCharacter
    return sf


def _next_month(d):
    return date(d.year + d.month // 12, d.month % 12 + 1, 1)


def ask_logo(title, version):
    """Ask for the logo to put on the Gantt chart; the choice is remembered
    (shared by all the Gantt tools). Returns an image path or None."""
    import os
    from pyrevit import forms, script
    from GUI.forms import select_from_buttons

    config = script.get_config("ADA_Gantt")
    last = config.get_option("logo_path", "")
    choose, none = "Choose a logo...", "No logo"
    options = [choose, none]
    keep = None
    if last and os.path.isfile(last):
        keep = u"Logo: {}".format(os.path.basename(last))
        options.insert(0, keep)
    choice = select_from_buttons(options, title=title,
                                 label="Logo on the Gantt chart:", version=version)
    if choice is None:
        script.exit()
    if choice == keep:
        return last
    if choice == none:
        return None
    path = forms.pick_file(files_filter="Images (*.png;*.jpg;*.jpeg;*.bmp)|"
                                        "*.png;*.jpg;*.jpeg;*.bmp")
    if not path:
        return last if keep else None
    config.logo_path = path
    script.save_config()
    return path


def load_logo(path):
    """Bitmap copy of an image file (the file is not kept locked), or None."""
    if not path:
        return None
    from System.Drawing import Image
    img = Image.FromFile(path)
    try:
        return Bitmap(img)
    finally:
        img.Dispose()


def draw_gantt(rows, width, marker=None, current=None, title=None, row_h=22,
               logo=None):
    """Bitmap of the Gantt chart.

    rows:    revit_phase_rows() (or dicts with 'name', 'start', 'end').
    marker:  (date, label) - red vertical line at the START of that date.
    current: index of the row to highlight (the phase the marker is in).
    logo:    Bitmap drawn top right, in the title band.
    """
    dated = [r for r in rows if r["start"]]
    if not dated:
        raise ValueError("no phase has dates in its Description")
    t0 = min(r["start"] for r in dated)
    t1 = max(r["end"] for r in dated) + ONE_DAY
    if marker:
        t0, t1 = min(t0, marker[0]), max(t1, marker[0] + ONE_DAY)
    days = float(max(1, (t1 - t0).days))

    font = Font("Segoe UI", row_h * 0.55, FontStyle.Regular, GraphicsUnit.Pixel)
    font_b = Font("Segoe UI", row_h * 0.55, FontStyle.Bold, GraphicsUnit.Pixel)
    font_t = Font("Segoe UI", row_h * 0.8, FontStyle.Bold, GraphicsUnit.Pixel)

    probe = Graphics.FromImage(Bitmap(1, 1))
    name_w = max(probe.MeasureString(r["name"], font).Width for r in rows) + 16
    name_w = float(min(name_w, width * 0.4))
    probe.Dispose()

    title_h = row_h * 1.8 if title else 0.0
    if logo is not None:
        title_h = max(title_h, row_h * 3.0)
    header_h = row_h * 2.0
    y_rows = title_h + header_h
    height = int(y_rows + row_h * len(rows) + row_h * 0.6)
    x0 = name_w
    tw = float(width) - name_w - 12

    def X(d):
        return x0 + tw * (d - t0).days / days

    bmp = Bitmap(int(width), height)
    g = Graphics.FromImage(bmp)
    g.SmoothingMode = SmoothingMode.AntiAlias
    g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit
    g.Clear(WHITE)
    text_b = SolidBrush(TEXT)
    grid = Pen(GRID, 1.0)
    left, centre = _string_format(), _string_format(StringAlignment.Center)

    logo_w = 0.0
    if logo is not None:
        box_h, box_w = title_h - 8.0, width * 0.25
        scale = min(box_h / logo.Height, box_w / logo.Width)
        logo_w, logo_h = logo.Width * scale, logo.Height * scale
        g.InterpolationMode = InterpolationMode.HighQualityBicubic
        g.DrawImage(logo, width - 8.0 - logo_w, (title_h - logo_h) / 2.0, logo_w, logo_h)
    if title:
        g.DrawString(title, font_t, text_b,
                     RectangleF(6.0, 0.0, width - 24.0 - logo_w, title_h), left)

    # zebra rows + highlighted row
    for i in range(len(rows)):
        y = y_rows + i * row_h
        if i == current:
            g.FillRectangle(SolidBrush(CURRENT_ROW), 0.0, y, float(width), float(row_h))
        elif i % 2:
            g.FillRectangle(SolidBrush(ZEBRA), 0.0, y, float(width), float(row_h))

    # month band + day / week ticks
    px_per_day = tw / days
    d = date(t0.year, t0.month, 1)
    alt = False
    while d < t1:
        nxt = _next_month(d)
        xa, xb = X(max(d, t0)), X(min(nxt, t1))
        g.FillRectangle(SolidBrush(HEADER_ALT if alt else HEADER), xa, title_h, xb - xa, float(row_h))
        label = d.strftime("%B %Y") if xb - xa > row_h * 6 else d.strftime("%m/%y")
        g.DrawString(label, font_b, text_b, RectangleF(xa, title_h, xb - xa, float(row_h)), centre)
        g.DrawLine(Pen(Color.FromArgb(190, 190, 190), 1.0), xa, title_h, xa, float(height))
        alt = not alt
        d = nxt
    if px_per_day >= row_h * 0.9:
        step = 1                                   # every day
    elif px_per_day * 7 >= row_h * 1.2:
        step = 7                                   # every week (Mondays)
    else:
        step = 0
    if step:
        d = t0
        if step == 7:
            d = t0 + timedelta(days=(7 - t0.weekday()) % 7)
        while d < t1:
            x = X(d)
            g.DrawLine(grid, x, title_h + row_h, x, float(height))
            g.DrawString(str(d.day), font, text_b,
                         RectangleF(x + (0 if step == 1 else 2), title_h + row_h,
                                    px_per_day * step, float(row_h)),
                         centre if step == 1 else left)
            d += timedelta(days=step)

    # names + bars
    marker_day = marker[0] if marker else None
    for i, row in enumerate(rows):
        y = y_rows + i * row_h
        g.DrawString(row["name"], font_b if i == current else font, text_b,
                     RectangleF(4.0, y, name_w - 8, float(row_h)), left)
        if not row["start"]:
            continue
        if i == current:
            colour = BAR_CURRENT
        elif marker_day is None:
            colour = BAR
        elif row["end"] + ONE_DAY <= marker_day:
            colour = BAR_DONE
        else:
            colour = BAR_TODO
        xa, xb = X(row["start"]), X(row["end"] + ONE_DAY)
        g.FillRectangle(SolidBrush(colour), xa, y + row_h * 0.2, max(2.0, xb - xa), row_h * 0.6)

    g.DrawLine(Pen(Color.FromArgb(160, 160, 160), 1.0), x0, title_h, x0, float(height))

    # red line
    if marker:
        x = X(marker[0])
        g.DrawLine(Pen(RED, max(2.0, row_h / 8.0)), x, title_h + row_h, x, float(height))
        label_w = g.MeasureString(marker[1], font_b).Width + 8
        lx = min(max(x - label_w / 2.0, x0), width - label_w - 2)
        g.FillRectangle(SolidBrush(RED), lx, title_h + row_h, label_w, float(row_h))
        g.DrawString(marker[1], font_b, SolidBrush(WHITE),
                     RectangleF(lx, title_h + row_h, label_w, float(row_h)), centre)

    g.Dispose()
    return bmp
