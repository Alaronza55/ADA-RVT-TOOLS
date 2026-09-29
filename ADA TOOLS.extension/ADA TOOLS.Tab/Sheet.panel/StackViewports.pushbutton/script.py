# -*- coding: utf-8 -*-
__doc__ = """Stack the viewports of the active sheet on top of each other, in
the order you choose.

  1. all viewports of the sheet are listed, top to bottom as they are
     now - select one and move it with Up / Down / To Top / To Bottom
     until the list shows the stack you want (last item = lowest),
  2. choose the horizontal alignment (left, centre, right, or keep
     each viewport's horizontal position),
  3. type the gap between viewports in millimetres.

The lowest viewport stays where it is. Each next viewport is placed
directly above the previous one, the viewport title included, with
the chosen gap. Pinned viewports are unpinned for the move and
re-pinned. The whole change is one undoable transaction."""
__title__ = "Stack\nViewports"
__version__ = "Version 1.0"
__author__ = "ADA"

import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import Transaction, ViewSheet, XYZ

from pyrevit import forms, revit, script

# Custom ADA GUI - order list (lib/GUI/OrderList.py), button-choice popup
# (lib/GUI/SelectFromButtons.py), text-input popup (lib/GUI/AskForInputs.py)
# and the shared dark/gold themed report (lib/GUI/ReportTheme.py)
from GUI.forms import order_list, select_from_buttons, ask_for_inputs
from GUI.ReportTheme import ADAReport

doc = revit.doc

TITLE = __title__.replace("\n", " ")
MM_TO_FEET = 1.0 / 304.8

ALIGN_LEFT = "Align left"
ALIGN_CENTER = "Align centre"
ALIGN_RIGHT = "Align right"
ALIGN_KEEP = "Keep horizontal position"


class Box(object):
    """2D extents on the sheet (feet)."""

    def __init__(self, min_x, min_y, max_x, max_y):
        self.min_x, self.min_y = min_x, min_y
        self.max_x, self.max_y = max_x, max_y

    @classmethod
    def from_outline(cls, outline):
        return cls(outline.MinimumPoint.X, outline.MinimumPoint.Y,
                   outline.MaximumPoint.X, outline.MaximumPoint.Y)

    def union(self, other):
        return Box(min(self.min_x, other.min_x), min(self.min_y, other.min_y),
                   max(self.max_x, other.max_x), max(self.max_y, other.max_y))

    @property
    def center_x(self):
        return (self.min_x + self.max_x) / 2.0


def view_name(viewport):
    return doc.GetElement(viewport.ViewId).Name


def viewport_boxes(viewport):
    """(frame box, full box including the viewport title)."""
    frame = Box.from_outline(viewport.GetBoxOutline())
    full = frame
    try:
        full = frame.union(Box.from_outline(viewport.GetLabelOutline()))
    except Exception:
        pass  # Revit < 2018 or no title shown
    return frame, full


def parse_mm(text):
    cleaned = (text or "").strip().replace(" ", "").replace(",", ".")
    try:
        return float(cleaned)
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# user input
# ---------------------------------------------------------------------------

def pick_order(viewports):
    # initial list: as they are now on the sheet, highest first
    current = sorted(viewports,
                     key=lambda vp: viewport_boxes(vp)[1].max_y,
                     reverse=True)
    items = []
    for vp in current:
        view = doc.GetElement(vp.ViewId)
        items.append(("{}  (1:{})".format(view.Name, view.Scale), vp))

    ordered = order_list(
        items,
        title=TITLE,
        label="Order the viewports - the last one is the lowest:",
        button_name="Next",
        top_text="TOP OF THE STACK",
        bottom_text="BOTTOM OF THE STACK (stays in place)",
        version=__version__)
    if not ordered:
        script.exit()
    return ordered


def pick_alignment():
    alignment = select_from_buttons(
        [ALIGN_LEFT, ALIGN_CENTER, ALIGN_RIGHT, ALIGN_KEEP],
        title=TITLE,
        label="Horizontal alignment of the stack:",
        version=__version__)
    if not alignment:
        script.exit()
    return alignment


def ask_gap():
    default = "10"
    while True:
        values = ask_for_inputs(
            [("Gap between viewports (mm)", default)],
            title=TITLE,
            label="Vertical gap between the viewports:",
            button_name="Stack",
            version=__version__)
        if values is None:
            script.exit()
        gap = parse_mm(values[0])
        if gap is not None and gap >= 0:
            return gap
        forms.alert("Type a number of millimetres, 0 or more.", title=TITLE)
        default = values[0]


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main():
    sheet = doc.ActiveView
    if not isinstance(sheet, ViewSheet):
        forms.alert("Open a sheet first - the viewports to stack must be on "
                    "the active sheet.", title=TITLE, exitscript=True)

    viewports = [doc.GetElement(vp_id) for vp_id in sheet.GetAllViewports()]
    if len(viewports) < 2:
        forms.alert("This sheet needs at least two viewports.", title=TITLE,
                    exitscript=True)

    top_to_bottom = pick_order(viewports)
    alignment = pick_alignment()
    gap_mm = ask_gap()
    gap = gap_mm * MM_TO_FEET

    bottom_to_top = list(reversed(top_to_bottom))
    anchor_frame, anchor_full = viewport_boxes(bottom_to_top[0])

    # compute every move first (moves are pure translations)
    moves = []
    prev_top = anchor_full.max_y
    for vp in bottom_to_top[1:]:
        frame, full = viewport_boxes(vp)
        dy = (prev_top + gap) - full.min_y
        if alignment == ALIGN_LEFT:
            dx = anchor_frame.min_x - frame.min_x
        elif alignment == ALIGN_RIGHT:
            dx = anchor_frame.max_x - frame.max_x
        elif alignment == ALIGN_CENTER:
            dx = anchor_frame.center_x - frame.center_x
        else:
            dx = 0.0
        moves.append((vp, XYZ(dx, dy, 0)))
        prev_top = full.max_y + dy

    report = ADAReport(TITLE)
    report.line("Sheet: <b>{} - {}</b>".format(sheet.SheetNumber, sheet.Name))
    report.line("Alignment: <b>{}</b> &nbsp;|&nbsp; Gap: <b>{:g} mm</b>".format(
        alignment, gap_mm))

    try:
        with Transaction(doc, "ADA - Stack Viewports") as t:
            t.Start()
            for vp, delta in moves:
                if delta.GetLength() < 1e-9:
                    continue
                was_pinned = vp.Pinned
                if was_pinned:
                    vp.Pinned = False
                vp.SetBoxCenter(vp.GetBoxCenter() + delta)
                if was_pinned:
                    vp.Pinned = True
            t.Commit()
    except Exception as ex:
        report.error("Failed, nothing was changed: {}".format(ex))
        report.flush()
        return

    rows = []
    for position, vp in enumerate(top_to_bottom, 1):
        if vp.Id == bottom_to_top[0].Id:
            note = "Bottom - stays in place"
        else:
            note = "Moved"
        rows.append([str(position), view_name(vp), note])

    report.subheader("Stack (top to bottom)")
    report.table(["#", "View", "Result"], rows)
    report.success("{} viewport(s) stacked above '{}'.".format(
        len(moves), view_name(bottom_to_top[0])))
    report.flush()


if __name__ == "__main__":
    main()
