# -*- coding: utf-8 -*-
__doc__ = """Set the draw order of the viewports on the active sheet - which
viewport is drawn on top of which where they overlap. Viewports do
NOT move.

All viewports of the sheet are listed in their current draw order
(top of the list = drawn on top). Select one and move it with
Up / Down / To Top / To Bottom until the list shows the order you
want, then confirm.

Revit draws viewports in the order they were placed and has no draw
order setting for them, so the tool re-places each viewport, lowest
first, at exactly the same position with the same viewport type,
rotation, detail number, title position and pinned state. The views
themselves are untouched. The whole change is one undoable
transaction."""
__title__ = "Viewport\nDraw Order"
__version__ = "Version 1.0"
__author__ = "ADA"

import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import (
    BuiltInParameter, Transaction, ViewSheet, Viewport
)

from pyrevit import forms, revit, script

# Custom ADA GUI - order list (lib/GUI/OrderList.py) and the shared
# dark/gold themed report (lib/GUI/ReportTheme.py)
from GUI.forms import order_list
from GUI.ReportTheme import ADAReport

doc = revit.doc

TITLE = __title__.replace("\n", " ")


def eid_value(eid):
    """ElementId numeric value, 2024+ safe."""
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue


def view_name(viewport):
    return doc.GetElement(viewport.ViewId).Name


class ViewportState(object):
    """Everything needed to re-place a viewport exactly as it was."""

    def __init__(self, viewport):
        self.view_id = viewport.ViewId
        self.name = view_name(viewport)
        self.center = viewport.GetBoxCenter()
        self.type_id = viewport.GetTypeId()
        self.rotation = viewport.Rotation
        self.pinned = viewport.Pinned
        param = viewport.get_Parameter(
            BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
        self.detail_number = param.AsString() if param else None
        # title position / line length - Revit 2022+
        self.label_offset = getattr(viewport, "LabelOffset", None)
        self.label_length = getattr(viewport, "LabelLineLength", None)

    def recreate(self, sheet):
        viewport = Viewport.Create(doc, sheet.Id, self.view_id, self.center)
        if viewport.GetTypeId() != self.type_id:
            viewport.ChangeTypeId(self.type_id)
        if viewport.Rotation != self.rotation:
            viewport.Rotation = self.rotation
        viewport.SetBoxCenter(self.center)
        if self.label_offset is not None:
            viewport.LabelOffset = self.label_offset
        if self.label_length is not None:
            viewport.LabelLineLength = self.label_length
        return viewport


def set_detail_number(viewport, number):
    param = viewport.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
    if param and not param.IsReadOnly and number:
        param.Set(number)


def pick_order(viewports):
    # current draw order: placed later = drawn on top = higher element id
    current = sorted(viewports, key=lambda vp: eid_value(vp.Id), reverse=True)
    items = []
    for vp in current:
        view = doc.GetElement(vp.ViewId)
        items.append(("{}  (1:{})".format(view.Name, view.Scale), vp))

    ordered = order_list(
        items,
        title=TITLE,
        label="Draw order - the first one is drawn on top:",
        button_name="Apply Order",
        top_text="DRAWN ON TOP",
        bottom_text="DRAWN AT THE BOTTOM",
        version=__version__)
    if not ordered:
        script.exit()
    return ordered, current


def main():
    sheet = doc.ActiveView
    if not isinstance(sheet, ViewSheet):
        forms.alert("Open a sheet first - the viewports must be on the "
                    "active sheet.", title=TITLE, exitscript=True)

    viewports = [doc.GetElement(vp_id) for vp_id in sheet.GetAllViewports()]
    if len(viewports) < 2:
        forms.alert("This sheet needs at least two viewports.", title=TITLE,
                    exitscript=True)

    top_to_bottom, current = pick_order(viewports)

    report = ADAReport(TITLE)
    report.line("Sheet: <b>{} - {}</b>".format(sheet.SheetNumber, sheet.Name))

    if [vp.Id for vp in top_to_bottom] == [vp.Id for vp in current]:
        report.success("The order is unchanged - nothing was done.")
        report.flush()
        return

    # re-placed lowest first, so the last one re-placed is drawn on top
    states = [ViewportState(vp) for vp in reversed(top_to_bottom)]

    try:
        with Transaction(doc, "ADA - Viewport Draw Order") as t:
            t.Start()
            for vp in top_to_bottom:
                if vp.Pinned:
                    vp.Pinned = False
                doc.Delete(vp.Id)

            new_viewports = [state.recreate(sheet) for state in states]

            # detail numbers must be unique on the sheet: park them on
            # temporary values first, then restore the originals
            for i, vp in enumerate(new_viewports):
                set_detail_number(vp, "ADA_TMP_{}".format(i))
            for vp, state in zip(new_viewports, states):
                set_detail_number(vp, state.detail_number)
                if state.pinned:
                    vp.Pinned = True
            t.Commit()
    except Exception as ex:
        report.error("Failed, nothing was changed: {}".format(ex))
        report.flush()
        return

    rows = [[str(i), state.name, state.detail_number or ""]
            for i, state in enumerate(reversed(states), 1)]
    report.subheader("Draw order (1 = on top)")
    report.table(["#", "View", "Detail Number"], rows)
    report.success("Draw order applied - {} viewport(s) re-placed at the same "
                   "position.".format(len(states)))
    report.flush()


if __name__ == "__main__":
    main()
