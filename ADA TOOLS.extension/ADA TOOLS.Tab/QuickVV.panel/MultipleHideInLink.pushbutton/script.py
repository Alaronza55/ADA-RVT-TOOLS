# -*- coding: utf-8 -*-
"""Click a linked model in the view, then pick elements inside it, then hide them."""

__title__ = "Hide Linked\nElements"
__author__ = "BESIX"
__doc__ = "Click a linked model in the view, pick elements inside it, hide them in the active view."

import sys
import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")
from System.Collections.Generic import List

from Autodesk.Revit import DB
from Autodesk.Revit.UI import Selection as UISelection
from pyrevit import revit, forms, script

doc         = revit.doc
uidoc       = revit.uidoc
active_view = doc.ActiveView
output      = script.get_output()

# ─── 1. Guard ──────────────────────────────────────────────────────────────────────────────
UNSUPPORTED = [
    DB.ViewType.Legend, DB.ViewType.Schedule,
    DB.ViewType.ColumnSchedule, DB.ViewType.PanelSchedule,
    DB.ViewType.DrawingSheet,
]
if active_view.ViewType in UNSUPPORTED:
    forms.alert("Active view does not support hiding elements.", exitscript=True)


# ─── 2. Selection filters ────────────────────────────────────────────────────────────
class LinkOnlyFilter(UISelection.ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, DB.RevitLinkInstance)
    def AllowReference(self, reference, point):
        return False


class LinkedElementFilter(UISelection.ISelectionFilter):
    def __init__(self, link_instance_id):
        self.link_instance_id = link_instance_id
    def AllowElement(self, element):
        return element.Id == self.link_instance_id
    def AllowReference(self, reference, point):
        return reference.ElementId == self.link_instance_id


# ─── 3. Step 1 — pick the link instance in the view ──────────────────────────────────────────────
try:
    link_ref = uidoc.Selection.PickObject(
        UISelection.ObjectType.Element,
        LinkOnlyFilter(),
        "Step 1/2 — Click on the linked model in the view."
    )
except Exception:
    sys.exit()

link_instance = doc.GetElement(link_ref.ElementId)
link_inst_id  = link_instance.Id
link_doc      = link_instance.GetLinkDocument()

if not link_doc:
    forms.alert("The selected link is not loaded.", exitscript=True)

output.print_md("**Link selected:** `{}`".format(link_instance.Name))


# ─── 4. Step 2 — pick elements inside that link ───────────────────────────────────────────────
try:
    refs = uidoc.Selection.PickObjects(
        UISelection.ObjectType.LinkedElement,
        LinkedElementFilter(link_inst_id),
        "Step 2/2 — Click elements to hide inside '{}' — press FINISH when done.".format(
            link_instance.Name
        )
    )
except Exception:
    sys.exit()

if not refs:
    forms.alert("No elements selected.", exitscript=True)

linked_elem_ids = []
for ref in refs:
    eid = ref.LinkedElementId
    if eid != DB.ElementId.InvalidElementId:
        linked_elem_ids.append(eid)

if not linked_elem_ids:
    forms.alert("Could not resolve any element IDs from the selection.", exitscript=True)

output.print_md("**Elements picked:** {}".format(len(linked_elem_ids)))


# ─── 5. Ensure link display is set to "Custom" (not "By Linked View") ──────────────────
# When a link is set to "By Linked View", per-element overrides and HideElements
# are completely ignored by Revit. We must switch the link to "Custom" display
# mode for the active view so that element-level visibility is respected.
#
# The API call is:  view.SetLinkOverrides(linkInstId, RevitLinkGraphicsSettings)
# RevitLinkGraphicsSettings.LinkVisibilityType controls By Linked View vs Custom.

def ensure_link_is_custom(view, link_inst_id):
    """
    If the link is set to ByLinkedView, switch it to Custom so that
    per-element overrides and HideElements() are respected.
    Returns True if a change was made (so caller can warn the user).
    """
    current = view.GetLinkOverrides(link_inst_id)
    if current is None:
        return False
    # LinkVisibilityType.ByLinkView == 1, Custom == 0
    if current.LinkVisibilityType == DB.RevitLinkGraphicsSettings.LinkVisibilityTypes.ByLinkView:
        custom_settings = DB.RevitLinkGraphicsSettings()
        # Default constructor creates a Custom setting
        view.SetLinkOverrides(link_inst_id, custom_settings)
        return True
    return False


# ─── 6. Hide elements ───────────────────────────────────────────────────────────────────────────────
def invisible_ogs():
    ogs = DB.OverrideGraphicSettings()
    ogs.SetSurfaceTransparency(100)
    white = DB.Color(255, 255, 255)
    ogs.SetProjectionLineColor(white)
    ogs.SetCutLineColor(white)
    ogs.SetSurfaceForegroundPatternColor(white)
    ogs.SetSurfaceBackgroundPatternColor(white)
    ogs.SetCutForegroundPatternColor(white)
    ogs.SetCutBackgroundPatternColor(white)
    return ogs

hidden_count  = 0
errors        = []
switched_mode = False

with revit.Transaction("Hide Linked Elements"):

    # Switch link to Custom display if currently set to By Linked View
    try:
        switched_mode = ensure_link_is_custom(active_view, link_inst_id)
        if switched_mode:
            output.print_md(
                "ℹ Link was set to **By Linked View** — switched to **Custom** "
                "so per-element visibility is respected."
            )
    except Exception as e:
        output.print_md("⚠ Could not check/set link display mode: `{}`".format(e))

    # Hide each element: try native HideElements, fallback to graphic override
    ogs = invisible_ogs()
    for eid in linked_elem_ids:
        try:
            active_view.HideElements(List[DB.ElementId]([eid]))
            hidden_count += 1
        except Exception:
            try:
                active_view.SetElementOverrides(eid, ogs)
                hidden_count += 1
            except Exception as e2:
                msg = "Element `{}`: {}".format(eid, e2)
                errors.append(msg)
                output.print_md("  ⚠ " + msg)

# ─── 7. Summary ───────────────────────────────────────────────────────────────────────────────────
output.print_md("\n---")
output.print_md("✅ **{} element(s) hidden** in view: *{}*".format(hidden_count, active_view.Name))
if switched_mode:
    output.print_md(
        "> ⚠ The link display was changed from **By Linked View** to **Custom** "
        "for this view. Other elements in the link will now follow the host view\'s "
        "Visibility/Graphics settings instead of the linked view."
    )
if errors:
    output.print_md("⚠ **{} error(s):**".format(len(errors)))
    for e in errors:
        output.print_md("- " + e)