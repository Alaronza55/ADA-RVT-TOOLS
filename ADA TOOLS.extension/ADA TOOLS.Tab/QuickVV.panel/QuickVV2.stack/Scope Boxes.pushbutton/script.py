# -*- coding: utf-8 -*-
__title__ = "Toggle\nScope Boxes"
__doc__ = "Toggle visibility of scope boxes in the active view"

from pyrevit import revit, DB, script

# Shared ADA-Tools dark/gold themed report (see lib/GUI/ReportTheme.py)
from GUI.ReportTheme import ADAReport

def toggle_scope_box_visibility():
    """Toggle the visibility of scope boxes in the active view"""

    doc = revit.doc
    active_view = doc.ActiveView

    # Check if we're in a valid view type
    if active_view.ViewType == DB.ViewType.Schedule or \
       active_view.ViewType == DB.ViewType.DrawingSheet or \
       active_view.ViewType == DB.ViewType.Legend:
        script.exit()

    # Get all scope box categories
    scope_box_category = DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_VolumeOfInterest)

    if not scope_box_category:
        ADAReport(__title__.replace(chr(10), " ")).error("Could not find Scope Boxes category").flush()
        return

    # Get current visibility state
    try:
        current_state = active_view.GetCategoryHidden(scope_box_category.Id)
        new_state = not current_state

        # Start transaction
        t = DB.Transaction(doc, "Toggle Scope Box Visibility")
        t.Start()

        try:
            # Toggle the visibility
            active_view.SetCategoryHidden(scope_box_category.Id, new_state)
            t.Commit()

        except Exception as e:
            t.RollBack()
            ADAReport(__title__.replace(chr(10), " ")).error(
                "Error toggling scope box visibility: {}".format(str(e))).flush()

    except Exception as e:
        ADAReport(__title__.replace(chr(10), " ")).error(
            "Error accessing category visibility: {}".format(str(e))).flush()

if __name__ == '__main__':
    toggle_scope_box_visibility()