# -*- coding: utf-8 -*-
__title__ = "Toggle\nLines"
__doc__ = "Toggle visibility of lines in the active view"

from pyrevit import revit, DB, script

# Shared ADA-Tools dark/gold themed report (see lib/GUI/ReportTheme.py)
from GUI.ReportTheme import ADAReport

def toggle_lines_visibility():
    """Toggle the visibility of lines in the active view"""

    doc = revit.doc
    active_view = doc.ActiveView

    # Check if we're in a valid view type
    if active_view.ViewType == DB.ViewType.Schedule or \
       active_view.ViewType == DB.ViewType.DrawingSheet or \
       active_view.ViewType == DB.ViewType.Legend:
        script.exit()

    # Get all lines categories
    lines_category = DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Lines)

    if not lines_category:
        ADAReport(__title__.replace(chr(10), " ")).error("Could not find Lines category").flush()
        return

    # Get current visibility state
    try:
        current_state = active_view.GetCategoryHidden(lines_category.Id)
        new_state = not current_state

        # Start transaction
        t = DB.Transaction(doc, "Toggle Lines Visibility")
        t.Start()

        try:
            # Toggle the visibility
            active_view.SetCategoryHidden(lines_category.Id, new_state)
            t.Commit()

        except Exception as e:
            t.RollBack()
            ADAReport(__title__.replace(chr(10), " ")).error(
                "Error toggling lines visibility: {}".format(str(e))).flush()

    except Exception as e:
        ADAReport(__title__.replace(chr(10), " ")).error(
            "Error accessing category visibility: {}".format(str(e))).flush()

if __name__ == '__main__':
    toggle_lines_visibility()