# -*- coding: utf-8 -*-
__title__ = "Toggle\nGrids"
__doc__ = "Toggle visibility of grids in the active view"

from pyrevit import revit, DB, script

def toggle_grid_visibility():
    """Toggle the visibility of grids in the active view"""
    
    doc = revit.doc
    active_view = doc.ActiveView
    
    # Check if we're in a valid view type
    if active_view.ViewType == DB.ViewType.Schedule or \
       active_view.ViewType == DB.ViewType.DrawingSheet or \
       active_view.ViewType == DB.ViewType.Legend:
        script.exit()
    
    # Get all grid categories
    grid_category = DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Grids)
    
    if not grid_category:
        print("Could not find Grids category")
        return
    
    # Get current visibility state
    try:
        current_state = active_view.GetCategoryHidden(grid_category.Id)
        new_state = not current_state
        
        # Start transaction
        t = DB.Transaction(doc, "Toggle Grid Visibility")
        t.Start()
        
        try:
            # Toggle the visibility
            active_view.SetCategoryHidden(grid_category.Id, new_state)
            t.Commit()
                
        except Exception as e:
            t.RollBack()
            print("Error toggling grid visibility: {}".format(str(e)))
            
    except Exception as e:
        print("Error accessing category visibility: {}".format(str(e)))

if __name__ == '__main__':
    toggle_grid_visibility()