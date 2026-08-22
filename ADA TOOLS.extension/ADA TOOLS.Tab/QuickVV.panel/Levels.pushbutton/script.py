# -*- coding: utf-8 -*-
__title__ = "Toggle\nLevels"
__doc__ = "Toggle visibility of levels in the active view"

from pyrevit import revit, DB, script

def toggle_level_visibility():
    """Toggle the visibility of levels in the active view"""
    
    doc = revit.doc
    active_view = doc.ActiveView
    
    # Check if we're in a valid view type
    if active_view.ViewType == DB.ViewType.Schedule or \
       active_view.ViewType == DB.ViewType.DrawingSheet or \
       active_view.ViewType == DB.ViewType.Legend:
        script.exit()
    
    # Get all level categories
    level_category = DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Levels)
    
    if not level_category:
        print("Could not find Levels category")
        return
    
    # Get current visibility state
    try:
        current_state = active_view.GetCategoryHidden(level_category.Id)
        new_state = not current_state
        
        # Start transaction
        t = DB.Transaction(doc, "Toggle Level Visibility")
        t.Start()
        
        try:
            # Toggle the visibility
            active_view.SetCategoryHidden(level_category.Id, new_state)
            t.Commit()
                
        except Exception as e:
            t.RollBack()
            print("Error toggling level visibility: {}".format(str(e)))
            
    except Exception as e:
        print("Error accessing category visibility: {}".format(str(e)))

if __name__ == '__main__':
    toggle_level_visibility()