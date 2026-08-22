# -*- coding: utf-8 -*-
"""Toggle Base Points and Site Visibility
Toggles visibility of Project Base Point, Survey Point, and Site in the active view.
Note: Internal Origin cannot be controlled via visibility settings in the API.
"""
__title__ = 'Toggle\nBase Points'
__author__ = 'Your Name'

from pyrevit import revit, DB, forms

# Get active document and view
doc = revit.doc
active_view = doc.ActiveView

# Category IDs for the base points and site
# Note: Internal Origin is not a controllable category in Revit API
category_ids = [
    DB.BuiltInCategory.OST_ProjectBasePoint,  # Project Base Point
    DB.BuiltInCategory.OST_SharedBasePoint,   # Survey Point
    DB.BuiltInCategory.OST_Site               # Site
]

# Check current visibility state of Project Base Point to determine toggle action
# We'll use the first category as reference
pbp_category = DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_ProjectBasePoint)
current_visibility = active_view.GetCategoryHidden(pbp_category.Id)

# Start transaction
t = DB.Transaction(doc, "Toggle Base Points Visibility")
t.Start()

try:
    for cat_id in category_ids:
        category = DB.Category.GetCategory(doc, cat_id)
        if category:
            # Toggle visibility - if currently hidden (True), show it (False), and vice versa
            active_view.SetCategoryHidden(category.Id, not current_visibility)
    
    t.Commit()
    
    # Show feedback to user
    if current_visibility:
        forms.alert("Base points and Site are now VISIBLE in the active view.", 
                   title="Base Points & Site Visibility")
    else:
        forms.alert("Base points and Site are now HIDDEN in the active view.", 
                   title="Base Points & Site Visibility")

except Exception as e:
    t.RollBack()
    forms.alert("Error toggling base points visibility:\n{}".format(str(e)), 
               title="Error", exitscript=True)