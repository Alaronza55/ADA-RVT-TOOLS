# -*- coding: utf-8 -*-
__doc__ = """Lists worksets with element counts and allows deletion of elements in selected worksets."""
__title__ = 'Workset\nElements Delete'
__version__ = "Version 1.0"
__author__ = 'ADA'

import clr
import System
from System.Collections.Generic import List

# Import RevitAPI
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import (
    FilteredWorksetCollector, WorksetKind, ElementWorksetFilter,
    FilteredElementCollector, BuiltInCategory, ElementId, Transaction,
    BuiltInParameter, ElementFilter, WorksetId
)

# Import RevitAPIUI
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult

# Import forms
from pyrevit import forms, revit, script

# Shared ADA-Tools dark/gold themed report and picker (see lib/GUI/ReportTheme.py,
# lib/GUI/SelectFromDict.py)
from GUI.ReportTheme import ADAReport
from GUI.forms import select_from_dict

# Create a logger output
output = script.get_output()
output.close_others()

# Get the current document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Check if document is workshared
if not doc.IsWorkshared:
    forms.alert('This document is not workshared. Worksets are only available in workshared documents.', exitscript=True)

# Get user worksets (excluding family worksets)
worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()

workset_data = []
for workset in worksets:
    # Create a workset filter for this workset
    workset_filter = ElementWorksetFilter(workset.Id, False)  # False means to include the elements in the workset
    
    # Count elements in the workset
    element_count = FilteredElementCollector(doc).WherePasses(workset_filter).GetElementCount()
    
    # Add to the workset_data list
    workset_data.append({
        'id': workset.Id,
        'name': workset.Name,
        'count': element_count
    })

# Sort worksets by name
workset_data.sort(key=lambda x: x['name'])

# Create a list of options for the checklist
options = ["{} ({} elements)".format(w['name'], w['count']) for w in workset_data]

# Show a form with checkboxes to let the user select worksets
selected_options = select_from_dict(
    options,
    title=__title__,
    label="Select worksets to clean:",
    button_name="Select Worksets",
    version=__version__,
    SelectMultiple=True
)

# If nothing was selected or the dialog was cancelled, exit
if not selected_options:
    forms.alert('No worksets were selected. Operation cancelled.', exitscript=True)

# Get the selected workset IDs
selected_workset_ids = []
selected_workset_names = []
for option in selected_options:
    index = options.index(option)
    selected_workset_ids.append(workset_data[index]['id'])
    selected_workset_names.append(workset_data[index]['name'])

# Confirm with the user
confirmation_message = "You're about to delete all deletable elements in the following worksets:\n\n"
for option in selected_options:
    confirmation_message += "- {}\n".format(option)
confirmation_message += "\nThis operation cannot be undone. Do you want to proceed?"

result = TaskDialog.Show("Confirm Deletion",
                         confirmation_message,
                         TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No,
                         TaskDialogResult.No)

if result == TaskDialogResult.No:
    forms.alert('Operation cancelled.', exitscript=True)

# Report
report = ADAReport(__title__.replace(chr(10), " "))
report.subheader("Processing Selected Worksets")
report.table(["Workset"], [[name] for name in selected_workset_names])

# Get all elements in the selected worksets
all_elements = []

for workset_id in selected_workset_ids:
    workset_filter = ElementWorksetFilter(workset_id, False)
    elements = FilteredElementCollector(doc).WherePasses(workset_filter).ToElements()
    all_elements.extend(elements)

report.line("Found <b>{}</b> total elements in selected worksets".format(len(all_elements)))

# Start a transaction
t = Transaction(doc, "Delete Elements in Selected Worksets")
t.Start()

# Counter for deleted and non-deletable elements
deleted_count = 0
non_deletable_count = 0
protected_categories = set()
dependent_elements = set()
pinned_elements = set()
other_reasons = set()

# Try to delete each element
for element in all_elements:
    try:
        # Skip null elements
        if element is None or not element.IsValidObject:
            non_deletable_count += 1
            other_reasons.add("Invalid or null elements")
            continue
        
        # Get category name if exists
        category_name = element.Category.Name if element.Category else "No Category"
        
        # Check if element is pinned - using HasParameter instead of direct access
        try:
            is_pinned = False
            if element.HasParameter(BuiltInParameter.ELEMENT_IS_PINNED_PARAM):
                param = element.get_Parameter(BuiltInParameter.ELEMENT_IS_PINNED_PARAM)
                if param and param.AsInteger() == 1:
                    is_pinned = True
                    pinned_elements.add(category_name)
                    non_deletable_count += 1
                    continue
        except Exception:
            # If parameter check fails, just continue with deletion attempt
            pass
        
        # Skip elements in protected categories
        if category_name in ["Levels", "Grids", "Views", "View Templates", "Viewports", "Sheets"]:
            protected_categories.add(category_name)
            non_deletable_count += 1
            continue
            
        # Try to delete the element
        try:
            # Try to check if element can be deleted
            can_delete = False
            try:
                can_delete = element.CanBeDeleted(doc)
            except:
                # If CanBeDeleted fails, we'll try direct deletion
                can_delete = True
                
            if can_delete:
                doc.Delete(element.Id)
                deleted_count += 1
            else:
                dependent_elements.add(category_name)
                non_deletable_count += 1
        except Exception as e:
            other_reasons.add(str(e))
            non_deletable_count += 1
    except Exception as e:
        other_reasons.add(str(e))
        non_deletable_count += 1

# Commit the transaction
t.Commit()

# Show detailed results in output window
report.subheader("Deletion Results")
report.success("<b>{}</b> elements successfully deleted".format(deleted_count))
if non_deletable_count:
    report.warn("<b>{}</b> elements could not be deleted".format(non_deletable_count))

if protected_categories:
    report.subheader("Protected Categories")
    for cat in protected_categories:
        report.line(cat)

if dependent_elements:
    report.subheader("Elements With Dependencies")
    for cat in dependent_elements:
        report.line(cat)

if pinned_elements:
    report.subheader("Pinned Elements")
    for cat in pinned_elements:
        report.line(cat)

if other_reasons:
    report.subheader("Other Reasons")
    for reason in other_reasons:
        report.warn(reason)

report.flush()

# Show the results in a dialog
result_message = "Operation completed:\n\n"
result_message += "- {} elements deleted\n".format(deleted_count)
result_message += "- {} elements could not be deleted".format(non_deletable_count)
result_message += "\n\nMore details are available in the output window."
forms.alert(result_message)
