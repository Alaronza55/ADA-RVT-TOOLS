# -*- coding: utf-8 -*-
__doc__ = """Select an element by pasting text that contains its Element ID in
brackets, e.g. 'Pipe Insulation [3290719]' (as copied from a schedule
or the status bar). The element is selected in the active view without
switching views. Perfect for copy paste from Navisworks!'"""
__title__ = "[ID]Isolate"
__author__ = "ADA"

import re
import System
from pyrevit import forms, revit, script
from Autodesk.Revit.DB import ElementId

doc = revit.doc
uidoc = revit.uidoc

# Ask user for input
user_input = forms.ask_for_string(
    prompt="Paste the element text (e.g. 'Pipe Insulation [3290719]'):",
    title="Select Element by Text"
)

if not user_input:
    script.exit()

# Extract the ID between brackets
match = re.search(r'\[(\d+)\]', user_input)

if not match:
    forms.alert("No element ID found between brackets.", title="Error")
    script.exit()

element_id_int = int(match.group(1))

# Get the element
element_id = ElementId(long(element_id_int))
element = doc.GetElement(element_id)

if not element:
    forms.alert("Element ID {} not found in the model.".format(element_id_int), title="Not Found")
    script.exit()

# Select the element in active view without switching views
selection = uidoc.Selection
selection.SetElementIds(System.Collections.Generic.List[ElementId]([element_id]))

forms.alert("Selected: {}".format(user_input.strip()), title="Done", warn_icon=False)