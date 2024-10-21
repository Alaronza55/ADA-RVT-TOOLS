# -*- coding: utf-8 -*-
__title__ = "Create Openings Tool"   # Name of the button displayed in Revit
__author__ = "Almog Davodson"
__doc__ = """Version = 1.0
_____________________________________________________________________
Description:

Create 3D Generic Model at the intersection between a Surface 
(Horizontal or Vertical) and a Pipe. 

Future Features / Roadmap :
- Make same process between two elements in link
- Integrate Conducts - HVAC Ducts
- Set Offset to Generic Model if needed.
- Make auto caracerization of opening
_____________________________________________________________________
How-to:

1. Select Element of surface
2. Select Pipe

_____________________________________________________________________
Last update:
- [20.10.2024] - 1.0 First coding attempt
_____________________________________________________________________
Author:     Almog Davidson"""


#IMPORTS
import os
import sys
import Autodesk
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import revit, forms

#VARIABLES
uidoc     = __revit__.ActiveUIDocument
doc       = __revit__.ActiveUIDocument.Document #type: Document

# CLASS
class Mult_Category1(Selection.ISelectionFilter):
	def __init__(self, nom_categorie_1, nom_categorie_2, nom_categorie_3, nom_categorie_4):
		self.nom_categorie_1 = nom_categorie_1
		self.nom_categorie_2 = nom_categorie_2
		self.nom_categorie_3 = nom_categorie_3
		self.nom_categorie_4 = nom_categorie_4

	def AllowElement(self, e):
		if e.Category.Name == self.nom_categorie_1 or e.Category.Name == self.nom_categorie_2 or e.Category.Name == self.nom_categorie_3 or e.Category.Name == self.nom_categorie_4:
			return True
		else:
			return False
	def AllowReference(self, ref, point):
		return True

class Mult_Category2(Selection.ISelectionFilter):
	def __init__(self, nom_categorie_1, nom_categorie_2):
		self.nom_categorie_1 = nom_categorie_1
		self.nom_categorie_2 = nom_categorie_2

	def AllowElement(self, e):
		if e.Category.Name == self.nom_categorie_1 or e.Category.Name == self.nom_categorie_2 :
			return True
		else:
			return False
	def AllowReference(self, ref, point):
		return True

#SELECT SURFACE ELEMENT
# Prompt the user to select multiple elements
try:
    selection1 = uidoc.Selection.PickObjects(ObjectType.Element,  Mult_Category1('Structural Columns', 'Floors', 'Structural Framing','Walls'))
except Exception as e:
    print("Selection canceled.")
    selection = None

# Check if the user made a selection
if not selection1:
    print("No elements selected.")
else:
    # Create a list to store element IDs
    element_ids_1 = []
    elements_1 = []

    # Loop through the selected elements to retrieve their IDs
    for sel in selection1:
        # Get the element by its Reference
        element = doc.GetElement(sel.ElementId)
        if element:
            # Append the ElementId's IntegerValue (the actual ID number)
            element_ids_1.append(sel.ElementId.IntegerValue)

#SELECT INTERSECTING ELEMENT
# Second selection: Prompt the user to select a second set of elements from allowed categories
try:
    selection2 = uidoc.Selection.PickObjects(ObjectType.Element, Mult_Category2('Pipes','Ducts'))
except Exception as e:
    print("Second selection canceled.")
    selection2 = None

# Check if the user made a second selection
if not selection2:
    print("No elements selected for the second selection.")
else:
    # Create a list to store the second set of element IDs
    element_ids_2 = []
    elements_2 = []

    # Loop through the second set of selected elements to retrieve their IDs
    for sel in selection2:
        # Get the element by its Reference
        element = doc.GetElement(sel.ElementId)
        if element:
            # Append the ElementId's IntegerValue (the actual ID number)
            element_ids_2.append(sel.ElementId.IntegerValue)

    # Print the second set of selected element IDs
    print("Selected Element IDs:", element_ids_1)
    print("Second set of selected Element IDs:", element_ids_2)



#PASS PARAMETER
# t = Transaction(doc, 'selection')
# t.Start()
# for i in els1:
# 	par1 = i.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
# 	par1.Set('Text_1')
# t.Commit()

