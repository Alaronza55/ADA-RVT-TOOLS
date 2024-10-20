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
		return true

# class Single_Category(Selection.ISelectionFilter):
# 	def __init__(self, nom_categorie):
# 		self.nom_categorie = nom_categorie
# 	def AllowElement(self, e):
# 		if e.Category.Name == self.nom_categorie:
# 			return True
# 		else:
# 			return False
# 	def AllowReference(self, ref, point):
# 		return true

#SELECT SURFACE ELEMENT
# Prompt the user to select multiple elements
try:
    sel1 = uidoc.Selection.PickObjects(ObjectType.Element,  Mult_Category1('Structural Columns', 'Floors', 'Structural Framing','Walls'))
except Exception as e:
    print("Selection canceled.")
    selection = None

# Check if the user made a selection
if not sel1:
    print("No elements selected.")
else:
    # Create a list to store element IDs
    element_ids = []

    # Loop through the selected elements to retrieve their IDs
    for sel in sel1:
        # Get the element by its Reference
        element = doc.GetElement(sel.ElementId)
        if element:
            # Append the ElementId's IntegerValue (the actual ID number)
            element_ids.append(sel.ElementId.IntegerValue)

    # Print the selected element IDs
    print("Selected Element IDs:", element_ids)



#SELECT INTERSECTING ELEMENT
# sel2 = uidoc.Selection.PickObjects(Selection.ObjectType.Element,  Single_Category('Pipes'))

#PASS PARAMETER
# t = Transaction(doc, 'selection')
# t.Start()
# for i in els1:
# 	par1 = i.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
# 	par1.Set('Text_1')
# t.Commit()

# # Retrieve the geometry of the SURFACE ELEMENT
# def get_geometry_of_element(sel1):
#     options = Options()
#     geom_element = elem.get_Geometry(options)
    
#     # Extracting geometry info, if available
#     geometry_data = []
#     if geom_element is not None:
#         for geom_obj in geom_element:
#             # Here we retrieve basic geometry info (you can extract more detailed data)
#             geometry_data.append(str(geom_obj))
#     return geometry_data

# # Retrieve the geometry of the INTERSECTING ELEMENT
# def get_geometry_of_element(sel2):
#     options = Options()
#     geom_element = elem.get_Geometry(options)
    
#     # Extracting geometry info, if available
#     geometry_data = []
#     if geom_element is not None:
#         for geom_obj in geom_element:
#             # Here we retrieve basic geometry info (you can extract more detailed data)
#             geometry_data.append(str(geom_obj))
#     return geometry_data






