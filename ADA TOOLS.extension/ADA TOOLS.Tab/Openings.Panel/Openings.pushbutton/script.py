# -*- coding: utf-8 -*-
__title__ = "Create Openings Tool"   # Name of the button displayed in Revit
__author__ = "Almog Davodson"
__doc__ = """Version = 1.0
_____________________________________________________________________
Description:

Create 3D Generic Model at the intersection between a Surface (Horizontal or Vertical)
and Pipe. 

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

#VARIABLES
uidoc     = __revit__.ActiveUIDocument
doc       = __revit__.ActiveUIDocument.Document #type: Document

# CLASS
class Mult_Category(Selection.ISelectionFilter):
	def __init__(self, nom_categorie_1, nom_categorie_2):
		self.nom_categorie_1 = nom_categorie_1
		self.nom_categorie_2 = nom_categorie_2
	def AllowElement(self, e):
		if e.Category.Name == self.nom_categorie_1 or e.Category.Name == self.nom_categorie_2:
			return True
		else:
			return False
	def AllowReference(self, ref, point):
		return true

#SELECT SURFACE ELEMENT
sel1 = uidoc.Selection.PickObject(Selection.ObjectType.Element, 'Choose Element')
el1 = doc.GetElement(sel1)

t = Transaction(doc, 'selection')
t.Start()
par1 = el1.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
par1.Set('')
t.Commit()