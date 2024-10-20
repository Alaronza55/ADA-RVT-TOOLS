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


#Imports
import sys
from Autodesk.Revit.DB import (Transaction)

#Variables
doc = __revit__.ActiveUIDocument.Document
view = doc.ActiveView
t = Transaction(doc,'change scale')

#Main
t.Start()

view.Scale = 50

t.Commit()