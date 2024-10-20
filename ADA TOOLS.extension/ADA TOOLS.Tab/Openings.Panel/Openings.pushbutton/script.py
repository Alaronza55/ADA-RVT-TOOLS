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