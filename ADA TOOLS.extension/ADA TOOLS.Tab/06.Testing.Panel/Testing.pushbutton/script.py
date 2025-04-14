import sys
from Autodesk.Revit.DB import *
from pyrevit import forms, revit


doc     =   __revit__.ActiveUIDocument.Document
uidoc   =   __revit__.ActiveUIDocument
app     =   __revit__.Application


with forms.WarningBar(title='Pick Element:'):
    element = revit.pick_element()
    
element_type = type(element)

print(element_type)
if element_type != Wall:
    forms.alert('You were supposed to pick a Wall.', exitscript=True)

e_cat                   =         element.Category.Name
e_id                    =         element.Id
e_level_id              =         element.LevelId
e_wall_type             =         element.WallType
e_width                 =         element.Width
get_level               =         doc.GetElement(e_level_id)
e_level_name            =         get_level.Name

print('Element Category:          {}'.format(e_cat       ))
print('Element ID:                {}'.format(e_id        ))
print('Level ID of Element :      {}'.format(e_level_id  ))
print('Level Name of Element :    {}'.format(e_level_name))
print('Wall Type :                {}'.format(e_wall_type ))
print('Element Width :            {}'.format(e_width     )) 