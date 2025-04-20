# -*- coding: utf-8 -*-
__title__   = "Testing2"
__version__ = 'Version = 0.2 (Beta)'
__doc__ = """Date    = 20.04.2024
_____________________________________________________________________
Description:
...

_____________________________________________________________________
Author: Erik Frits"""
# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#==================================================
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import *
import traceback


# pyRevit Imports
from pyrevit import script, forms, EXEC_PARAMS

# Custom EF Imports
from Snippets._vectors import rotate_vector
from Snippets._views import SectionGenerator
from GUI.forms import select_from_dict

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#==================================================
uidoc     = __revit__.ActiveUIDocument
doc       = __revit__.ActiveUIDocument.Document #type: Document

output = script.get_output()

# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝
#----------------------------------------------------------------------
# 1️⃣ User Input - Select Structural Element
bic_structural = BuiltInCategory
select_opts_structural = {'Walls'             : bic_structural.OST_Walls,
              'Columns'            : [bic_structural.OST_Columns, bic_structural.OST_StructuralColumns],
              'Beams/Framing'      : bic_structural.OST_StructuralFraming,
              'Floors'             : bic_structural.OST_Floors,
              }

#👉 Pick Selection Opts
selected_opts_structural = select_from_dict(select_opts_structural,
                            title = __title__,
                            label = 'Select Structural Element',
                            version = __version__,
                            SelectMultiple = False)


#⬇️ Flatten List to break nested lists
def flatten_list(lst1):
    new_lst_structural = []
    for i in lst1:
        if isinstance(i,list):
            new_lst_structural += i
        else:
            new_lst_structural.append(i)
    return new_lst_structural

selected_opts_structural = flatten_list(selected_opts_structural)

if not selected_opts_structural:
    forms.alert('No Category was selected. Please Try Again.', title=__title__, exitscript=True)

#----------------------------------------------------------------------
#2️⃣ Select Elements

class ADA_SelectionFilter_Structural(ISelectionFilter):
    def __init__(self, list_types_or_cats):
        """ ISelectionFilter made to filter with types
        :param allowed_types: list of allowed Types"""

        # Convert BuiltInCategories to ElementIds, Keep Types the Same.
        self.list_types_or_cats = [ElementId(i) if type(i) == BuiltInCategory else i for i in list_types_or_cats]

    def AllowElement(self, element_structural):
        if element_structural.ViewSpecific:
            return False

        #🅰️ Check if Element's Type in Allowed List
        if type(element_structural) in self.list_types_or_cats:
            return True

        #🅱️ Check if Element's Category in Allowed List
        elif element_structural.Category.Id in self.list_types_or_cats:
            return True

selected_elems_structural = []

try:
    ISF = ADA_SelectionFilter_Structural(selected_opts_structural)
    with forms.WarningBar(title='Select Elements and click "Finish"'):
        ref_selected_elems_structural = uidoc.Selection.PickObjects(ObjectType.Element,ISF)

    selected_elems_structural         = [doc.GetElement(ref_str) for ref_str in ref_selected_elems_structural]

except:
    if EXEC_PARAMS.debug_mode:
        import traceback

        print(traceback.format_exc())

#----------------------------------------------------------------------


# 1️⃣ User Input - Select MEP Element
bic_MEP = BuiltInCategory
select_opts_MEP = {'Plumbing Fixtures'  : [bic_MEP.OST_Furniture, bic_MEP.OST_PlumbingFixtures],
              'Lighting Fixtures'  : bic_MEP.OST_LightingFixtures,
              'Electrical Fixtures, Equipment, Circuits' : [bic_MEP.OST_ElectricalFixtures, bic_MEP.OST_ElectricalEquipment, bic_MEP.OST_ElectricalCircuit],
              'Pipes'              : bic_MEP.OST_PipeCurves,
              'Pipe Fittings'      : bic_MEP.OST_PipeFitting,
              'Pipe Accessories'   : bic_MEP.OST_PipeAccessory,
              'Ducts'              : bic_MEP.OST_DuctCurves,
              'Duct Fittings'      : bic_MEP.OST_DuctFitting,
              'Duct Accessories'   : bic_MEP.OST_DuctAccessory,
              'Cable Trays'        : bic_MEP.OST_CableTray,
              'Conduits'           : bic_MEP.OST_Conduit,
              'Mechanical Equipment': bic_MEP.OST_MechanicalEquipment
              }

#👉 Pick Selection Opts
selected_opts_MEP = select_from_dict(select_opts_MEP,
                            title = __title__,
                            label = 'Select MEP Element',
                            version = __version__,
                            SelectMultiple = False)


#⬇️ Flatten List to break nested lists
def flatten_list(lst2):
    new_lst_MEP = []
    for i in lst2:
        if isinstance(i,list):
            new_lst_MEP += i
        else:
            new_lst_MEP.append(i)
    return new_lst_MEP

selected_opts_MEP = flatten_list(selected_opts_MEP)

if not selected_opts_MEP:
    forms.alert('No Category was selected. Please Try Again.', title=__title__, exitscript=True)

#----------------------------------------------------------------------

class ADA_SelectionFilter_MEP(ISelectionFilter):
    def __init__(self, list_types_or_cats_MEP):
        """ ISelectionFilter made to filter with types
        :param allowed_types: list of allowed Types"""

        # Convert BuiltInCategories to ElementIds, Keep Types the Same.
        self.list_types_or_cats_MEP = [ElementId(c) if type(c) == BuiltInCategory else c for c in list_types_or_cats_MEP]

    def AllowElement(self, element_MEP):
        if element_MEP.ViewSpecific:
            return False

        #🅰️ Check if Element's Type in Allowed List
        if type(element_MEP) in self.list_types_or_cats_MEP:
            return True

        #🅱️ Check if Element's Category in Allowed List
        elif element_MEP.Category.Id in self.list_types_or_cats_MEP:
            return True

selected_elems_MEP = []

try:
    ISF_MEP = ADA_SelectionFilter_MEP(selected_opts_MEP)
    with forms.WarningBar(title='Select Elements and click "Finish"'):
        ref_selected_elems_MEP = uidoc.Selection.PickObjects(ObjectType.Element,ISF_MEP)

    selected_elems_MEP                = [doc.GetElement(ref_mep) for ref_mep in ref_selected_elems_MEP]

except:
    if EXEC_PARAMS.debug_mode:
        import traceback

        print(traceback.format_exc())

#----------------------------------------------------------------------
#3️⃣ Ensure Elements Selected, Exit if Not
if not selected_elems_structural:
    error_msg = 'No Elements were selected.\nPlease Try Again'
    forms.alert(error_msg, title='Selection has Failed.', exitscript=True)

if not selected_elems_MEP:
    error_msg = 'No Elements were selected.\nPlease Try Again'
    forms.alert(error_msg, title='Selection has Failed.', exitscript=True)
#----------------------------------------------------------------------
print("Structural Element ID:")
for i, elem in enumerate(selected_elems_structural):
    elem_id_structural = elem.Id.IntegerValue
    print(elem_id_structural)

print("MEP Element ID:")
for i, elem in enumerate(selected_elems_MEP):
    elem_id_MEP = elem.Id.IntegerValue
    print(elem_id_MEP)