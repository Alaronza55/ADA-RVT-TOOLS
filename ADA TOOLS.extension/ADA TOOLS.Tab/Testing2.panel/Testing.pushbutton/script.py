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

#----------------------------------------------------------------------

# def get_solids_from_element(element):
#     """Gets all solids from a Revit element"""
    
#     solids = []  # Initialize as an empty list
#     opts = Options()
#     opts.ComputeReferences = True
#     opts.DetailLevel = ViewDetailLevel.Fine
    
#     geometry = element.get_Geometry(opts)
    
#     if geometry:
#         for geo_obj in geometry:
#             if isinstance(geo_obj, Solid) and geo_obj.Volume > 0:
#                 solids.append(geo_obj)
#             elif isinstance(geo_obj, GeometryInstance):
#                 instance_geometry = geo_obj.GetInstanceGeometry()
#                 for instance_geo_obj in instance_geometry:
#                     if isinstance(instance_geo_obj, Solid) and instance_geo_obj.Volume > 0:
#                         solids.append(instance_geo_obj)
                    
#     return solids

# def analyze_face(face, index):
#     """Analyzes a face and returns information about it"""
#     face_info = {
#         "index": index,
#         "type": face.__class__.__name__,
#         "area": face.Area,
#         "edges": []
#     }
    
#     # Get face-specific properties
#     if isinstance(face, PlanarFace):
#         face_info["normal"] = face.FaceNormal  # Use FaceNormal instead of Normal
#         face_info["origin"] = face.Origin
#     elif isinstance(face, CylindricalFace):
#         face_info["axis"] = face.Axis
#         face_info["radius"] = face.Radius
    
#     # Get edge information
#     for edge in face.EdgeLoops.Item[0]:
#         edge_info = {
#             "type": edge.__class__.__name__,
#             "length": edge.AsCurve().Length,
#             "start": edge.AsCurve().GetEndPoint(0),
#             "end": edge.AsCurve().GetEndPoint(1)
#         }
#         face_info["edges"].append(edge_info)
    
#     return face_info




# def print_face_info(face_info):
#     """Prints face information in a structured format with metric units"""
#     indent = "    "
    
#     output.print_md("### Face #{} - {}".format(face_info["index"] + 1, face_info["type"]))
    
#     # Convert area from sq ft to sq m (1 sq ft = 0.092903 sq m)
#     area_in_sqm = face_info["area"] * 0.092903
#     output.print_md("**Area:** {:.4f} sq m".format(area_in_sqm))
    
#     # Print face-specific properties
#     if "normal" in face_info:
#         normal = face_info["normal"]
#         output.print_md("**Normal Vector:** ({:.4f}, {:.4f}, {:.4f})".format(normal.X, normal.Y, normal.Z))
    
#     if "origin" in face_info:
#         origin = face_info["origin"]
#         # Convert coordinates from ft to m (1 ft = 0.3048 m)
#         x_m = origin.X * 0.3048
#         y_m = origin.Y * 0.3048
#         z_m = origin.Z * 0.3048
#         output.print_md("**Origin Point:** ({:.4f} m, {:.4f} m, {:.4f} m)".format(x_m, y_m, z_m))
    
#     if "axis" in face_info:
#         axis = face_info["axis"]
#         output.print_md("**Axis Direction:** ({:.4f}, {:.4f}, {:.4f})".format(axis.X, axis.Y, axis.Z))
    
#     if "radius" in face_info:
#         # Make sure radius is a number before formatting
#         radius = face_info["radius"]
#         if isinstance(radius, (int, float)):
#             # Convert radius from ft to m
#             radius_m = radius * 0.3048
#             output.print_md("**Radius:** {:.4f} m".format(radius_m))
#         else:
#             output.print_md("**Radius:** {}".format(radius))
    
#     # Print edge information
#     output.print_md("**Edges:** {}".format(len(face_info["edges"])))
#     for i, edge in enumerate(face_info["edges"]):
#         # Convert length from ft to m
#         length_m = edge["length"] * 0.3048
#         output.print_md("{}Edge #{} - {} (Length: {:.4f} m)".format(
#             indent, i + 1, edge["type"], length_m))
        
#         start = edge["start"]
#         # Convert coordinates from ft to m
#         start_x_m = start.X * 0.3048
#         start_y_m = start.Y * 0.3048
#         start_z_m = start.Z * 0.3048
#         output.print_md("{}    Start: ({:.4f} m, {:.4f} m, {:.4f} m)".format(
#             indent, start_x_m, start_y_m, start_z_m))
        
#         end = edge["end"]
#         # Convert coordinates from ft to m
#         end_x_m = end.X * 0.3048
#         end_y_m = end.Y * 0.3048
#         end_z_m = end.Z * 0.3048
#         output.print_md("{}    End: ({:.4f} m, {:.4f} m, {:.4f} m)".format(
#             indent, end_x_m, end_y_m, end_z_m))


# # Get solids from all structural elements
# structural_solids = []
# for elem in selected_elems_structural:
#     structural_solids.extend(get_solids_from_element(elem))

# # Get solids from all MEP elements
# mep_solids = []
# for elem in selected_elems_MEP:
#     mep_solids.extend(get_solids_from_element(elem))

# def main():
#     # Find intersections
#     intersection_found = False
#     intersection_count = 0

#     for i, struct_solid in enumerate(structural_solids):
#         for j, mep_solid in enumerate(mep_solids):
#             try:
#                 # Execute boolean operation to find intersection
#                 intersection_solid = BooleanOperationsUtils.ExecuteBooleanOperation(
#                     struct_solid, mep_solid, BooleanOperationsType.Intersect)
                
#                 if intersection_solid and intersection_solid.Volume > 0:
#                     intersection_found = True
#                     intersection_count += 1
                    
#                     output.print_md("---")
#                     output.print_md("## Intersection #{} (Structural Solid #{} ∩ MEP Solid #{})".format(
#                         intersection_count, i+1, j+1))
#                     output.print_md("**Volume:** {:.4f} cubic ft".format(intersection_solid.Volume))
                    
#                     # Analyze faces of the intersection solid
#                     face_count = 0
#                     for face in intersection_solid.Faces:
#                         face_info = analyze_face(face, face_count)
#                         print_face_info(face_info)
#                         face_count += 1
                    
#                     if face_count == 0:
#                         output.print_md("⚠️ No faces found in the intersection solid.")
            
#             except Exception as e:
#                 output.print_md("⚠️ **Error in intersection calculation:** {}".format(str(e)))
#                 continue

#     if not intersection_found:
#         output.print_md("❌ **No intersection found between the selected elements.**")
#         output.print_md("The structural and MEP elements do not intersect geometrically.")

# # Call the main function at the end of your script
# if __name__ == "__main__":
#     try:
#         main()
#     except Exception as e:
#         print("Error: {}".format(str(e)))
#         print(traceback.format_exc())
