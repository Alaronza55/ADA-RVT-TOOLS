# -*- coding: utf-8 -*-
__title__ = "Set Opening\nHost to View"
__author__ = "Your Name"
__doc__ = """Changes the host of generic models containing 'BES_Opening' 
from linked models to the active view's level by recreating them with exact same position and parameters."""

from Autodesk.Revit.DB import (
    FilteredElementCollector, 
    BuiltInCategory,
    Transaction,
    FamilyInstance,
    Level,
    BuiltInParameter,
    XYZ,
    ElementId,
    Line,
    LocationPoint,
    StorageType
)
from Autodesk.Revit.DB.Structure import StructuralType
from pyrevit import revit, forms
import math

doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView

# Get the active view's associated level
view_level = active_view.GenLevel

if not view_level:
    forms.alert("Active view does not have an associated level.", exitscript=True)

# Collect all generic models in the active view
generic_models = FilteredElementCollector(doc, active_view.Id) \
    .OfCategory(BuiltInCategory.OST_GenericModel) \
    .WhereElementIsNotElementType() \
    .ToElements()

# Filter for elements containing "BES_Opening" in type or family name
matching_elements = []
for element in generic_models:
    if isinstance(element, FamilyInstance):
        type_name = element.Symbol.FamilyName if element.Symbol else ""
        family_name = element.Symbol.Family.Name if element.Symbol and element.Symbol.Family else ""
        
        if "BES_Opening" in type_name or "BES_Opening" in family_name:
            matching_elements.append(element)

if not matching_elements:
    forms.alert("No generic models found with 'BES_Opening' in type or family name in active view.", exitscript=True)

# Store element data before transaction
elements_to_change = []
elements_to_recreate = []

for element in matching_elements:
    # Check if level parameter is read-only
    level_param = element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
    if level_param and not level_param.IsReadOnly:
        elements_to_change.append({
            'element': element,
            'id': element.Id.IntegerValue
        })
        continue
    
    # Need to recreate
    element_data = {
        'id': element.Id,
        'id_int': element.Id.IntegerValue,
        'symbol': element.Symbol,
        'symbol_id': element.Symbol.Id,
        'can_recreate': False
    }
    
    location = element.Location
    if hasattr(location, 'Point'):
        point = location.Point
        element_data['target_point'] = XYZ(point.X, point.Y, point.Z)
        element_data['can_recreate'] = True
        
        # Get rotation
        angle = 0
        if isinstance(location, LocationPoint):
            try:
                transform = element.GetTransform()
                basis_x = transform.BasisX
                angle = math.atan2(basis_x.Y, basis_x.X)
            except:
                angle = 0
        element_data['angle'] = angle
        
        # Get flips
        try:
            element_data['hand_flipped'] = element.HandFlipped if hasattr(element, 'HandFlipped') else False
            element_data['facing_flipped'] = element.FacingFlipped if hasattr(element, 'FacingFlipped') else False
        except:
            element_data['hand_flipped'] = False
            element_data['facing_flipped'] = False
        
        # Store ALL parameters
        params_to_copy = {}
        for param in element.Parameters:
            try:
                param_name = param.Definition.Name
                
                try:
                    builtin_param = param.Id.IntegerValue
                except:
                    builtin_param = None
                
                param_info = {
                    'storage_type': param.StorageType,
                    'value': None,
                    'is_readonly': param.IsReadOnly,
                    'has_value': param.HasValue,
                    'builtin_param': builtin_param
                }
                
                if param.HasValue:
                    if param.StorageType == StorageType.Integer:
                        param_info['value'] = param.AsInteger()
                    elif param.StorageType == StorageType.Double:
                        param_info['value'] = param.AsDouble()
                    elif param.StorageType == StorageType.String:
                        param_info['value'] = param.AsString()
                    elif param.StorageType == StorageType.ElementId:
                        param_info['value'] = param.AsElementId()
                
                params_to_copy[param_name] = param_info
            except:
                pass
        
        element_data['parameters'] = params_to_copy
    
    if element_data['can_recreate']:
        elements_to_recreate.append(element_data)

changed_count = 0
recreated_count = 0
failed_count = 0
failed_elements = []
new_element_ids = []

# Parameters to skip
SKIP_PARAM_NAMES = ['Elevation from Level', 'Level', 'Offset from Host', 'Host']
SKIP_BUILTIN_IDS = [
    BuiltInParameter.FAMILY_LEVEL_PARAM.value__,
    BuiltInParameter.INSTANCE_ELEVATION_PARAM.value__,
    BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM.value__,
    BuiltInParameter.HOST_ID_PARAM.value__
]

# TRANSACTION 1: Create elements and copy non-position parameters
with Transaction(doc, "Create Elements") as t:
    t.Start()
    
    for elem_data in elements_to_change:
        try:
            element = elem_data['element']
            level_param = element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
            level_param.Set(view_level.Id)
            changed_count += 1
        except Exception as e:
            failed_count += 1
            failed_elements.append(elem_data['id'])
    
    for element_data in elements_to_recreate:
        element_id = element_data['id_int']
        try:
            symbol = doc.GetElement(element_data['symbol_id'])
            if symbol and not symbol.IsActive:
                symbol.Activate()
                doc.Regenerate()
            
            target_point = element_data['target_point']
            insertion_point = XYZ(target_point.X, target_point.Y, view_level.Elevation)
            
            doc.Delete(element_data['id'])
            doc.Regenerate()
            
            new_element = doc.Create.NewFamilyInstance(
                insertion_point,
                symbol,
                view_level,
                StructuralType.NonStructural
            )
            
            if new_element is None:
                raise Exception("NewFamilyInstance returned None")
            
            # Store new element ID and target point for second transaction
            new_element_ids.append({
                'element': new_element,
                'id': new_element.Id,
                'target_point': target_point,
                'angle': element_data['angle'],
                'hand_flipped': element_data['hand_flipped'],
                'facing_flipped': element_data['facing_flipped'],
                'parameters': element_data['parameters']
            })
            
            recreated_count += 1
                
        except Exception as e:
            failed_count += 1
            failed_elements.append(element_id)
            print("FAILED Element {}: {}".format(element_id, str(e)))
    
    t.Commit()

# TRANSACTION 2: Position and parameter adjustment
with Transaction(doc, "Position and Parameters") as t:
    t.Start()
    
    for new_elem_data in new_element_ids:
        try:
            new_element = doc.GetElement(new_elem_data['id'])
            target_point = new_elem_data['target_point']
            
            # Copy non-position parameters FIRST
            for param_name, param_data in new_elem_data['parameters'].items():
                try:
                    if param_data['is_readonly']:
                        continue
                    
                    if param_name in SKIP_PARAM_NAMES:
                        continue
                    
                    builtin_param_id = param_data.get('builtin_param')
                    if builtin_param_id in SKIP_BUILTIN_IDS:
                        continue
                    
                    new_param = new_element.LookupParameter(param_name)
                    if new_param and not new_param.IsReadOnly and param_data['has_value'] and param_data['value'] is not None:
                        if param_data['storage_type'] == StorageType.Integer:
                            new_param.Set(param_data['value'])
                        elif param_data['storage_type'] == StorageType.Double:
                            new_param.Set(param_data['value'])
                        elif param_data['storage_type'] == StorageType.String:
                            if param_data['value']:
                                new_param.Set(param_data['value'])
                        elif param_data['storage_type'] == StorageType.ElementId:
                            if param_data['value'] != ElementId.InvalidElementId:
                                new_param.Set(param_data['value'])
                except:
                    pass
            
            # NOW move to target position
            current_location = new_element.Location
            if hasattr(current_location, 'Point'):
                current_point = current_location.Point
                translation = XYZ(0, 0, target_point.Z - current_point.Z)
                current_location.Move(translation)
                
                # Verify
                final_point = new_element.Location.Point
                print("Element {}: Target Z={}, Final Z={}, Diff={}".format(
                    new_elem_data['id'].IntegerValue, target_point.Z, final_point.Z, abs(final_point.Z - target_point.Z)))
            
            # Apply rotation
            if new_elem_data['angle'] != 0:
                try:
                    loc = new_element.Location
                    if hasattr(loc, 'Point'):
                        pt = loc.Point
                        axis = Line.CreateBound(pt, XYZ(pt.X, pt.Y, pt.Z + 10))
                        loc.Rotate(axis, new_elem_data['angle'])
                except:
                    pass
            
            # Apply flips
            try:
                if new_elem_data['hand_flipped'] and hasattr(new_element, 'HandFlipped'):
                    new_element.HandFlipped = True
                if new_elem_data['facing_flipped'] and hasattr(new_element, 'FacingFlipped'):
                    new_element.FacingFlipped = True
            except:
                pass
                
        except Exception as e:
            print("Position adjustment failed: {}".format(str(e)))
    
    t.Commit()

# Report results
message = "Host Level Change Complete:\n\n"
message += "Successfully changed: {} elements\n".format(changed_count)
message += "Recreated: {} elements\n".format(recreated_count)
message += "Failed: {} elements\n\n".format(failed_count)
message += "New host level: {}".format(view_level.Name)

if failed_elements:
    message += "\n\nFailed element IDs (first 10):\n{}".format(", ".join(map(str, failed_elements[:10])))

forms.alert(message, title="Set Opening Host Results")