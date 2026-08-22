# -*- coding: utf-8 -*-
__title__ = "Transfer to BES Parameters"
__doc__ = "Transfers Volume, Area, Assembly Code, Assembly Description, Family Name, Material, Category, Level, Type Mark, and Weight to BES_DATA EXTRACTION parameters for visible elements"

from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInParameter, BuiltInCategory, FamilyInstance, Options, Solid
from pyrevit import revit, forms

doc = revit.doc
uidoc = revit.uidoc

def get_parameter_value(element, built_in_param):
    """Get the value of a built-in parameter"""
    param = element.get_Parameter(built_in_param)
    if param and param.HasValue:
        if param.StorageType.ToString() == "Double":
            return param.AsDouble()
        elif param.StorageType.ToString() == "String":
            return param.AsString()
        elif param.StorageType.ToString() == "Integer":
            return param.AsInteger()
        elif param.StorageType.ToString() == "ElementId":
            return param.AsValueString()
    return None

def get_type_parameter_value(element, built_in_param):
    """Get the value of a built-in parameter from the element's type"""
    try:
        element_type = doc.GetElement(element.GetTypeId())
        if element_type:
            param = element_type.get_Parameter(built_in_param)
            if param and param.HasValue:
                if param.StorageType.ToString() == "Double":
                    return param.AsDouble()
                elif param.StorageType.ToString() == "String":
                    return param.AsString()
                elif param.StorageType.ToString() == "Integer":
                    return param.AsInteger()
                elif param.StorageType.ToString() == "ElementId":
                    return param.AsValueString()
    except:
        pass
    return None

def get_type_parameter_by_name(element, param_name):
    """Get a type parameter value by its name"""
    try:
        element_type = doc.GetElement(element.GetTypeId())
        if element_type:
            param = element_type.LookupParameter(param_name)
            if param and param.HasValue:
                if param.StorageType.ToString() == "Double":
                    return param.AsDouble()
                elif param.StorageType.ToString() == "String":
                    return param.AsString()
                elif param.StorageType.ToString() == "Integer":
                    return param.AsInteger()
                elif param.StorageType.ToString() == "ElementId":
                    return param.AsValueString()
    except:
        pass
    return None

def set_parameter_value(element, param_name, value):
    """Set the value of a parameter by name"""
    param = element.LookupParameter(param_name)
    if param and not param.IsReadOnly:
        if value is not None:
            if param.StorageType.ToString() == "Double":
                param.Set(float(value))
                return True
            elif param.StorageType.ToString() == "String":
                param.Set(str(value))
                return True
            elif param.StorageType.ToString() == "Integer":
                param.Set(int(value))
                return True
    return False

def get_family_name(element):
    """Get the family name of an element, or category name for system families"""
    try:
        # First check if element has a family
        if hasattr(element, 'Symbol') and element.Symbol:
            family_name = element.Symbol.Family.Name
            # If family name exists and is not empty, return it
            if family_name and family_name.strip():
                return family_name
        elif hasattr(element, 'Family'):
            family_name = element.Family.Name
            if family_name and family_name.strip():
                return family_name
    except:
        pass
    
    # If no family name found, return category name (for system families)
    try:
        if element.Category and element.Category.Name:
            return element.Category.Name
    except:
        pass
    
    return None

def get_category_name(element):
    """Get the category name of an element"""
    try:
        if element.Category and element.Category.Name:
            return element.Category.Name
    except:
        pass
    return None

def has_visible_area_parameter(element):
    """Check if element has a user-visible 'Area' parameter"""
    try:
        # Look for a parameter named "Area" that is visible to the user
        area_param = element.LookupParameter("Area")
        if area_param:
            # Check if it's not a hidden parameter and has a definition
            try:
                # If we can access the definition name, it's likely user-visible
                if area_param.Definition and area_param.Definition.Name == "Area":
                    return True
            except:
                pass
    except:
        pass
    return False

def calculate_geometry_volume(element):
    """Calculate volume from element's solid geometry (for stairs, railings, etc.)"""
    try:
        options = Options()
        options.ComputeReferences = True
        options.DetailLevel = 0  # Coarse detail level
        
        geom_element = element.get_Geometry(options)
        
        if not geom_element:
            return None
        
        total_volume = 0.0
        
        # Iterate through geometry objects
        for geom_obj in geom_element:
            # Check if it's a solid
            if isinstance(geom_obj, Solid):
                if geom_obj.Volume > 0:
                    total_volume += geom_obj.Volume
            # Check if it's a geometry instance
            elif hasattr(geom_obj, 'GetInstanceGeometry'):
                instance_geom = geom_obj.GetInstanceGeometry()
                if instance_geom:
                    for inst_obj in instance_geom:
                        if isinstance(inst_obj, Solid):
                            if inst_obj.Volume > 0:
                                total_volume += inst_obj.Volume
        
        if total_volume > 0:
            return total_volume
        
    except:
        pass
    
    return None

def get_element_volume(element):
    """Get volume for an element - includes nested family volumes"""
    total_volume = 0.0
    
    # First get the main element's volume
    volume = get_parameter_value(element, BuiltInParameter.HOST_VOLUME_COMPUTED)
    if volume is not None:
        total_volume += volume
    
    # If this is a FamilyInstance, add volumes from nested families
    if isinstance(element, FamilyInstance):
        try:
            # Get all sub-components (nested family instances)
            sub_component_ids = element.GetSubComponentIds()
            if sub_component_ids:
                for sub_id in sub_component_ids:
                    sub_element = doc.GetElement(sub_id)
                    if sub_element:
                        # Get volume from nested family
                        sub_volume = get_parameter_value(sub_element, BuiltInParameter.HOST_VOLUME_COMPUTED)
                        if sub_volume is not None:
                            total_volume += sub_volume
        except:
            pass
    
    # If still no volume, try calculating from geometry (for stairs, railings, etc.)
    if total_volume == 0.0:
        if element.Category:
            category_name = element.Category.Name
            # Categories that typically need geometry calculation
            if category_name in ["Stairs", "Railings", "Ramps"]:
                calculated_volume = calculate_geometry_volume(element)
                if calculated_volume is not None:
                    total_volume = calculated_volume
    
    # Return total volume if we found any
    if total_volume > 0.0:
        return total_volume
    
    return None

def get_element_weight(element):
    """Get weight for an element - includes nested family weights"""
    total_weight = 0.0
    
    # First get the main element's weight using LookupParameter
    weight_param = element.LookupParameter("Weight")
    if weight_param and weight_param.HasValue:
        total_weight += weight_param.AsDouble()
    
    # If this is a FamilyInstance, add weights from nested families
    if isinstance(element, FamilyInstance):
        try:
            # Get all sub-components (nested family instances)
            sub_component_ids = element.GetSubComponentIds()
            if sub_component_ids:
                for sub_id in sub_component_ids:
                    sub_element = doc.GetElement(sub_id)
                    if sub_element:
                        # Get weight from nested family
                        sub_weight_param = sub_element.LookupParameter("Weight")
                        if sub_weight_param and sub_weight_param.HasValue:
                            total_weight += sub_weight_param.AsDouble()
        except:
            pass
    
    # Return total weight if we found any
    if total_weight > 0.0:
        return total_weight
    
    return None

def get_materials_from_element(element, materials_dict):
    """Extract materials from an element and add to materials_dict"""
    try:
        # Get materials from GetMaterialIds (geometry materials)
        material_ids = element.GetMaterialIds(False)
        if material_ids and material_ids.Count > 0:
            for mat_id in material_ids:
                material = doc.GetElement(mat_id)
                if material:
                    materials_dict[material.Name] = True
        
        # Get structural material
        struct_material_param = element.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
        if struct_material_param and struct_material_param.HasValue:
            material_id = struct_material_param.AsElementId()
            if material_id and material_id.IntegerValue != -1:
                material = doc.GetElement(material_id)
                if material:
                    materials_dict[material.Name] = True
        
        # Get material from MATERIAL_ID_PARAM
        material_id_param = element.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
        if material_id_param and material_id_param.HasValue:
            material_id = material_id_param.AsElementId()
            if material_id and material_id.IntegerValue != -1:
                material = doc.GetElement(material_id)
                if material:
                    materials_dict[material.Name] = True
    except:
        pass

def get_material_name(element):
    """Get the material name(s) of an element including nested families"""
    materials_dict = {}
    
    try:
        # Get materials from the main element
        get_materials_from_element(element, materials_dict)
        
        # If this is a FamilyInstance, check for nested families (sub-components)
        if isinstance(element, FamilyInstance):
            try:
                # Get all sub-components (nested family instances)
                sub_component_ids = element.GetSubComponentIds()
                if sub_component_ids:
                    for sub_id in sub_component_ids:
                        sub_element = doc.GetElement(sub_id)
                        if sub_element:
                            get_materials_from_element(sub_element, materials_dict)
            except:
                pass
        
        # If no materials found yet, check type material
        if not materials_dict:
            element_type = doc.GetElement(element.GetTypeId())
            if element_type:
                type_material_param = element_type.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
                if type_material_param and type_material_param.HasValue:
                    material_id = type_material_param.AsElementId()
                    if material_id and material_id.IntegerValue != -1:
                        material = doc.GetElement(material_id)
                        if material:
                            materials_dict[material.Name] = True
        
        # Return all materials joined with " / "
        if materials_dict:
            return " / ".join(materials_dict.keys())
    
    except:
        pass
    
    return None

def get_project_base_point_elevation():
    """Get the project base point elevation in meters"""
    try:
        collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint)
        for pbp in collector:
            elev_param = pbp.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM)
            if elev_param and elev_param.HasValue:
                # Return elevation in meters
                return elev_param.AsDouble() * 0.3048
    except:
        pass
    return 0.0  # Default to 0 if not found

def get_element_level_name(element):
    """
    METHOD 2: Get the level name from the element's level parameter.
    Tries multiple level parameters depending on element type.
    Returns None if no level parameter is found.
    """
    try:
        # List of possible level parameters to check (in priority order)
        level_params = [
            BuiltInParameter.ROOF_BASE_LEVEL_PARAM,            # Roof Base Level
            BuiltInParameter.FAMILY_BASE_LEVEL_PARAM,          # Base Level (columns, framing)
            BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,   # Reference Level (structural framing)
            BuiltInParameter.SCHEDULE_LEVEL_PARAM,             # Level (walls, floors, etc.)
            BuiltInParameter.FAMILY_LEVEL_PARAM,               # Level (some family instances)
            BuiltInParameter.WALL_BASE_CONSTRAINT,             # Wall Base Constraint
            BuiltInParameter.LEVEL_PARAM,                      # Level parameter
            BuiltInParameter.STAIRS_BASE_LEVEL_PARAM,          # Stairs base level
        ]
        
        for level_param_id in level_params:
            level_param = element.get_Parameter(level_param_id)
            if level_param and level_param.HasValue:
                level_id = level_param.AsElementId()
                if level_id and level_id.IntegerValue != -1:
                    level = doc.GetElement(level_id)
                    if level and hasattr(level, 'Name'):
                        return level.Name
        
    except:
        pass
    
    return None

def get_element_level_designation_by_elevation(element, base_elevation):
    """
    METHOD 1 (Fallback): Get the level designation based on element's absolute elevation.
    Used only when element has no level parameter.
    """
    try:
        elevation = None
        
        # Try to get base level + offset (for columns, framing, etc.)
        try:
            base_level_param = element.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM)
            if not base_level_param or not base_level_param.HasValue:
                base_level_param = element.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM)
            
            if base_level_param and base_level_param.HasValue:
                level_id = base_level_param.AsElementId()
                level = doc.GetElement(level_id)
                if level:
                    level_elevation = level.Elevation
                    
                    # Get base offset if it exists
                    base_offset = 0.0
                    offset_param = element.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM)
                    if offset_param and offset_param.HasValue:
                        base_offset = offset_param.AsDouble()
                    
                    elevation = level_elevation + base_offset
        except:
            pass
        
        # Try location point
        if elevation is None:
            if hasattr(element, 'Location') and element.Location:
                try:
                    location = element.Location
                    if hasattr(location, 'Point'):
                        elevation = location.Point.Z
                except:
                    pass
        
        # Try bounding box (fallback)
        if elevation is None:
            try:
                bbox = element.get_BoundingBox(None)
                if bbox:
                    elevation = bbox.Min.Z
            except:
                pass
        
        # If we found an elevation, determine level designation
        if elevation is not None:
            # Convert from feet to meters
            elevation_meters = elevation * 0.3048
            
            # Add base point elevation to get absolute elevation
            absolute_elevation = base_elevation + elevation_meters
            
            # Check elevation ranges
            if 76.42 <= absolute_elevation < 81.4:
                return "Level -1"
            elif 81.4 <= absolute_elevation < 86.2:
                return "Level 0"
            elif 86.2 <= absolute_elevation < 91.0:
                return "Level 1"
            elif 91.0 <= absolute_elevation < 95.8:
                return "Level 2"
            elif 95.8 <= absolute_elevation < 99.35:
                return "Level 3"
            elif absolute_elevation >= 99.35:
                return "Roof above level 3"
            else:
                return None
        
    except:
        pass
    
    return None

def get_element_level_designation(element, base_elevation):
    """
    Get level designation for element.
    METHOD 2: First tries to get the level name from element's level parameter.
    METHOD 1 (Fallback): If no level parameter, calculates by elevation.
    """
    # METHOD 2: Try to get level name from level parameter
    level_name = get_element_level_name(element)
    if level_name:
        return level_name
    
    # METHOD 1 (Fallback): Calculate by elevation
    return get_element_level_designation_by_elevation(element, base_elevation)

# Get project base point elevation once at the start
project_base_elevation = get_project_base_point_elevation()

# Get current view
active_view = doc.ActiveView

# Get current selection or all visible elements in view
selection = uidoc.Selection.GetElementIds()

if selection:
    elements = [doc.GetElement(id) for id in selection]
    process_selection = True
else:
    # Ask user if they want to process all visible elements in the view
    result = forms.alert(
        "No elements selected. Do you want to process ALL visible elements in the current view?",
        title="No Selection",
        ok=False,
        yes=True,
        no=True
    )
    
    if result:
        # Get all elements visible in the active view
        elements = FilteredElementCollector(doc, active_view.Id).WhereElementIsNotElementType().ToElements()
        process_selection = False
    else:
        forms.alert("Operation cancelled.", title="Cancelled")
        import sys
        sys.exit()

# Process elements
success_count = 0
fail_count = 0
no_param_count = 0

t = Transaction(doc, "Transfer to BES Parameters")
t.Start()

try:
    for element in elements:
        element_processed = False
        
        # Transfer Volume (instance parameter or calculated from geometry, including nested families)
        volume = get_element_volume(element)
        if volume is not None:
            if set_parameter_value(element, "BES_DATA EXTRACTION_Volume", volume):
                element_processed = True
        
        # Transfer Weight (instance parameter, including nested families)
        weight = get_element_weight(element)
        if weight is not None:
            if set_parameter_value(element, "BES_DATA EXTRACTION_Weight", weight):
                element_processed = True
        
        # Transfer Area (instance parameter) - ONLY if element has visible Area parameter
        # Pass the raw internal value (square feet) - Revit will handle the display conversion
        if has_visible_area_parameter(element):
            area = get_parameter_value(element, BuiltInParameter.HOST_AREA_COMPUTED)
            if area is not None:
                if set_parameter_value(element, "BES_DATA EXTRACTION_Area", area):
                    element_processed = True
        
        # Transfer Assembly Code (type parameter)
        assembly_code = get_type_parameter_value(element, BuiltInParameter.UNIFORMAT_CODE)
        if assembly_code:
            if set_parameter_value(element, "BES_DATA EXTRACTION_Assembly Code", assembly_code):
                element_processed = True
        
        # Transfer Assembly Description (type parameter)
        assembly_desc = get_type_parameter_value(element, BuiltInParameter.UNIFORMAT_DESCRIPTION)
        if assembly_desc:
            if set_parameter_value(element, "BES_DATA EXTRACTION_Assembly Description", assembly_desc):
                element_processed = True
        
        # Transfer Family Name (or Category Name for system families)
        family_name = get_family_name(element)
        if family_name:
            if set_parameter_value(element, "BES_DATA EXTRACTION_Family Name", family_name):
                element_processed = True
        
        # Transfer Category
        category_name = get_category_name(element)
        if category_name:
            if set_parameter_value(element, "BES_DATA EXTRACTION_Category", category_name):
                element_processed = True
        
        # Transfer Material (including nested families)
        material_name = get_material_name(element)
        if material_name:
            if set_parameter_value(element, "BES_DATA EXTRACTION_Material", material_name):
                element_processed = True
        
        # Transfer Level designation (METHOD 2 with METHOD 1 fallback)
        level_designation = get_element_level_designation(element, project_base_elevation)
        if level_designation:
            if set_parameter_value(element, "BES_DATA EXTRACTION_Level", level_designation):
                element_processed = True
        
        # Transfer Type Mark (look up by name from type)
        type_mark = get_type_parameter_by_name(element, "Type Mark")
        if type_mark:
            if set_parameter_value(element, "BES_DATA EXTRACTION.Type Mark", type_mark):
                element_processed = True
        
        if element_processed:
            success_count += 1
        else:
            # Check if element has any of the BES parameters
            has_bes_param = False
            for param_name in ["BES_DATA EXTRACTION_Volume",
                             "BES_DATA EXTRACTION_Weight",
                             "BES_DATA EXTRACTION_Area",
                             "BES_DATA EXTRACTION_Assembly Code",
                             "BES_DATA EXTRACTION_Assembly Description",
                             "BES_DATA EXTRACTION_Family Name",
                             "BES_DATA EXTRACTION_Category",
                             "BES_DATA EXTRACTION_Material",
                             "BES_DATA EXTRACTION_Level",
                             "BES_DATA EXTRACTION.Type Mark"]:
                if element.LookupParameter(param_name):
                    has_bes_param = True
                    break
            
            if has_bes_param:
                fail_count += 1
            else:
                no_param_count += 1
    
    t.Commit()
    
    # Report results
    message = "Transfer completed!\n\n"
    message += "Project Base Elevation: {} m\n".format(project_base_elevation)
    message += "Elements updated: {}\n".format(success_count)
    
    if fail_count > 0:
        message += "Elements with BES parameters but no source data: {}\n".format(fail_count)
    
    if no_param_count > 0:
        message += "Elements without BES parameters: {}".format(no_param_count)
    
    forms.alert(message, title="Success")

except Exception as e:
    t.RollBack()
    forms.alert("Error: {}".format(str(e)), title="Error")