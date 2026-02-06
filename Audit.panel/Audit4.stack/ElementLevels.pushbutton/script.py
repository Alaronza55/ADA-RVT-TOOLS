"""
Retrieve Model Elements with ID and Level
"""
__title__ = "Export Elements\nwith Level"
__author__ = "Your Name"

# Import required modules
from pyrevit import revit, DB, forms
import os

# Get current document
doc = revit.doc
folder_name = doc.Title

def clean_text(text):
    """Clean text to remove problematic characters"""
    try:
        # Convert to string and replace problematic characters
        text = str(text)
        # Replace common problematic characters
        text = text.replace('\u25B2', 'triangle')  # Triangle symbol
        text = text.replace('\u00B0', 'deg')       # Degree symbol
        text = text.replace('\u00B2', '2')         # Superscript 2
        text = text.replace('\u00B3', '3')         # Superscript 3
        text = text.replace('\u2013', '-')         # En dash
        text = text.replace('\u2014', '-')         # Em dash
        text = text.replace('\u2019', "'")         # Right single quotation mark
        text = text.replace('\u201C', '"')         # Left double quotation mark
        text = text.replace('\u201D', '"')         # Right double quotation mark

        # Remove any remaining non-ASCII characters
        text = ''.join(char if ord(char) < 128 else '?' for char in text)
        return text
    except:
        return "Unknown"

def get_parent_element(element):
    """Get the parent element if this is a nested element"""
    try:
        # Check if element has SuperComponent property (for nested elements)
        if hasattr(element, 'SuperComponent') and element.SuperComponent:
            return element.SuperComponent
        
        # For FamilyInstances, check if it's nested
        if isinstance(element, DB.FamilyInstance):
            # Check if it has a host family
            if hasattr(element, 'Host') and element.Host:
                # If the host is also a FamilyInstance, it might be a nested situation
                if isinstance(element.Host, DB.FamilyInstance):
                    return element.Host
        
        # Try to get parent through GetSubComponentIds (reverse lookup)
        try:
            # Get all FamilyInstances and check which one contains this as a subcomponent
            all_family_instances = DB.FilteredElementCollector(doc).OfClass(DB.FamilyInstance).ToElements()
            for family_instance in all_family_instances:
                try:
                    sub_component_ids = family_instance.GetSubComponentIds()
                    if element.Id in sub_component_ids:
                        return family_instance
                except:
                    continue
        except:
            pass
            
        return None
    except:
        return None

def get_schedule_level_from_element(element, element_description=""):
    """Generic function to get schedule level from an element"""
    # Method 1: Try to find Schedule Level parameter by name
    try:
        all_params = element.Parameters
        for param in all_params:
            param_name = param.Definition.Name.lower()
            if 'schedule level' in param_name or 'schedule_level' in param_name:
                if param.HasValue:
                    if param.StorageType == DB.StorageType.ElementId:
                        element_id = param.AsElementId()
                        if element_id != DB.ElementId.InvalidElementId:
                            level_element = doc.GetElement(element_id)
                            if level_element and hasattr(level_element, 'Name'):
                                return clean_text(level_element.Name) + " (schedule level{})".format(element_description)
                    elif param.StorageType == DB.StorageType.String:
                        level_name = param.AsString()
                        if level_name:
                            return clean_text(level_name) + " (schedule level{})".format(element_description)
                    elif param.StorageType == DB.StorageType.Double or param.StorageType == DB.StorageType.Integer:
                        level_value = param.AsValueString()
                        if level_value:
                            return clean_text(level_value) + " (schedule level{})".format(element_description)
    except:
        pass
    
    # Method 2: Try built-in schedule level parameter
    try:
        schedule_level_param = element.get_Parameter(DB.BuiltInParameter.SCHEDULE_LEVEL_PARAM)
        if schedule_level_param and schedule_level_param.HasValue:
            if schedule_level_param.StorageType == DB.StorageType.ElementId:
                element_id = schedule_level_param.AsElementId()
                if element_id != DB.ElementId.InvalidElementId:
                    level_element = doc.GetElement(element_id)
                    if level_element:
                        return clean_text(level_element.Name) + " (schedule level{})".format(element_description)
            elif schedule_level_param.StorageType == DB.StorageType.String:
                level_name = schedule_level_param.AsString()
                if level_name:
                    return clean_text(level_name) + " (schedule level{})".format(element_description)
    except:
        pass
    
    # Method 3: Check shared parameters for schedule level
    try:
        for param in element.Parameters:
            if param.IsShared:
                param_name = param.Definition.Name.lower()
                if 'schedule level' in param_name or 'schedul' in param_name and 'level' in param_name:
                    if param.HasValue:
                        if param.StorageType == DB.StorageType.ElementId:
                            element_id = param.AsElementId()
                            if element_id != DB.ElementId.InvalidElementId:
                                level_element = doc.GetElement(element_id)
                                if level_element and hasattr(level_element, 'Name'):
                                    return clean_text(level_element.Name) + " (schedule level{})".format(element_description)
                        elif param.StorageType == DB.StorageType.String:
                            level_name = param.AsString()
                            if level_name:
                                return clean_text(level_name) + " (schedule level{})".format(element_description)
    except:
        pass
    
    return None

def get_element_level(element):
    """Get the level associated with an element - with special handling for structural framing, planting, generic models, plumbing fixtures, mechanical equipment, lighting fixtures, security devices, electrical fixtures, communication devices, furniture, and nested elements"""
    try:
        # First check if this is a nested element and handle accordingly
        parent_element = get_parent_element(element)
        
        # Special handling for Generic Models - check Schedule Level first
        if element.Category and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_GenericModel):
            
            # Check schedule level on the element itself
            schedule_level = get_schedule_level_from_element(element)
            if schedule_level:
                return schedule_level
            
            # If nested, check parent element for schedule level
            if parent_element:
                parent_schedule_level = get_schedule_level_from_element(parent_element, " from parent")
                if parent_schedule_level:
                    return parent_schedule_level
            
            # Standard level parameters for generic models
            generic_level_params = [
                DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
                DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                DB.BuiltInParameter.LEVEL_PARAM,
                DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM
            ]
            
            for param_type in generic_level_params:
                try:
                    level_param = element.get_Parameter(param_type)
                    if level_param and level_param.HasValue:
                        element_id = level_param.AsElementId()
                        if element_id != DB.ElementId.InvalidElementId:
                            level_element = doc.GetElement(element_id)
                            if level_element:
                                return clean_text(level_element.Name)
                except:
                    continue
            
            # If nested and no level found on element, try parent's standard levels
            if parent_element:
                for param_type in generic_level_params:
                    try:
                        level_param = parent_element.get_Parameter(param_type)
                        if level_param and level_param.HasValue:
                            element_id = level_param.AsElementId()
                            if element_id != DB.ElementId.InvalidElementId:
                                level_element = doc.GetElement(element_id)
                                if level_element:
                                    return clean_text(level_element.Name) + " (from parent)"
                    except:
                        continue

        # Special handling for Plumbing Fixtures - check Schedule Level first
        elif element.Category and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_PlumbingFixtures):
            
            # Check schedule level on the element itself
            schedule_level = get_schedule_level_from_element(element)
            if schedule_level:
                return schedule_level
            
            # If nested, check parent element for schedule level
            if parent_element:
                parent_schedule_level = get_schedule_level_from_element(parent_element, " from parent")
                if parent_schedule_level:
                    return parent_schedule_level
            
            # Standard level parameters for plumbing fixtures
            plumbing_level_params = [
                DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
                DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                DB.BuiltInParameter.LEVEL_PARAM,
                DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM
            ]
            
            for param_type in plumbing_level_params:
                try:
                    level_param = element.get_Parameter(param_type)
                    if level_param and level_param.HasValue:
                        element_id = level_param.AsElementId()
                        if element_id != DB.ElementId.InvalidElementId:
                            level_element = doc.GetElement(element_id)
                            if level_element:
                                return clean_text(level_element.Name)
                except:
                    continue
            
            # If nested and no level found on element, try parent's standard levels
            if parent_element:
                for param_type in plumbing_level_params:
                    try:
                        level_param = parent_element.get_Parameter(param_type)
                        if level_param and level_param.HasValue:
                            element_id = level_param.AsElementId()
                            if element_id != DB.ElementId.InvalidElementId:
                                level_element = doc.GetElement(element_id)
                                if level_element:
                                    return clean_text(level_element.Name) + " (from parent)"
                    except:
                        continue

        # Special handling for Mechanical Equipment - check Schedule Level first
        elif element.Category and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_MechanicalEquipment):
            
            # Check schedule level on the element itself
            schedule_level = get_schedule_level_from_element(element)
            if schedule_level:
                return schedule_level
            
            # If nested, check parent element for schedule level
            if parent_element:
                parent_schedule_level = get_schedule_level_from_element(parent_element, " from parent")
                if parent_schedule_level:
                    return parent_schedule_level
            
            # Standard level parameters for mechanical equipment
            mech_level_params = [
                DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
                DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                DB.BuiltInParameter.LEVEL_PARAM,
                DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM
            ]
            
            for param_type in mech_level_params:
                try:
                    level_param = element.get_Parameter(param_type)
                    if level_param and level_param.HasValue:
                        element_id = level_param.AsElementId()
                        if element_id != DB.ElementId.InvalidElementId:
                            level_element = doc.GetElement(element_id)
                            if level_element:
                                return clean_text(level_element.Name)
                except:
                    continue
            
            # If nested and no level found on element, try parent's standard levels
            if parent_element:
                for param_type in mech_level_params:
                    try:
                        level_param = parent_element.get_Parameter(param_type)
                        if level_param and level_param.HasValue:
                            element_id = level_param.AsElementId()
                            if element_id != DB.ElementId.InvalidElementId:
                                level_element = doc.GetElement(element_id)
                                if level_element:
                                    return clean_text(level_element.Name) + " (from parent)"
                    except:
                        continue

        # Special handling for Lighting Fixtures - check Schedule Level first
        elif element.Category and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_LightingFixtures):
            
            # Check schedule level on the element itself
            schedule_level = get_schedule_level_from_element(element)
            if schedule_level:
                return schedule_level
            
            # If nested, check parent element for schedule level
            if parent_element:
                parent_schedule_level = get_schedule_level_from_element(parent_element, " from parent")
                if parent_schedule_level:
                    return parent_schedule_level
            
            # Standard level parameters for lighting fixtures
            lighting_level_params = [
                DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
                DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                DB.BuiltInParameter.LEVEL_PARAM,
                DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM
            ]
            
            for param_type in lighting_level_params:
                try:
                    level_param = element.get_Parameter(param_type)
                    if level_param and level_param.HasValue:
                        element_id = level_param.AsElementId()
                        if element_id != DB.ElementId.InvalidElementId:
                            level_element = doc.GetElement(element_id)
                            if level_element:
                                return clean_text(level_element.Name)
                except:
                    continue
            
            # If nested and no level found on element, try parent's standard levels
            if parent_element:
                for param_type in lighting_level_params:
                    try:
                        level_param = parent_element.get_Parameter(param_type)
                        if level_param and level_param.HasValue:
                            element_id = level_param.AsElementId()
                            if element_id != DB.ElementId.InvalidElementId:
                                level_element = doc.GetElement(element_id)
                                if level_element:
                                    return clean_text(level_element.Name) + " (from parent)"
                    except:
                        continue

        # Special handling for Security Devices - check Schedule Level first
        elif element.Category and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_SecurityDevices):
            
            # Check schedule level on the element itself
            schedule_level = get_schedule_level_from_element(element)
            if schedule_level:
                return schedule_level
            
            # If nested, check parent element for schedule level
            if parent_element:
                parent_schedule_level = get_schedule_level_from_element(parent_element, " from parent")
                if parent_schedule_level:
                    return parent_schedule_level
            
            # Standard level parameters for security devices
            security_level_params = [
                DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
                DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                DB.BuiltInParameter.LEVEL_PARAM,
                DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM
            ]
            
            for param_type in security_level_params:
                try:
                    level_param = element.get_Parameter(param_type)
                    if level_param and level_param.HasValue:
                        element_id = level_param.AsElementId()
                        if element_id != DB.ElementId.InvalidElementId:
                            level_element = doc.GetElement(element_id)
                            if level_element:
                                return clean_text(level_element.Name)
                except:
                    continue
            
            # If nested and no level found on element, try parent's standard levels
            if parent_element:
                for param_type in security_level_params:
                    try:
                        level_param = parent_element.get_Parameter(param_type)
                        if level_param and level_param.HasValue:
                            element_id = level_param.AsElementId()
                            if element_id != DB.ElementId.InvalidElementId:
                                level_element = doc.GetElement(element_id)
                                if level_element:
                                    return clean_text(level_element.Name) + " (from parent)"
                    except:
                        continue

        # Special handling for Electrical Fixtures - check Schedule Level first
        elif element.Category and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_ElectricalFixtures):
            
            # Check schedule level on the element itself
            schedule_level = get_schedule_level_from_element(element)
            if schedule_level:
                return schedule_level
            
            # If nested, check parent element for schedule level
            if parent_element:
                parent_schedule_level = get_schedule_level_from_element(parent_element, " from parent")
                if parent_schedule_level:
                    return parent_schedule_level
            
            # Standard level parameters for electrical fixtures
            electrical_level_params = [
                DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
                DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                DB.BuiltInParameter.LEVEL_PARAM,
                DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM
            ]
            
            for param_type in electrical_level_params:
                try:
                    level_param = element.get_Parameter(param_type)
                    if level_param and level_param.HasValue:
                        element_id = level_param.AsElementId()
                        if element_id != DB.ElementId.InvalidElementId:
                            level_element = doc.GetElement(element_id)
                            if level_element:
                                return clean_text(level_element.Name)
                except:
                    continue
            
            # If nested and no level found on element, try parent's standard levels
            if parent_element:
                for param_type in electrical_level_params:
                    try:
                        level_param = parent_element.get_Parameter(param_type)
                        if level_param and level_param.HasValue:
                            element_id = level_param.AsElementId()
                            if element_id != DB.ElementId.InvalidElementId:
                                level_element = doc.GetElement(element_id)
                                if level_element:
                                    return clean_text(level_element.Name) + " (from parent)"
                    except:
                        continue

        # Special handling for Communication Devices - check Schedule Level first
        elif element.Category and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_CommunicationDevices):
            
            # Check schedule level on the element itself
            schedule_level = get_schedule_level_from_element(element)
            if schedule_level:
                return schedule_level
            
            # If nested, check parent element for schedule level
            if parent_element:
                parent_schedule_level = get_schedule_level_from_element(parent_element, " from parent")
                if parent_schedule_level:
                    return parent_schedule_level
            
            # Standard level parameters for communication devices
            communication_level_params = [
                DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
                DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                DB.BuiltInParameter.LEVEL_PARAM,
                DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM
            ]
            
            for param_type in communication_level_params:
                try:
                    level_param = element.get_Parameter(param_type)
                    if level_param and level_param.HasValue:
                        element_id = level_param.AsElementId()
                        if element_id != DB.ElementId.InvalidElementId:
                            level_element = doc.GetElement(element_id)
                            if level_element:
                                return clean_text(level_element.Name)
                except:
                    continue
            
            # If nested and no level found on element, try parent's standard levels
            if parent_element:
                for param_type in communication_level_params:
                    try:
                        level_param = parent_element.get_Parameter(param_type)
                        if level_param and level_param.HasValue:
                            element_id = level_param.AsElementId()
                            if element_id != DB.ElementId.InvalidElementId:
                                level_element = doc.GetElement(element_id)
                                if level_element:
                                    return clean_text(level_element.Name) + " (from parent)"
                    except:
                        continue

        # Special handling for Furniture - check Schedule Level first
        elif element.Category and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_Furniture):
            
            # Check schedule level on the element itself
            schedule_level = get_schedule_level_from_element(element)
            if schedule_level:
                return schedule_level
            
            # If nested, check parent element for schedule level
            if parent_element:
                parent_schedule_level = get_schedule_level_from_element(parent_element, " from parent")
                if parent_schedule_level:
                    return parent_schedule_level
            
            # Standard level parameters for furniture
            furniture_level_params = [
                DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
                DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                DB.BuiltInParameter.LEVEL_PARAM,
                DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM
            ]
            
            for param_type in furniture_level_params:
                try:
                    level_param = element.get_Parameter(param_type)
                    if level_param and level_param.HasValue:
                        element_id = level_param.AsElementId()
                        if element_id != DB.ElementId.InvalidElementId:
                            level_element = doc.GetElement(element_id)
                            if level_element:
                                return clean_text(level_element.Name)
                except:
                    continue
            
            # If nested and no level found on element, try parent's standard levels
            if parent_element:
                for param_type in furniture_level_params:
                    try:
                        level_param = parent_element.get_Parameter(param_type)
                        if level_param and level_param.HasValue:
                            element_id = level_param.AsElementId()
                            if element_id != DB.ElementId.InvalidElementId:
                                level_element = doc.GetElement(element_id)
                                if level_element:
                                    return clean_text(level_element.Name) + " (from parent)"
                    except:
                        continue

        # Special handling for Planting category - check Schedule Level first, then Host
        elif element.Category and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_Planting):
            
            # Check schedule level on the element itself FIRST
            schedule_level = get_schedule_level_from_element(element)
            if schedule_level:
                return schedule_level
            
            # If nested, check parent element for schedule level
            if parent_element:
                parent_schedule_level = get_schedule_level_from_element(parent_element, " from parent")
                if parent_schedule_level:
                    return parent_schedule_level
            
            # Method 2: Try to get level from Host element 
            try:
                if hasattr(element, 'Host') and element.Host:
                    host = element.Host
                    
                    # Try to get level from host's level parameters
                    host_level_params = [
                        DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
                        DB.BuiltInParameter.LEVEL_PARAM,
                        DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM,
                        DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM
                    ]
                    
                    for param_type in host_level_params:
                        try:
                            host_level_param = host.get_Parameter(param_type)
                            if host_level_param and host_level_param.HasValue:
                                element_id = host_level_param.AsElementId()
                                if element_id != DB.ElementId.InvalidElementId:
                                    level_element = doc.GetElement(element_id)
                                    if level_element:
                                        return clean_text(level_element.Name) + " (from host)"
                        except:
                            continue
                    
                    # Try host's direct Level property
                    if hasattr(host, 'Level') and host.Level:
                        return clean_text(host.Level.Name) + " (from host)"
                    
                    # Try host's LevelId property
                    if hasattr(host, 'LevelId') and host.LevelId != DB.ElementId.InvalidElementId:
                        level = doc.GetElement(host.LevelId)
                        if level:
                            return clean_text(level.Name) + " (from host)"
            except:
                pass
            
            # Method 3: Try standard level parameters for planting
            planting_level_params = [
                DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
                DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                DB.BuiltInParameter.LEVEL_PARAM,
                DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM
            ]
            
            for param_type in planting_level_params:
                try:
                    level_param = element.get_Parameter(param_type)
                    if level_param and level_param.HasValue:
                        element_id = level_param.AsElementId()
                        if element_id != DB.ElementId.InvalidElementId:
                            level_element = doc.GetElement(element_id)
                            if level_element:
                                return clean_text(level_element.Name)
                except:
                    continue
            
            # Method 4: Location-based approach for planting
            try:
                location = element.Location
                if hasattr(location, 'Point'):
                    point = location.Point
                    # Find the level that this planting element is closest to
                    all_levels = DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements()
                    if all_levels:
                        # Sort levels by elevation
                        sorted_levels = sorted(all_levels, key=lambda x: x.Elevation)
                        
                        # Find the appropriate level based on Z coordinate
                        element_z = point.Z
                        for i, level in enumerate(sorted_levels):
                            # If element is above this level and below next level (or if it's the top level)
                            if i == len(sorted_levels) - 1 or element_z < sorted_levels[i + 1].Elevation:
                                if element_z >= level.Elevation - 1.0:  # 1 foot tolerance below level
                                    return clean_text(level.Name) + " (calculated)"
            except:
                pass

        # Special handling for structural framing - try multiple approaches
        elif element.Category and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_StructuralFraming):
            
            # Method 1: Try Reference Level built-in parameters
            reference_level_params = [
                DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
                DB.BuiltInParameter.LEVEL_PARAM,
                DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM
            ]
            
            for param_type in reference_level_params:
                try:
                    level_param = element.get_Parameter(param_type)
                    if level_param and level_param.HasValue:
                        element_id = level_param.AsElementId()
                        if element_id != DB.ElementId.InvalidElementId:
                            level_element = doc.GetElement(element_id)
                            if level_element:
                                return clean_text(level_element.Name)
                except:
                    continue
            
            # If nested, try parent element
            if parent_element:
                for param_type in reference_level_params:
                    try:
                        level_param = parent_element.get_Parameter(param_type)
                        if level_param and level_param.HasValue:
                            element_id = level_param.AsElementId()
                            if element_id != DB.ElementId.InvalidElementId:
                                level_element = doc.GetElement(element_id)
                                if level_element:
                                    return clean_text(level_element.Name) + " (from parent)"
                    except:
                        continue
            
            # Method 2: Try to get reference level by parameter name
            try:
                all_params = element.Parameters
                for param in all_params:
                    param_name = param.Definition.Name.lower()
                    if 'reference level' in param_name or 'ref level' in param_name:
                        if param.HasValue and param.StorageType == DB.StorageType.ElementId:
                            element_id = param.AsElementId()
                            if element_id != DB.ElementId.InvalidElementId:
                                level_element = doc.GetElement(element_id)
                                if level_element and level_element.GetType().Name == "Level":
                                    return clean_text(level_element.Name)
            except:
                pass

        # Standard handling for all other elements (and fallback for special categories)
        common_level_params = [
            DB.BuiltInParameter.FAMILY_LEVEL_PARAM,
            DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM,
            DB.BuiltInParameter.LEVEL_PARAM,
            DB.BuiltInParameter.WALL_BASE_CONSTRAINT,
            DB.BuiltInParameter.STAIRS_BASE_LEVEL_PARAM,
            DB.BuiltInParameter.ROOF_BASE_LEVEL_PARAM,
            DB.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM
        ]

        for param_type in common_level_params:
            try:
                level_param = element.get_Parameter(param_type)
                if level_param and level_param.HasValue:
                    element_id = level_param.AsElementId()
                    if element_id != DB.ElementId.InvalidElementId:
                        level_element = doc.GetElement(element_id)
                        if level_element:
                            return clean_text(level_element.Name)
            except:
                continue

        # If no level found on element and it's nested, try parent element
        if parent_element:
            for param_type in common_level_params:
                try:
                    level_param = parent_element.get_Parameter(param_type)
                    if level_param and level_param.HasValue:
                        element_id = level_param.AsElementId()
                        if element_id != DB.ElementId.InvalidElementId:
                            level_element = doc.GetElement(element_id)
                            if level_element:
                                return clean_text(level_element.Name) + " (from parent)"
                except:
                    continue

        # Try direct Level property
        if hasattr(element, 'Level') and element.Level:
            return clean_text(element.Level.Name)
        
        # If nested, try parent's direct Level property
        if parent_element and hasattr(parent_element, 'Level') and parent_element.Level:
            return clean_text(parent_element.Level.Name) + " (from parent)"

        # Try LevelId property
        if hasattr(element, 'LevelId') and element.LevelId != DB.ElementId.InvalidElementId:
            level = doc.GetElement(element.LevelId)
            if level:
                return clean_text(level.Name)
        
        # If nested, try parent's LevelId property
        if parent_element and hasattr(parent_element, 'LevelId') and parent_element.LevelId != DB.ElementId.InvalidElementId:
            level = doc.GetElement(parent_element.LevelId)
            if level:
                return clean_text(level.Name) + " (from parent)"

        # For hosted elements, try to get level from host (general fallback)
        try:
            if hasattr(element, 'Host') and element.Host:
                host_level_param = element.Host.get_Parameter(DB.BuiltInParameter.FAMILY_LEVEL_PARAM)
                if host_level_param and host_level_param.HasValue:
                    element_id = host_level_param.AsElementId()
                    if element_id != DB.ElementId.InvalidElementId:
                        level_element = doc.GetElement(element_id)
                        if level_element:
                            return clean_text(level_element.Name) + " (from host)"
        except:
            pass

        return "No Level"

    except Exception as e:
        return "Error: {}".format(str(e))

def main():
    """Main function to export elements with levels"""

    # Define categories to collect
    categories = [
        DB.BuiltInCategory.OST_Walls,
        DB.BuiltInCategory.OST_Doors,
        DB.BuiltInCategory.OST_Windows,
        DB.BuiltInCategory.OST_Floors,
        DB.BuiltInCategory.OST_Ceilings,
        DB.BuiltInCategory.OST_Roofs,
        DB.BuiltInCategory.OST_Stairs,
        DB.BuiltInCategory.OST_StairsRailing,
        DB.BuiltInCategory.OST_Ramps,
        DB.BuiltInCategory.OST_CurtainWallMullions,
        DB.BuiltInCategory.OST_CurtainWallPanels,
        DB.BuiltInCategory.OST_Columns,
        DB.BuiltInCategory.OST_Furniture,
        DB.BuiltInCategory.OST_FurnitureSystems,
        DB.BuiltInCategory.OST_Casework,
        DB.BuiltInCategory.OST_StructuralColumns,
        DB.BuiltInCategory.OST_StructuralFraming,
        DB.BuiltInCategory.OST_StructuralFoundation,
        DB.BuiltInCategory.OST_Rebar,
        DB.BuiltInCategory.OST_MechanicalEquipment,
        DB.BuiltInCategory.OST_ElectricalEquipment,
        DB.BuiltInCategory.OST_ElectricalFixtures,
        DB.BuiltInCategory.OST_LightingFixtures,
        DB.BuiltInCategory.OST_PlumbingFixtures,
        DB.BuiltInCategory.OST_Sprinklers,
        DB.BuiltInCategory.OST_CommunicationDevices,
        DB.BuiltInCategory.OST_SecurityDevices,
        DB.BuiltInCategory.OST_NurseCallDevices,
        DB.BuiltInCategory.OST_Planting,
        DB.BuiltInCategory.OST_Site,
        DB.BuiltInCategory.OST_Parking,
        DB.BuiltInCategory.OST_Entourage,
        DB.BuiltInCategory.OST_GenericModel,
        DB.BuiltInCategory.OST_Mass,
        DB.BuiltInCategory.OST_SpecialityEquipment
    ]

    # Collect elements from all categories
    all_elements = []
    for category in categories:
        try:
            category_elements = DB.FilteredElementCollector(doc)\
                                 .OfCategory(category)\
                                 .WhereElementIsNotElementType()\
                                 .ToElements()
            all_elements.extend(category_elements)
        except:
            # Some categories might not exist in all documents
            continue

    # Prepare data for export
    data = []
    for element in all_elements:
        try:
            # Get element name
            element_name = "Unnamed"
            try:
                name_param = element.get_Parameter(DB.BuiltInParameter.ELEM_FAMILY_AND_TYPE_PARAM)
                if name_param and name_param.HasValue:
                    element_name = name_param.AsValueString()
            except:
                pass

            # Get element ID
            element_id = str(element.Id.IntegerValue)

            # Get element level
            element_level = get_element_level(element)

            # Get category name
            category_name = "Unknown Category"
            try:
                if element.Category:
                    category_name = element.Category.Name
            except:
                pass

            # Add to data
            data.append([element_name, element_id, element_level, category_name])

        except:
            # Skip problematic elements
            continue

    if not data:
        forms.alert("No elements found to export.")
        return

    # Create filename
    output_folder = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker\{}".format(folder_name)
    csv_filename = "Model_Elements_with_Levels.csv"
    file_path = os.path.join(output_folder, csv_filename)
                
    try:
        # Create directory if it doesn't exist
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            
        # Create CSV file with explicit encoding
        with open(file_path, 'w') as f:
            # Write headers
            f.write("Element Name,Element ID,Level,Category\n")

            # Write data
            for row_data in data:
                # Clean and escape data
                escaped_row = []
                for item in row_data:
                    item_str = clean_text(str(item))
                    if ',' in item_str or '"' in item_str:
                        # Escape quotes and wrap in quotes
                        item_str = item_str.replace('"', '""')
                        escaped_row.append('"{}"'.format(item_str))
                    else:
                        escaped_row.append(item_str)

                f.write("{}\n".format(",".join(escaped_row)))

        # Show success message
        message = "Export completed successfully!\n\n"
        message += "File saved to: {}\n".format(file_path)
        message += "Total elements exported: {}\n\n".format(len(data))
        message += "Categories included: Walls, Doors, Windows, Floors, Ceilings, Roofs, Stairs, Railings, Ramps, "
        message += "Curtain Wall Mullions, Curtain Panels, Columns, Furniture, Furniture Systems, Casework, "
        message += "Structural Columns, Structural Framing, Structural Foundations, Structural Rebar, "
        message += "Mechanical Equipment, Electrical Equipment, Electrical Fixtures, Lighting Fixtures, "
        message += "Plumbing Fixtures, Fire Protection, Communication Devices, Security Devices, "
        message += "Nurse Call Devices, Planting, Site, Parking, Entourage, Generic Models, Mass, Specialty Equipment\n\n"
        message += "Special handling:\n"
        message += "- Structural Framing: Uses reference level when available\n"
        message += "- Planting: Checks Schedule Level parameter first, then host element level, then calculates from position\n"
        message += "- Generic Models: Checks Schedule Level parameter first, then standard level parameters\n"
        message += "- Plumbing Fixtures: Checks Schedule Level parameter first, then standard level parameters\n"
        message += "- Mechanical Equipment: Checks Schedule Level parameter first, then standard level parameters\n"
        message += "- Lighting Fixtures: Checks Schedule Level parameter first, then standard level parameters\n"
        message += "- Security Devices: Checks Schedule Level parameter first, then standard level parameters\n"
        message += "- Electrical Fixtures: Checks Schedule Level parameter first, then standard level parameters\n"
        message += "- Communication Devices: Checks Schedule Level parameter first, then standard level parameters\n"
        message += "- Furniture: Checks Schedule Level parameter first, then standard level parameters\n"
        message += "- Nested Elements: Checks parent element properties when child element has no level\n\n"
        message += "The CSV file can be opened in Excel or any spreadsheet application."

        forms.alert(message)

    except Exception as e:
        forms.alert("Error creating CSV file: {}".format(str(e)))
        print("Detailed error: {}".format(str(e)))

if __name__ == "__main__":
    main()
