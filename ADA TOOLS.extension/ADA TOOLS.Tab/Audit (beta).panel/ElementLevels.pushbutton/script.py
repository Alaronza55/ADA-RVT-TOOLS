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

def get_element_level(element):
    """Get the level associated with an element - with special handling for structural framing, planting, and generic models"""
    try:
        # Special handling for Generic Models - check Schedule Level first
        if element.Category and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_GenericModel):
            
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
                                        return clean_text(level_element.Name) + " (schedule level)"
                            elif param.StorageType == DB.StorageType.String:
                                level_name = param.AsString()
                                if level_name:
                                    return clean_text(level_name) + " (schedule level)"
                            elif param.StorageType == DB.StorageType.Double or param.StorageType == DB.StorageType.Integer:
                                # Sometimes schedule level might be stored as text that needs to be converted
                                level_value = param.AsValueString()
                                if level_value:
                                    return clean_text(level_value) + " (schedule level)"
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
                                return clean_text(level_element.Name) + " (schedule level)"
                    elif schedule_level_param.StorageType == DB.StorageType.String:
                        level_name = schedule_level_param.AsString()
                        if level_name:
                            return clean_text(level_name) + " (schedule level)"
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
                                            return clean_text(level_element.Name) + " (schedule level)"
                                elif param.StorageType == DB.StorageType.String:
                                    level_name = param.AsString()
                                    if level_name:
                                        return clean_text(level_name) + " (schedule level)"
            except:
                pass
            
            # Method 4: Standard level parameters for generic models
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

        # Special handling for Planting category - check Host first
        elif element.Category and element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_Planting):
            
            # Method 1: Try to get level from Host element
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
            
            # Method 2: Try standard level parameters for planting
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
            
            # Method 3: Location-based approach for planting
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
            
            # Method 3: For structural framing, try to get start and end levels
            try:
                if hasattr(element, 'GetAnalyticalModel'):
                    analytical_model = element.GetAnalyticalModel()
                    if analytical_model:
                        # Try to get the curve and calculate level from Z coordinate
                        curve = analytical_model.GetCurve()
                        if curve:
                            start_point = curve.GetEndPoint(0)
                            # Find closest level to this Z coordinate
                            all_levels = DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements()
                            closest_level = None
                            min_distance = float('inf')
                            
                            for level in all_levels:
                                distance = abs(level.Elevation - start_point.Z)
                                if distance < min_distance:
                                    min_distance = distance
                                    closest_level = level
                            
                            if closest_level:
                                return clean_text(closest_level.Name)
            except:
                pass
            
            # Method 4: Try location-based approach for structural framing
            try:
                location = element.Location
                if hasattr(location, 'Curve'):
                    curve = location.Curve
                    if curve:
                        start_point = curve.GetEndPoint(0)
                        # Find the level that this element is closest to
                        all_levels = DB.FilteredElementCollector(doc).OfClass(DB.Level).ToElements()
                        if all_levels:
                            # Sort levels by elevation
                            sorted_levels = sorted(all_levels, key=lambda x: x.Elevation)
                            
                            # Find the appropriate level based on Z coordinate
                            element_z = start_point.Z
                            for i, level in enumerate(sorted_levels):
                                # If element is above this level and below next level (or if it's the top level)
                                if i == len(sorted_levels) - 1 or element_z < sorted_levels[i + 1].Elevation:
                                    if element_z >= level.Elevation - 1.0:  # 1 foot tolerance below level
                                        return clean_text(level.Name)
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

        # Try direct Level property
        if hasattr(element, 'Level') and element.Level:
            return clean_text(element.Level.Name)

        # Try LevelId property
        if hasattr(element, 'LevelId') and element.LevelId != DB.ElementId.InvalidElementId:
            level = doc.GetElement(element.LevelId)
            if level:
                return clean_text(level.Name)

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
    # Create directory if it doesn't exist

    try:
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
        message += "- Planting: Checks host element level first, then calculates from position\n"
        message += "- Generic Models: Checks Schedule Level parameter first, then standard level parameters\n\n"
        message += "The CSV file can be opened in Excel or any spreadsheet application."

        forms.alert(message)

    except Exception as e:
        forms.alert("Error creating CSV file: {}".format(str(e)))
        print("Detailed error: {}".format(str(e)))

if __name__ == "__main__":
    main()
