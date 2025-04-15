# -*- coding: utf-8 -*-
"""
Openings Caracterisation
"""
#__title__ = 'Openings Caract#erisation'
__author__ = 'ADA55'


import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from pyrevit import script

# Get the current document
doc = __revit__.ActiveUIDocument.Document
output = script.get_output()
output.print_md("# Face Opening Elements - Survey Elevations")

# Get project location and active project position
project_location = doc.ActiveProjectLocation
project_position = project_location.GetProjectPosition(XYZ.Zero)
site_elevation = project_position.Elevation
angle = project_position.Angle
east_west = project_position.EastWest
north_south = project_position.NorthSouth

# Conversion factors
foot_to_meter = 0.3048
meter_to_mm = 1000

# Get all generic model family instances
generic_models = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_GenericModel)\
    .WhereElementIsNotElementType()\
    .ToElements()

# Filter for Face Opening models
face_opening_models = []
for element in generic_models:
    try:
        if isinstance(element, FamilyInstance):
            family_name = element.Symbol.Family.Name
            if "Face Opening" in family_name:
                face_opening_models.append(element)
    except:
        pass

# Start a transaction
t = Transaction(doc, "Update Element IDs and Elevations")
t.Start()

updated_id = 0
updated_elev = 0
not_found_id = 0
not_found_elev = 0
failed_id = 0
failed_elev = 0

# Update each element and get elevation
for element in face_opening_models:
    try:
        # Get the element ID
        element_id = element.Id.IntegerValue
        
        # Update element ID parameter
        id_param = element.LookupParameter("VAL_Element_ID_Instance")
        if id_param:
            try:
                id_param.Set(element_id)
                updated_id += 1
            except:
                failed_id += 1
        else:
            not_found_id += 1
        
        # Get elevation parameter
        elev_param = element.LookupParameter("VAL_Arrase AXE_Instance")
        
        # Calculate elevation
        location = element.Location
        if isinstance(location, LocationPoint):
            point = location.Point
            
            # Get element Z coordinate in project coordinates
            element_z = point.Z
            
            # Convert to survey coordinates by adding the site elevation offset
            survey_z = element_z + site_elevation
            
            # Convert to millimeters
            elevation_in_mm = survey_z * foot_to_meter * meter_to_mm
            
            # Convert mm value to feet for Revit's internal units
            mm_value_in_feet = elevation_in_mm / (meter_to_mm * foot_to_meter)
            
            # Update elevation parameter
            if elev_param:
                try:
                    # For dimension parameters, set the value in internal units (feet)
                    elev_param.Set(mm_value_in_feet)
                    updated_elev += 1
                except:
                    failed_elev += 1
            else:
                not_found_elev += 1
    
    except:
        failed_id += 1
        failed_elev += 1

t.Commit()

# Print survey coordinate system information
output.print_md("## Survey Coordinate System Information")
output.print_md("- Site Elevation: {:.3f} m ({:,.0f} mm)".format(
    site_elevation * foot_to_meter, 
    site_elevation * foot_to_meter * meter_to_mm))
output.print_md("- East-West: {:.3f} m".format(east_west * foot_to_meter))
output.print_md("- North-South: {:.3f} m".format(north_south * foot_to_meter))
output.print_md("- Angle to True North: {:.6f} radians ({:.2f} degrees)".format(
    angle, angle * 180 / 3.14159))

# Print results
output.print_md("## Update Summary")
output.print_md("Found {} Face Opening elements".format(len(face_opening_models)))

output.print_md("### Updated Parameters")
output.print_md("- Element ID (VAL_Element_ID_Instance): {} of {} updated".format(
    updated_id, len(face_opening_models)))
output.print_md("- Elevation (VAL_Arrase AXE_Instance): {} of {} updated".format(
    updated_elev, len(face_opening_models)))

if failed_id > 0 or failed_elev > 0 or not_found_id > 0 or not_found_elev > 0:
    output.print_md("### Issues")
    if not_found_id > 0:
        output.print_md("- Element ID parameter not found in {} elements".format(not_found_id))
    if not_found_elev > 0:
        output.print_md("- Elevation parameter not found in {} elements".format(not_found_elev))
    if failed_id > 0:
        output.print_md("- Failed to update Element ID for {} elements".format(failed_id))
    if failed_elev > 0:
        output.print_md("- Failed to update Elevation for {} elements".format(failed_elev))

output.print_md("## Notes")
output.print_md("- All elevation values are stored in parameters as millimeter values")
output.print_md("- The actual displayed value will depend on the project's unit settings")
output.print_md("- The script converts survey coordinates (Z) to millimeters above sea level")
