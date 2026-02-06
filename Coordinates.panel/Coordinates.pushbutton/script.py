# -*- coding: utf-8 -*-
"""Retrieves Project Base Point and Shared Site Coordinates with selectable units.

This script retrieves and displays the coordinates of the Project Base Point
and the Shared Site Point in the current Revit project, allowing the user
to select the desired output units.
"""
__title__ = "Get Project\nCoordinates"
__author__ = "ADA"
__doc__ = "Retrieves and displays the Project Base Point and Shared Site Point coordinates with selectable units"

import clr
from Autodesk.Revit.DB import *
from pyrevit import forms
from pyrevit import script
import math

# Get the current document
doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

# Conversion factors from Revit internal units (feet) to other units
CONVERSION_FACTORS = {
    "Feet": 1.0,
    "Meters": 0.3048,
    "Centimeters": 30.48,
    "Millimeters": 304.8
}

def get_unit_selection():
    """Prompt user to select the output unit."""
    options = ["Meters", "Centimeters", "Millimeters"]
    selected_option = forms.CommandSwitchWindow.show(
        options, 
        message="Select the output unit:"
    )
    
    if not selected_option:
        # User canceled, default to meters
        return "Meters"
    
    return selected_option

def convert_value(value, unit):
    """Convert a value from Revit internal units (feet) to the specified unit."""
    if value is None:
        return None
        
    return value * CONVERSION_FACTORS[unit]

def get_project_base_point(unit):
    """Retrieves the Project Base Point element and its coordinates."""
    # Get all the Project Base Point elements in the project
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType()
    base_points = list(collector)
    
    if not base_points:
        return "No Project Base Point found in the project."
    
    # Usually there's only one Project Base Point
    base_point = base_points[0]
    
    # Get the parameter values
    ns_value = base_point.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble()
    ew_value = base_point.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsDouble()
    elev_value = base_point.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsDouble()
    
    # Convert to selected unit
    ns_converted = convert_value(ns_value, unit)
    ew_converted = convert_value(ew_value, unit)
    elev_converted = convert_value(elev_value, unit)
    
    # Check if we can get the angle as well
    angle_param = base_point.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM)
    angle_value = angle_param.AsDouble() if angle_param else None
    
    # Convert angle to degrees if available
    if angle_value is not None:
        angle_degrees = math.degrees(angle_value)
    else:
        angle_degrees = None
    
    pbp_info = {
        "North/South ({0})".format(unit): "{:.4f}".format(ns_converted),
        "East/West ({0})".format(unit): "{:.4f}".format(ew_converted),
        "Elevation ({0})".format(unit): "{:.4f}".format(elev_converted)
    }
    
    if angle_degrees is not None:
        pbp_info["Angle to True North (degrees)"] = "{:.4f}".format(angle_degrees)
    
    return pbp_info

def get_survey_point(unit):
    """Retrieves the Survey Point element and its coordinates."""
    # Get all the Survey Point elements in the project
    collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_SharedBasePoint).WhereElementIsNotElementType()
    survey_points = list(collector)
    
    if not survey_points:
        return "No Survey Point found in the project."
    
    # Usually there's only one Survey Point
    survey_point = survey_points[0]
    
    # Get the parameter values
    ns_value = survey_point.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble()
    ew_value = survey_point.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsDouble()
    elev_value = survey_point.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsDouble()
    
    # Convert to selected unit
    ns_converted = convert_value(ns_value, unit)
    ew_converted = convert_value(ew_value, unit)
    elev_converted = convert_value(elev_value, unit)
    
    # Check if we can get the angle as well
    angle_param = survey_point.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM)
    angle_value = angle_param.AsDouble() if angle_param else None
    
    # Convert angle to degrees if available
    if angle_value is not None:
        angle_degrees = math.degrees(angle_value)
    else:
        angle_degrees = None
    
    sp_info = {
        "North/South ({0})".format(unit): "{:.4f}".format(ns_converted),
        "East/West ({0})".format(unit): "{:.4f}".format(ew_converted),
        "Elevation ({0})".format(unit): "{:.4f}".format(elev_converted)
    }
    
    if angle_degrees is not None:
        sp_info["Angle to True North (degrees)"] = "{:.4f}".format(angle_degrees)
    
    return sp_info

def get_project_position(unit):
    """Retrieves the Project Position information."""
    project_loc = doc.ActiveProjectLocation
    project_position = project_loc.GetProjectPosition(XYZ(0, 0, 0))
    
    if not project_position:
        return "No Project Position information available."
    
    # Convert values to selected unit
    ew_converted = convert_value(project_position.EastWest, unit)
    ns_converted = convert_value(project_position.NorthSouth, unit)
    elev_converted = convert_value(project_position.Elevation, unit)
    
    # Convert angle to degrees
    angle_degrees = math.degrees(project_position.Angle)
    
    return {
        "East/West ({0})".format(unit): "{:.4f}".format(ew_converted),
        "North/South ({0})".format(unit): "{:.4f}".format(ns_converted),
        "Elevation ({0})".format(unit): "{:.4f}".format(elev_converted),
        "Angle to True North (degrees)": "{:.4f}".format(angle_degrees)
    }

def print_properties(title, properties):
    """Prints properties with consistent formatting."""
    output.print_md("### " + title)
    
    if isinstance(properties, str):
        output.print_md(properties)
    else:
        for key, value in properties.items():
            output.print_md("**{}:** {}".format(key, value))
    
    output.print_md("---")

def main():
    """Main function to execute the script."""
    # Get user's unit preference
    unit = get_unit_selection()
    
    if not unit:
        return  # User canceled
    
    output.print_md("# Project Coordinate Information")
    output.print_md("Project: **{}**".format(doc.Title))
    output.print_md("Selected Unit: **{}**".format(unit))
    output.print_md("---")
    
    # Get Project Base Point information
    pbp_info = get_project_base_point(unit)
    print_properties("Project Base Point", pbp_info)
    
    # Get Survey Point information
    sp_info = get_survey_point(unit)
    print_properties("Survey/Shared Point", sp_info)
    
    # Get the project location information
    proj_pos_info = get_project_position(unit)
    print_properties("Project Position", proj_pos_info)

if __name__ == "__main__":
    main()
