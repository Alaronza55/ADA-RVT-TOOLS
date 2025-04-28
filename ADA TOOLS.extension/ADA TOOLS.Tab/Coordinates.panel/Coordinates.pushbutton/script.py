# Import necessary modules
import clr
import math

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *

# Get the active Revit document
doc = __revit__.ActiveUIDocument.Document

# Function to convert from internal Revit coordinates to displayable values
def convert_internal_to_display_units(value, doc):
    # Convert from internal units (feet) to the project's display units
    display_units = doc.GetUnits().GetFormatOptions(UnitType.UT_Length).DisplayUnits
    return UnitUtils.ConvertFromInternalUnits(value, display_units)

# Find the Project Base Point and Survey Point
base_point = None
survey_point = None

# Use FilteredElementCollector to find the points
collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType()
base_points = list(collector)

collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_SharedBasePoint).WhereElementIsNotElementType()
survey_points = list(collector)

if len(base_points) > 0:
    base_point = base_points[0]
    
if len(survey_points) > 0:
    survey_point = survey_points[0]

# Function to get point coordinates
def get_point_info(point, point_name):
    if point is None:
        print("{0} not found in the project.".format(point_name))
        return
        
    # Create bold text effect using ANSI escape codes (works in some console environments)
    # Using repeated characters for "bolding" effect (visible in all environments)
    bold_title = "**" + point_name.upper() + "**"
    
    # Print point name with bold formatting
    print("\n")
    print("+" + "-" * (len(bold_title) + 4) + "+")
    print("|  " + bold_title + "  |")
    print("+" + "-" * (len(bold_title) + 4) + "+")
    
    # Try to get the coordinates directly using built-in parameters
    try:
        # Get parameter values properly
        # Note: These parameters are read differently based on how they're stored in Revit
        e = point.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsDouble()
        n = point.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble()
        elev = point.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsDouble()
        
        # Convert to display units
        e_display = convert_internal_to_display_units(e, doc)
        n_display = convert_internal_to_display_units(n, doc)
        elev_display = convert_internal_to_display_units(elev, doc)
        
        # Format with two decimal places and use comma as decimal separator
        e_formatted = "{:.3f}".format(e_display)
        n_formatted = "{:.3f}".format(n_display)
        elev_formatted = "{:.3f}".format(elev_display)
        
        print("\nCOORDINATES:")
        print("  East/West: {0}".format(e_formatted))
        print("  North/South: {0}".format(n_formatted))
        print("  Elevation: {0}".format(elev_formatted))
    except Exception as ex:
        # If can't get coordinates using the standard method, try alternative approach
        print("Could not retrieve standard coordinates. Error: {0}".format(str(ex)))
        print("Trying alternative coordinate retrieval method...")
        
        try:
            # Alternative method - get location directly from element
            location = point.Location
            if location:
                point_loc = location.Point
                x = convert_internal_to_display_units(point_loc.X, doc)
                y = convert_internal_to_display_units(point_loc.Y, doc)
                z = convert_internal_to_display_units(point_loc.Z, doc)
                
                # Format with two decimal places and use comma as decimal separator
                x_formatted = "{:.3f}".format(x)
                y_formatted = "{:.3f}".format(y)
                z_formatted = "{:.3f}".format(z)
                
                print("\nCOORDINATES (Alternative Method):")
                print("  X: {0}".format(x_formatted))
                print("  Y: {0}".format(y_formatted))
                print("  Z: {0}".format(z_formatted))
        except:
            print("Could not retrieve coordinates using alternative methods either.")

# Print a header for the report
print("\n" + "=" * 50)
print("        REVIT PROJECT POINT COORDINATES")
print("=" * 50)

# Get and print information for both points
get_point_info(base_point, "Project Base Point")
get_point_info(survey_point, "Project Survey Point")
