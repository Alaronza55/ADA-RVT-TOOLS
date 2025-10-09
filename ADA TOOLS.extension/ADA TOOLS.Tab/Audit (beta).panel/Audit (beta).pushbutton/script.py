"""Retrieve Survey Point Coordinates from Project Information
Gets the values of B6_Survey Point Lat and B6_Survey Point Long parameters"""

__title__ = "Get Survey\nCoordinates"
__author__ = "Your Name"

from pyrevit import revit, DB
from pyrevit import script

# Get the current document
doc = revit.doc
output = script.get_output()

# Parameter names to retrieve
param_lat = "B6_Survey Point Lat"
param_long = "B6_Survey Point Long"

# Get Project Information
proj_info = doc.ProjectInformation

lat_value = None
long_value = None

# Iterate through all parameters in Project Information
for param in proj_info.Parameters:
    param_name = param.Definition.Name
    
    if param_name == param_lat:
        if param.HasValue:
            # Get the value as string
            lat_value = param.AsString()
    
    elif param_name == param_long:
        if param.HasValue:
            # Get the value as string
            long_value = param.AsString()

# Display results
output.print_md("## Survey Point Coordinates")
output.print_md("---")

if lat_value:
    output.print_md("**{}:** {}".format(param_lat, lat_value))
else:
    output.print_md("**{}:** *Not found or empty*".format(param_lat))

if long_value:
    output.print_md("**{}:** {}".format(param_long, long_value))
else:
    output.print_md("**{}:** *Not found or empty*".format(param_long))

# Copy to clipboard if both values exist
if lat_value and long_value:
    coords = "Lat: {}, Long: {}".format(lat_value, long_value)
    output.print_md("\n---")
    output.print_md("**Combined:** `{}`".format(coords))
    
    # Uncomment to copy to clipboard automatically
    # script.clipboard_copy(coords)
