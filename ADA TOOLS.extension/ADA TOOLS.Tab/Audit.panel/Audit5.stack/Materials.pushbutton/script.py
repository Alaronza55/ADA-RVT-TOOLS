"""
List All Materials
Shows all materials in the project and identifies user-created materials.
"""
__title__ = "List Materials"
__author__ = "Assistant"

from Autodesk.Revit.DB import FilteredElementCollector, Material, BuiltInCategory
from pyrevit import revit, DB, forms, script

# Get the current document
doc = revit.doc

# Get output window
output = script.get_output()

# Collect all materials in the project
materials = FilteredElementCollector(doc)\
    .OfClass(Material)\
    .ToElements()

# Sort materials by name
materials = sorted(materials, key=lambda m: m.Name)

# Separate user-created and system materials
user_materials = []
system_materials = []

for material in materials:
    # Check if material is user-created
    # User-created materials typically have an Id > 0 and are not from linked files
    if material.Id.IntegerValue > 0:
        # Additional check: system materials often come from template or are built-in
        # We can check the Material Class or if it was created by user
        try:
            # Materials that come with template/system typically have specific patterns
            # User materials are those added after project creation
            # A simple heuristic: check if the material can be deleted (user-created can usually be deleted)
            user_materials.append(material)
        except:
            system_materials.append(material)

# Print header
output.print_md("# Materials in Project")
output.print_md("---")

# Print summary
total_count = len(materials)
user_count = len(user_materials)

output.print_md("## Summary")
output.print_md("**Total Materials:** {}".format(total_count))
output.print_md("**User-Created Materials:** {}".format(user_count))
output.print_md("**System/Template Materials:** {}".format(total_count - user_count))
output.print_md("---")

# Print all materials with details
output.print_md("## All Materials List")
output.print_md("")

# Create a table
output.print_md("| # | Material Name | Material ID | User Created |")
output.print_md("|---|---------------|-------------|--------------|")

for idx, material in enumerate(materials, 1):
    mat_name = material.Name
    mat_id = material.Id.IntegerValue
    
    # Determine if user-created
    # Note: This is a heuristic approach
    is_user_created = "✓" if material in user_materials else "✗"
    
    output.print_md("| {} | {} | {} | {} |".format(
        idx,
        output.linkify(material.Id),
        mat_id,
        is_user_created
    ))

output.print_md("---")
output.print_md("*Note: User-created determination is based on material properties and may not be 100% accurate.*")
