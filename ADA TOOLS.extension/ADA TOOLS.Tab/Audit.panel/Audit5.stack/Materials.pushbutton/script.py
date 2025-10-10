"""
List All Materials
Shows all materials in the project and identifies user-created materials.
"""
__title__ = "List Materials"
__author__ = "Assistant"

from Autodesk.Revit.DB import FilteredElementCollector, Material
from pyrevit import revit, script

# Get the current document
doc = revit.doc
output = script.get_output()

try:
    # Collect all materials
    materials = FilteredElementCollector(doc).OfClass(Material).ToElements()
    materials = sorted(materials, key=lambda m: m.Name)
    
    # ID Thresholds - adjust these based on your observations
    SYSTEM_ID_THRESHOLD = 500000
    
    user_materials = []
    system_materials = []
    
    for material in materials:
        mat_id = material.Id.IntegerValue
        mat_name = material.Name
        
        # Simple heuristic: High IDs are typically user-created
        if mat_id > SYSTEM_ID_THRESHOLD:
            user_materials.append(material)
        else:
            system_materials.append(material)
    
    # Print Results
    output.print_md("# Materials Analysis Report")
    output.print_md("---")
    
    # Summary
    total = len(materials)
    user_count = len(user_materials)
    system_count = len(system_materials)
    
    output.print_md("## Summary Statistics")
    output.print_md("")
    output.print_md("- **Total Materials:** {0}".format(total))
    output.print_md("- **User-Created (ID > {0}):** {1} ({2}%)".format(
        SYSTEM_ID_THRESHOLD,
        user_count, 
        round((user_count*100.0/total), 1) if total > 0 else 0))
    output.print_md("- **System/Template (ID <= {0}):** {1} ({2}%)".format(
        SYSTEM_ID_THRESHOLD,
        system_count,
        round((system_count*100.0/total), 1) if total > 0 else 0))
    output.print_md("")
    output.print_md("---")
    
    # User Materials
    if user_materials:
        output.print_md("## USER-CREATED MATERIALS ({0})".format(user_count))
        output.print_md("")
        user_materials_sorted = sorted(user_materials, key=lambda x: x.Id.IntegerValue, reverse=True)
        for idx, mat in enumerate(user_materials_sorted, 1):
            output.print_md("{0}. {1} **{2}** - ID: {3}".format(
                idx,
                output.linkify(mat.Id),
                mat.Name,
                mat.Id.IntegerValue
            ))
        output.print_md("")
        output.print_md("---")
    
    # System Materials (show first 50)
    if system_materials:
        output.print_md("## SYSTEM/TEMPLATE MATERIALS ({0})".format(system_count))
        output.print_md("")
        system_materials_sorted = sorted(system_materials, key=lambda x: x.Id.IntegerValue)
        display_count = min(50, system_count)
        
        for idx, mat in enumerate(system_materials_sorted[:display_count], 1):
            output.print_md("{0}. {1} **{2}** - ID: {3}".format(
                idx,
                output.linkify(mat.Id),
                mat.Name,
                mat.Id.IntegerValue
            ))
        
        if system_count > display_count:
            output.print_md("")
            output.print_md("*...and {0} more system materials*".format(system_count - display_count))
        
        output.print_md("")
        output.print_md("---")
    
    # Show ID distribution
    output.print_md("## ID Distribution")
    output.print_md("")
    if materials:
        sorted_mats = sorted(materials, key=lambda x: x.Id.IntegerValue)
        output.print_md("**Lowest Material ID:** {0}".format(sorted_mats[0].Id.IntegerValue))
        output.print_md("**Highest Material ID:** {0}".format(sorted_mats[-1].Id.IntegerValue))
    output.print_md("")
    output.print_md("*Threshold used: {0}*".format(SYSTEM_ID_THRESHOLD))
    output.print_md("")
    output.print_md("*Adjust SYSTEM_ID_THRESHOLD in the script if needed based on the ID distribution above.*")

except Exception as e:
    import traceback
    output.print_md("# Error occurred:")
    output.print_md("```")
    output.print_md(str(e))
    output.print_md("")
    output.print_md(traceback.format_exc())
    output.print_md("```")
