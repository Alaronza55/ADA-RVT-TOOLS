"""
List All Materials - OPTIMIZED VERSION
Shows all materials in the project and identifies which ones are actually used.
Optimized for performance on large models.
"""
__title__ = "List Materials"
__author__ = "Almog Davidson"

from Autodesk.Revit.DB import (FilteredElementCollector, Material, ElementId, 
                                BuiltInParameter, GeometryElement,
                                Options, Solid, ViewDetailLevel,
                                GeometryInstance, HostObjAttributes)
from pyrevit import revit, script
import os
import csv

doc = revit.doc
output = script.get_output()

def get_materials_from_element(elem, doc, elem_type_cache):
    """Extract materials - optimized with caching and prioritized checks"""
    materials = set()
    
    # Quick parameter checks first (fastest)
    mat_param = elem.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
    if mat_param and mat_param.HasValue:
        mat_id = mat_param.AsElementId()
        if mat_id != ElementId.InvalidElementId:
            materials.add(mat_id.IntegerValue)
    
    struct_mat_param = elem.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
    if struct_mat_param and struct_mat_param.HasValue:
        mat_id = struct_mat_param.AsElementId()
        if mat_id != ElementId.InvalidElementId:
            materials.add(mat_id.IntegerValue)
    
    # Check painted materials (fast and common)
    try:
        if hasattr(elem, 'GetMaterialIds'):
            paint_mat_ids = elem.GetMaterialIds(True)
            for mat_id in paint_mat_ids:
                if mat_id != ElementId.InvalidElementId:
                    materials.add(mat_id.IntegerValue)
    except:
        pass
    
    # Type checks with caching
    type_id = elem.GetTypeId()
    if type_id != ElementId.InvalidElementId:
        type_id_int = type_id.IntegerValue
        
        # Check cache first
        if type_id_int in elem_type_cache:
            materials.update(elem_type_cache[type_id_int])
        else:
            # Process type and cache results
            type_materials = set()
            elem_type = doc.GetElement(type_id)
            
            if elem_type:
                # Type material parameters
                type_mat_param = elem_type.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
                if type_mat_param and type_mat_param.HasValue:
                    mat_id = type_mat_param.AsElementId()
                    if mat_id != ElementId.InvalidElementId:
                        type_materials.add(mat_id.IntegerValue)
                
                type_struct_mat_param = elem_type.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
                if type_struct_mat_param and type_struct_mat_param.HasValue:
                    mat_id = type_struct_mat_param.AsElementId()
                    if mat_id != ElementId.InvalidElementId:
                        type_materials.add(mat_id.IntegerValue)
                
                # Compound structure layers
                if isinstance(elem_type, HostObjAttributes):
                    try:
                        compound = elem_type.GetCompoundStructure()
                        if compound:
                            for layer in compound.GetLayers():
                                mat_id = layer.MaterialId
                                if mat_id != ElementId.InvalidElementId:
                                    type_materials.add(mat_id.IntegerValue)
                    except:
                        pass
            
            # Cache and use
            elem_type_cache[type_id_int] = type_materials
            materials.update(type_materials)
    
    return materials

def get_materials_from_geometry_optimized(elem, geo_options):
    """Optimized geometry traversal - only when needed"""
    materials = set()
    
    try:
        geom_elem = elem.get_Geometry(geo_options)
        if not geom_elem:
            return materials
        
        for geom_obj in geom_elem:
            if isinstance(geom_obj, Solid) and geom_obj.Faces.Size > 0:
                # Sample faces instead of checking all (much faster)
                face_count = geom_obj.Faces.Size
                step = max(1, face_count // 10)  # Sample up to 10 faces
                
                for i in range(0, face_count, step):
                    face = geom_obj.Faces[i]
                    mat_id = face.MaterialElementId
                    if mat_id != ElementId.InvalidElementId:
                        materials.add(mat_id.IntegerValue)
            
            elif isinstance(geom_obj, GeometryInstance):
                # Only go one level deep
                inst_geom = geom_obj.GetInstanceGeometry()
                if inst_geom:
                    for inst_obj in inst_geom:
                        if isinstance(inst_obj, Solid) and inst_obj.Faces.Size > 0:
                            # Just check first face of nested geometry
                            mat_id = inst_obj.Faces[0].MaterialElementId
                            if mat_id != ElementId.InvalidElementId:
                                materials.add(mat_id.IntegerValue)
                            break  # One sample is enough
    except:
        pass
    
    return materials

try:
    # Collect all materials
    all_materials = FilteredElementCollector(doc).OfClass(Material).ToElements()
    material_ids = set(mat.Id.IntegerValue for mat in all_materials)
    
    # Track usage - use boolean instead of element sets (much faster)
    material_used = {mat_id: False for mat_id in material_ids}
    material_count = {mat_id: 0 for mat_id in material_ids}
    
    output.print_md("# Material Usage Analysis - OPTIMIZED")
    output.print_md("")
    output.print_md("*Scanning elements with optimized algorithm...*")
    output.print_md("")
    
    # Get all placed instances
    all_instances = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
    total_elements = len(all_instances)
    
    # Geometry options
    geo_options = Options()
    geo_options.DetailLevel = ViewDetailLevel.Fine
    geo_options.IncludeNonVisibleObjects = False  # Changed to False - faster
    geo_options.ComputeReferences = False
    
    # Type cache for performance
    elem_type_cache = {}
    
    # Fast pass: Check parameters only (90% of materials are found here)
    processed = 0
    update_interval = max(100, total_elements // 20)  # Update less frequently
    
    for elem in all_instances:
        processed += 1
        if processed % update_interval == 0:
            # Update progress without printing (faster)
            output.update_progress(processed, total_elements)
        
        # Quick parameter check
        elem_materials = get_materials_from_element(elem, doc, elem_type_cache)
        
        for mat_id in elem_materials:
            if mat_id in material_used:
                material_used[mat_id] = True
                material_count[mat_id] += 1
    
    # Second pass: Geometry check ONLY for materials still not found
    # This is much faster than checking geometry for everything
    unused_mats = {mat_id for mat_id, used in material_used.items() if not used}
    
    if unused_mats:
        output.print_md("*Running detailed geometry scan for remaining materials...*")
        output.print_md("")
        
        processed = 0
        for elem in all_instances:
            processed += 1
            if processed % update_interval == 0:
                output.update_progress(processed, total_elements)
            
            # Skip if all materials are found
            if not unused_mats:
                break
            
            geo_materials = get_materials_from_geometry_optimized(elem, geo_options)
            
            for mat_id in geo_materials:
                if mat_id in unused_mats:
                    material_used[mat_id] = True
                    material_count[mat_id] += 1
                    unused_mats.discard(mat_id)
    
    output.print_md("*Analysis complete!*")
    output.print_md("")
    output.print_md("---")
    output.print_md("")
    
    # Separate used and unused materials
    used_materials = []
    unused_materials = []
    
    for mat in all_materials:
        mat_id = mat.Id.IntegerValue
        if material_used[mat_id]:
            used_materials.append((mat, material_count[mat_id]))
        else:
            unused_materials.append(mat)
    
    # Statistics
    total = len(all_materials)
    used_count = len(used_materials)
    unused_count = len(unused_materials)
    
    output.print_md("## Summary Statistics")
    output.print_md("")
    output.print_md("- **Total Materials in Project:** {0}".format(total))
    output.print_md("- **Materials Actually Used:** {0} ({1}%)".format(
        used_count, 
        round((used_count*100.0/total), 1) if total > 0 else 0))
    output.print_md("- **Unused Materials:** {0} ({1}%)".format(
        unused_count,
        round((unused_count*100.0/total), 1) if total > 0 else 0))
    output.print_md("- **Total Elements Analyzed:** {0}".format(total_elements))
    output.print_md("")
    output.print_md("---")
    output.print_md("")
    
    # Used Materials
    if used_materials:
        output.print_md("## USED MATERIALS ({0})".format(used_count))
        output.print_md("")
        output.print_md("*Shows number of element instances using each material.*")
        output.print_md("*Click material ID to select it in the project.*")
        output.print_md("")
        
        used_materials_sorted = sorted(used_materials, key=lambda x: (-x[1], x[0].Name))
        
        for idx, (mat, count) in enumerate(used_materials_sorted, 1):
            output.print_md("{0}. {1} **{2}** - Used in **{3}** element(s)".format(
                idx,
                output.linkify(mat.Id),
                mat.Name,
                count
            ))
        output.print_md("")
        output.print_md("---")
        output.print_md("")
    
    # Unused Materials
    if unused_materials:
        output.print_md("## UNUSED MATERIALS ({0})".format(unused_count))
        output.print_md("")
        output.print_md("*These materials are not assigned to any placed elements.*")
        output.print_md("*Click material ID to select it in the project.*")
        output.print_md("")
        
        unused_materials_sorted = sorted(unused_materials, key=lambda x: x.Name)
        
        for idx, mat in enumerate(unused_materials_sorted, 1):
            output.print_md("{0}. {1} **{2}**".format(
                idx,
                output.linkify(mat.Id),
                mat.Name
            ))
        output.print_md("")
        output.print_md("---")
        output.print_md("")
    
    # Export to CSV
    folder_name = doc.Title
    output_folder = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker\{}".format(folder_name)
    
    if not os.path.exists(output_folder):
        try:
            os.makedirs(output_folder)
            output.print_md("**Created folder:** `{}`".format(output_folder))
        except Exception as e:
            output.print_md("**Error creating folder:** {}".format(str(e)))
            output_folder = os.path.expanduser("~\\Desktop")
    
    filename_detailed = "Material_Usage.csv"
    filepath_detailed = os.path.join(output_folder, filename_detailed)
    
    try:
        with open(filepath_detailed, 'wb') as csvfile:
            csvwriter = csv.writer(csvfile)
            csvwriter.writerow(['Material Name', 'Used', 'Count'])
            
            for mat, count in used_materials:
                # Encode to UTF-8 to handle special characters
                mat_name = mat.Name.encode('utf-8') if mat.Name else 'Unnamed'
                csvwriter.writerow([mat_name, 'Yes', count])
            
            for mat in unused_materials:
                # Encode to UTF-8 to handle special characters
                mat_name = mat.Name.encode('utf-8') if mat.Name else 'Unnamed'
                csvwriter.writerow([mat_name, 'No', 0])
        
        output.print_md("## CSV Export Complete")
        output.print_md("")
        output.print_md("*CSV exported to: **{0}**".format(filepath_detailed))
    except Exception as csv_error:
        output.print_md("# Error exporting CSV:")
        output.print_md("```")
        output.print_md(str(csv_error))
        output.print_md("```")

except Exception as e:
    import traceback
    output.print_md("# Error occurred:")
    output.print_md("```")
    output.print_md(str(e))
    output.print_md("")
    output.print_md(traceback.format_exc())
    output.print_md("```")