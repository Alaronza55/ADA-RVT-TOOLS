"""
List All Materials
Shows all materials in the project and identifies which ones are actually used.
This is the most robust version with multiple verification methods.
"""
__title__ = "List Materials"
__author__ = "Almog Davidson"

from Autodesk.Revit.DB import (FilteredElementCollector, Material, ElementId, 
                                BuiltInParameter, FamilyInstance, GeometryElement,
                                Options, Solid, Face, StorageType, ViewDetailLevel,
                                GeometryInstance, HostObjAttributes)
from pyrevit import revit, script
import os
import csv

# Get the current document
doc = revit.doc
output = script.get_output()

def get_materials_from_element(elem, doc, geo_options):
    """Extract all materials used by an element - most comprehensive check"""
    materials = set()
    
    # 1. Check direct Material parameter on instance
    mat_param = elem.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
    if mat_param and mat_param.HasValue:
        mat_id = mat_param.AsElementId()
        if mat_id != ElementId.InvalidElementId:
            materials.add(mat_id.IntegerValue)
    
    # 2. Check Structural Material parameter on instance
    struct_mat_param = elem.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
    if struct_mat_param and struct_mat_param.HasValue:
        mat_id = struct_mat_param.AsElementId()
        if mat_id != ElementId.InvalidElementId:
            materials.add(mat_id.IntegerValue)
    
    # 3. Get the element's type
    elem_type = doc.GetElement(elem.GetTypeId())
    if elem_type:
        # 3a. Check type Material parameter
        type_mat_param = elem_type.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
        if type_mat_param and type_mat_param.HasValue:
            mat_id = type_mat_param.AsElementId()
            if mat_id != ElementId.InvalidElementId:
                materials.add(mat_id.IntegerValue)
        
        # 3b. Check type Structural Material parameter
        type_struct_mat_param = elem_type.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
        if type_struct_mat_param and type_struct_mat_param.HasValue:
            mat_id = type_struct_mat_param.AsElementId()
            if mat_id != ElementId.InvalidElementId:
                materials.add(mat_id.IntegerValue)
        
        # 3c. Check compound structure layers (walls, floors, roofs, ceilings)
        if isinstance(elem_type, HostObjAttributes):
            try:
                compound = elem_type.GetCompoundStructure()
                if compound:
                    for layer in compound.GetLayers():
                        mat_id = layer.MaterialId
                        if mat_id != ElementId.InvalidElementId:
                            materials.add(mat_id.IntegerValue)
            except:
                pass
        
        # 3d. Check ALL parameters on type that might reference materials
        for param in elem_type.Parameters:
            if param.StorageType == StorageType.ElementId and param.HasValue:
                param_elem_id = param.AsElementId()
                if param_elem_id != ElementId.InvalidElementId:
                    try:
                        param_elem = doc.GetElement(param_elem_id)
                        if isinstance(param_elem, Material):
                            materials.add(param_elem_id.IntegerValue)
                    except:
                        pass
    
    # 4. Check ALL parameters on instance that might reference materials
    for param in elem.Parameters:
        if param.StorageType == StorageType.ElementId and param.HasValue:
            param_elem_id = param.AsElementId()
            if param_elem_id != ElementId.InvalidElementId:
                try:
                    param_elem = doc.GetElement(param_elem_id)
                    if isinstance(param_elem, Material):
                        materials.add(param_elem_id.IntegerValue)
                except:
                    pass
    
    # 5. Check painted materials on faces
    try:
        if hasattr(elem, 'GetMaterialIds'):
            paint_mat_ids = elem.GetMaterialIds(True)  # True = get painted materials only
            for mat_id in paint_mat_ids:
                if mat_id != ElementId.InvalidElementId:
                    materials.add(mat_id.IntegerValue)
    except:
        pass
    
    # 6. Check geometry for materials (most thorough)
    try:
        geom_elem = elem.get_Geometry(geo_options)
        if geom_elem:
            materials.update(extract_materials_from_geometry(geom_elem))
    except:
        pass
    
    return materials

def extract_materials_from_geometry(geom_elem):
    """Recursively extract materials from geometry"""
    materials = set()
    
    try:
        for geom_obj in geom_elem:
            # Check solids
            if isinstance(geom_obj, Solid):
                if geom_obj.Faces.Size > 0:
                    for face in geom_obj.Faces:
                        mat_id = face.MaterialElementId
                        if mat_id != ElementId.InvalidElementId:
                            materials.add(mat_id.IntegerValue)
            
            # Check faces directly
            elif isinstance(geom_obj, Face):
                mat_id = geom_obj.MaterialElementId
                if mat_id != ElementId.InvalidElementId:
                    materials.add(mat_id.IntegerValue)
            
            # Check geometry instances (for families) - RECURSIVE
            elif isinstance(geom_obj, GeometryInstance):
                # Get instance geometry
                inst_geom = geom_obj.GetInstanceGeometry()
                if inst_geom:
                    materials.update(extract_materials_from_geometry(inst_geom))
                
                # Also get symbol geometry
                symbol_geom = geom_obj.GetSymbolGeometry()
                if symbol_geom:
                    materials.update(extract_materials_from_geometry(symbol_geom))
    except:
        pass
    
    return materials

try:
    # Collect all materials in the project
    all_materials = FilteredElementCollector(doc).OfClass(Material).ToElements()
    material_ids = set(mat.Id.IntegerValue for mat in all_materials)
    
    # Create a dictionary to track which instances use which materials
    material_elements = {mat_id: set() for mat_id in material_ids}
    
    output.print_md("# Material Usage Analysis")
    output.print_md("")
    output.print_md("*Performing comprehensive scan of all placed elements...*")
    output.print_md("*This may take a few moments for large projects.*")
    output.print_md("")
    
    # Get all placed element instances (not types)
    all_instances = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
    
    # Geometry options for material extraction
    geo_options = Options()
    geo_options.DetailLevel = ViewDetailLevel.Fine
    geo_options.IncludeNonVisibleObjects = True
    geo_options.ComputeReferences = False  # Faster
    
    # Progress tracking
    total_elements = len(all_instances)
    processed = 0
    
    # Check each placed element
    for elem in all_instances:
        processed += 1
        if processed % 1000 == 0:
            output.print_md("*Processed {0}/{1} elements...*".format(processed, total_elements))
        
        elem_id = elem.Id.IntegerValue
        
        # Get all materials used by this element
        elem_materials = get_materials_from_element(elem, doc, geo_options)
        
        # Add this element to each material's usage list
        for mat_id in elem_materials:
            if mat_id in material_elements:
                material_elements[mat_id].add(elem_id)
    
    output.print_md("*Analysis complete!*")
    output.print_md("")
    output.print_md("---")
    output.print_md("")
    
    # Separate used and unused materials
    used_materials = []
    unused_materials = []
    
    for mat in all_materials:
        mat_id = mat.Id.IntegerValue
        count = len(material_elements[mat_id])
        if count > 0:
            used_materials.append((mat, count))
        else:
            unused_materials.append(mat)
    
    # Statistics
    total = len(all_materials)
    used_count = len(used_materials)
    unused_count = len(unused_materials)
    total_element_count = sum(count for _, count in used_materials)
    
    output.print_md("## Summary Statistics")
    output.print_md("")
    output.print_md("- **Total Materials in Project:** {0}".format(total))
    output.print_md("- **Materials Actually Used:** {0} ({1}%)".format(
        used_count, 
        round((used_count*100.0/total), 1) if total > 0 else 0))
    output.print_md("- **Unused Materials:** {0} ({1}%)".format(
        unused_count,
        round((unused_count*100.0/total), 1) if total > 0 else 0))
    output.print_md("- **Total Placed Elements Analyzed:** {0}".format(total_elements))
    output.print_md("")
    output.print_md("---")
    output.print_md("")
    
    # Used Materials
    if used_materials:
        output.print_md("## USED MATERIALS ({0})".format(used_count))
        output.print_md("")
        output.print_md("*Shows number of placed element instances using each material.*")
        output.print_md("*Click material ID to select it in the project.*")
        output.print_md("")
        
        # Sort by usage count (descending), then by name
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
        output.print_md("*They may be used in unplaced families or can be safely deleted.*")
        output.print_md("*Click material ID to select it in the project.*")
        output.print_md("")
        
        # Sort alphabetically
        unused_materials_sorted = sorted(unused_materials, key=lambda x: x.Name)
        
        for idx, mat in enumerate(unused_materials_sorted, 1):
            output.print_md("{0}. {1} **{2}**".format(
                idx,
                output.linkify(mat.Id),
                mat.Name
            ))
        output.print_md("")
        output.print_md("---")

        # Export to CSV
    folder_name = doc.Title
    output_folder = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker\{}".format(folder_name)

    # Create the folder if it doesn't exist
    if not os.path.exists(output_folder):
        try:
            os.makedirs(output_folder)
            print("**Created folder:** `{}`".format(output_folder))
        except Exception as e:
            print("**Error creating folder:** {}".format(str(e)))
            print("**Attempting to save to default location...**")
            output_folder = os.path.expanduser("~\\Desktop")

    # Save detailed element breakdown
    filename_detailed = "Material_Usage.csv"
    filepath_detailed = os.path.join(output_folder, filename_detailed)
    try:
        with open(filepath_detailed, 'wb') as csvfile:
            csvwriter = csv.writer(csvfile)
            # Write header
            csvwriter.writerow(['Material Name', 'Used', 'Count'])
            
            # Write used materials
            for mat, count in used_materials:
                csvwriter.writerow([mat.Name, 'Yes', count])
            
            # Write unused materials
            for mat in unused_materials:
                csvwriter.writerow([mat.Name, 'No', 0])
        
        output.print_md("## CSV Export Complete")
        output.print_md("")
        output.print_md("*A CSV file has been exported to: **{0}**".format(filepath_detailed))
        output.print_md("*Columns: 'Material Name', 'Used' (Yes/No), and 'Count' (number of elements using the material).")
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
