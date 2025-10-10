"""
List All Materials
Shows all materials in the project and identifies which ones are actually used.
"""
__title__ = "List Materials"
__author__ = "Almog Davidson"

from Autodesk.Revit.DB import (FilteredElementCollector, Material, ElementId, 
                                BuiltInParameter, FamilyInstance, GeometryElement,
                                Options, Solid, Face, StorageType, ViewDetailLevel)
from pyrevit import revit, script

# Get the current document
doc = revit.doc
output = script.get_output()

try:
    # Collect all materials
    all_materials = FilteredElementCollector(doc).OfClass(Material).ToElements()
    
    # Create a dictionary to track which INSTANCES use which materials
    # Key: material_id, Value: set of element_ids (instances only)
    material_elements = {}
    for mat in all_materials:
        material_elements[mat.Id.IntegerValue] = set()
    
    output.print_md("*Scanning project for material usage (this may take a moment)...*")
    output.print_md("")
    
    # Get all element instances only (not types)
    all_instances = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
    
    # Geometry options for extracting materials from geometry
    geo_options = Options()
    geo_options.DetailLevel = ViewDetailLevel.Fine
    geo_options.IncludeNonVisibleObjects = True
    
    # Check all instances
    output.print_md("*Checking element instances...*")
    for elem in all_instances:
        elem_id = elem.Id.IntegerValue
        materials_found_in_elem = set()  # Track unique materials in this element
        
        # Check Material parameter on instance
        mat_param = elem.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
        if mat_param and mat_param.HasValue:
            mat_id = mat_param.AsElementId()
            if mat_id != ElementId.InvalidElementId and mat_id.IntegerValue in material_elements:
                materials_found_in_elem.add(mat_id.IntegerValue)
        
        # Check Structural Material parameter on instance
        struct_mat_param = elem.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
        if struct_mat_param and struct_mat_param.HasValue:
            mat_id = struct_mat_param.AsElementId()
            if mat_id != ElementId.InvalidElementId and mat_id.IntegerValue in material_elements:
                materials_found_in_elem.add(mat_id.IntegerValue)
        
        # Check type's material parameters
        elem_type = doc.GetElement(elem.GetTypeId())
        if elem_type:
            # Check type Material parameter
            type_mat_param = elem_type.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
            if type_mat_param and type_mat_param.HasValue:
                mat_id = type_mat_param.AsElementId()
                if mat_id != ElementId.InvalidElementId and mat_id.IntegerValue in material_elements:
                    materials_found_in_elem.add(mat_id.IntegerValue)
            
            # Check type Structural Material parameter
            type_struct_mat_param = elem_type.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
            if type_struct_mat_param and type_struct_mat_param.HasValue:
                mat_id = type_struct_mat_param.AsElementId()
                if mat_id != ElementId.InvalidElementId and mat_id.IntegerValue in material_elements:
                    materials_found_in_elem.add(mat_id.IntegerValue)
            
            # Check for compound structures in the type
            if hasattr(elem_type, 'GetCompoundStructure'):
                try:
                    compound = elem_type.GetCompoundStructure()
                    if compound:
                        for layer in compound.GetLayers():
                            mat_id = layer.MaterialId
                            if mat_id != ElementId.InvalidElementId and mat_id.IntegerValue in material_elements:
                                materials_found_in_elem.add(mat_id.IntegerValue)
                except:
                    pass
            
            # Check all type parameters for materials
            for param in elem_type.Parameters:
                if param.StorageType == StorageType.ElementId and param.HasValue:
                    param_elem_id = param.AsElementId()
                    if param_elem_id != ElementId.InvalidElementId:
                        try:
                            param_elem = doc.GetElement(param_elem_id)
                            if isinstance(param_elem, Material) and param_elem_id.IntegerValue in material_elements:
                                materials_found_in_elem.add(param_elem_id.IntegerValue)
                        except:
                            pass
        
        # Check for paint materials on faces
        try:
            if hasattr(elem, 'GetMaterialIds'):
                paint_mat_ids = elem.GetMaterialIds(True)  # True = get painted materials
                for mat_id in paint_mat_ids:
                    if mat_id != ElementId.InvalidElementId and mat_id.IntegerValue in material_elements:
                        materials_found_in_elem.add(mat_id.IntegerValue)
        except:
            pass
        
        # Check all instance parameters for materials
        for param in elem.Parameters:
            if param.StorageType == StorageType.ElementId and param.HasValue:
                param_elem_id = param.AsElementId()
                if param_elem_id != ElementId.InvalidElementId:
                    try:
                        param_elem = doc.GetElement(param_elem_id)
                        if isinstance(param_elem, Material) and param_elem_id.IntegerValue in material_elements:
                            materials_found_in_elem.add(param_elem_id.IntegerValue)
                    except:
                        pass
        
        # Check geometry for materials
        try:
            geom_elem = elem.get_Geometry(geo_options)
            if geom_elem:
                for geom_obj in geom_elem:
                    # Check solids
                    if isinstance(geom_obj, Solid) and geom_obj.Faces.Size > 0:
                        for face in geom_obj.Faces:
                            mat_id = face.MaterialElementId
                            if mat_id != ElementId.InvalidElementId and mat_id.IntegerValue in material_elements:
                                materials_found_in_elem.add(mat_id.IntegerValue)
                    
                    # Check geometry instances (for families)
                    elif isinstance(geom_obj, GeometryInstance):
                        inst_geom = geom_obj.GetInstanceGeometry()
                        if inst_geom:
                            for inst_obj in inst_geom:
                                if isinstance(inst_obj, Solid) and inst_obj.Faces.Size > 0:
                                    for face in inst_obj.Faces:
                                        mat_id = face.MaterialElementId
                                        if mat_id != ElementId.InvalidElementId and mat_id.IntegerValue in material_elements:
                                            materials_found_in_elem.add(mat_id.IntegerValue)
        except:
            pass
        
        # Add this element ID to all materials found in it
        for mat_id in materials_found_in_elem:
            material_elements[mat_id].add(elem_id)
    
    output.print_md("*Analysis complete!*")
    output.print_md("")
    output.print_md("---")
    output.print_md("")
    
    # Separate used and unused materials
    used_materials = [(mat, len(material_elements[mat.Id.IntegerValue])) 
                      for mat in all_materials 
                      if len(material_elements[mat.Id.IntegerValue]) > 0]
    unused_materials = [mat for mat in all_materials 
                        if len(material_elements[mat.Id.IntegerValue]) == 0]
    
    # Statistics
    total = len(all_materials)
    used_count = len(used_materials)
    unused_count = len(unused_materials)
    
    output.print_md("## Summary Statistics")
    output.print_md("")
    output.print_md("- **Total Materials:** {0}".format(total))
    output.print_md("- **Used Materials:** {0} ({1}%)".format(
        used_count, 
        round((used_count*100.0/total), 1) if total > 0 else 0))
    output.print_md("- **Unused Materials:** {0} ({1}%)".format(
        unused_count,
        round((unused_count*100.0/total), 1) if total > 0 else 0))
    output.print_md("")
    output.print_md("---")
    
    # Used Materials
    if used_materials:
        output.print_md("## USED MATERIALS ({0})".format(used_count))
        output.print_md("")
        output.print_md("*Count shows number of placed element instances using each material.*")
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
    
    # Unused Materials
    if unused_materials:
        output.print_md("## UNUSED MATERIALS ({0})".format(unused_count))
        output.print_md("")
        output.print_md("*These materials are not assigned to any placed elements and can potentially be deleted.*")
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

except Exception as e:
    import traceback
    output.print_md("# Error occurred:")
    output.print_md("```")
    output.print_md(str(e))
    output.print_md("")
    output.print_md(traceback.format_exc())
    output.print_md("```")
