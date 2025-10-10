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
    
    # Create a set to track used material IDs
    used_material_ids = set()
    
    output.print_md("*Scanning project for material usage (this may take a moment)...*")
    output.print_md("")
    
    # Get all element types and instances
    all_types = FilteredElementCollector(doc).WhereElementIsElementType().ToElements()
    all_instances = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
    
    # Geometry options for extracting materials from geometry
    geo_options = Options()
    geo_options.DetailLevel = ViewDetailLevel.Fine
    geo_options.IncludeNonVisibleObjects = True
    
    # Check all types
    output.print_md("*Checking element types...*")
    for elem_type in all_types:
        # Check Material parameter
        mat_param = elem_type.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
        if mat_param and mat_param.HasValue:
            mat_id = mat_param.AsElementId()
            if mat_id != ElementId.InvalidElementId:
                used_material_ids.add(mat_id.IntegerValue)
        
        # Check Structural Material parameter
        struct_mat_param = elem_type.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
        if struct_mat_param and struct_mat_param.HasValue:
            mat_id = struct_mat_param.AsElementId()
            if mat_id != ElementId.InvalidElementId:
                used_material_ids.add(mat_id.IntegerValue)
        
        # Check for compound structures (walls, floors, roofs, ceilings)
        if hasattr(elem_type, 'GetCompoundStructure'):
            try:
                compound = elem_type.GetCompoundStructure()
                if compound:
                    for layer in compound.GetLayers():
                        mat_id = layer.MaterialId
                        if mat_id != ElementId.InvalidElementId:
                            used_material_ids.add(mat_id.IntegerValue)
            except:
                pass
        
        # Check all parameters for ElementId that might reference materials
        for param in elem_type.Parameters:
            if param.StorageType == StorageType.ElementId and param.HasValue:
                elem_id = param.AsElementId()
                if elem_id != ElementId.InvalidElementId:
                    try:
                        elem = doc.GetElement(elem_id)
                        if isinstance(elem, Material):
                            used_material_ids.add(elem_id.IntegerValue)
                    except:
                        pass
    
    # Check all instances
    output.print_md("*Checking element instances...*")
    for elem in all_instances:
        # Check Material parameter
        mat_param = elem.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
        if mat_param and mat_param.HasValue:
            mat_id = mat_param.AsElementId()
            if mat_id != ElementId.InvalidElementId:
                used_material_ids.add(mat_id.IntegerValue)
        
        # Check Structural Material parameter
        struct_mat_param = elem.get_Parameter(BuiltInParameter.STRUCTURAL_MATERIAL_PARAM)
        if struct_mat_param and struct_mat_param.HasValue:
            mat_id = struct_mat_param.AsElementId()
            if mat_id != ElementId.InvalidElementId:
                used_material_ids.add(mat_id.IntegerValue)
        
        # Check for paint materials on faces
        try:
            if hasattr(elem, 'GetMaterialIds'):
                paint_mat_ids = elem.GetMaterialIds(True)  # True = get painted materials
                for mat_id in paint_mat_ids:
                    if mat_id != ElementId.InvalidElementId:
                        used_material_ids.add(mat_id.IntegerValue)
        except:
            pass
        
        # Check all parameters for ElementId that might reference materials
        for param in elem.Parameters:
            if param.StorageType == StorageType.ElementId and param.HasValue:
                elem_id = param.AsElementId()
                if elem_id != ElementId.InvalidElementId:
                    try:
                        mat_elem = doc.GetElement(elem_id)
                        if isinstance(mat_elem, Material):
                            used_material_ids.add(elem_id.IntegerValue)
                    except:
                        pass
        
        # Check geometry for materials (most comprehensive check)
        try:
            geom_elem = elem.get_Geometry(geo_options)
            if geom_elem:
                for geom_obj in geom_elem:
                    # Check geometry instance (for family instances)
                    if hasattr(geom_obj, 'GetInstanceGeometry'):
                        try:
                            inst_geom = geom_obj.GetInstanceGeometry()
                            if inst_geom:
                                for inst_obj in inst_geom:
                                    if isinstance(inst_obj, Solid):
                                        for face in inst_obj.Faces:
                                            mat_id = face.MaterialElementId
                                            if mat_id != ElementId.InvalidElementId:
                                                used_material_ids.add(mat_id.IntegerValue)
                        except:
                            pass
                    
                    # Check symbol geometry (for family instances)
                    if hasattr(geom_obj, 'GetSymbolGeometry'):
                        try:
                            sym_geom = geom_obj.GetSymbolGeometry()
                            if sym_geom:
                                for sym_obj in sym_geom:
                                    if isinstance(sym_obj, Solid):
                                        for face in sym_obj.Faces:
                                            mat_id = face.MaterialElementId
                                            if mat_id != ElementId.InvalidElementId:
                                                used_material_ids.add(mat_id.IntegerValue)
                        except:
                            pass
                    
                    # Check solid geometry directly
                    if isinstance(geom_obj, Solid):
                        for face in geom_obj.Faces:
                            mat_id = face.MaterialElementId
                            if mat_id != ElementId.InvalidElementId:
                                used_material_ids.add(mat_id.IntegerValue)
        except:
            pass
    
    # Categorize materials
    used_materials = []
    unused_materials = []
    
    for material in all_materials:
        if material.Id.IntegerValue in used_material_ids:
            used_materials.append(material)
        else:
            unused_materials.append(material)
    
    # Print Results
    output.print_md("# Materials Usage Report")
    output.print_md("---")
    
    # Summary
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
        used_materials_sorted = sorted(used_materials, key=lambda x: x.Name)
        for idx, mat in enumerate(used_materials_sorted, 1):
            output.print_md("{0}. {1} **{2}**".format(
                idx,
                output.linkify(mat.Id),
                mat.Name
            ))
        output.print_md("")
        output.print_md("---")
    
    # Unused Materials
    if unused_materials:
        output.print_md("## UNUSED MATERIALS ({0})".format(unused_count))
        output.print_md("")
        output.print_md("*These materials are not assigned to any elements and can potentially be deleted.*")
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
