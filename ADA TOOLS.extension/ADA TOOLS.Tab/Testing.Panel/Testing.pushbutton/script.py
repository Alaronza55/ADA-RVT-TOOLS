# -*- coding: utf-8 -*-
"""Replace and Align Family Instances
Replaces all instances of family type 'BES_RESA_CIRC: BES_RESA CIRC HORIZONTAL' 
with 'BES_Opening_Horizontal-circular' and aligns their geometry centers in the active view.
"""

__title__ = 'Replace & Align\nCircular Openings'
__author__ = 'Your Name'

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType
from pyrevit import revit, DB

doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView


def get_geometry_center(element):
    """Get the geometry center of an element using its geometry."""
    try:
        options = Options()
        options.ComputeReferences = True
        options.IncludeNonVisibleObjects = True
        options.DetailLevel = ViewDetailLevel.Fine
        
        geom = element.get_Geometry(options)
        if geom:
            # Collect all solids
            solids = []
            
            for geom_obj in geom:
                if isinstance(geom_obj, Solid):
                    if geom_obj.Volume > 0.0001:  # Minimum volume threshold
                        solids.append(geom_obj)
                elif isinstance(geom_obj, GeometryInstance):
                    inst_geom = geom_obj.GetInstanceGeometry()
                    if inst_geom:
                        for inst_obj in inst_geom:
                            if isinstance(inst_obj, Solid) and inst_obj.Volume > 0.0001:
                                solids.append(inst_obj)
            
            # If we found solids, calculate weighted centroid
            if solids:
                total_volume = 0.0
                weighted_center = XYZ(0, 0, 0)
                
                for solid in solids:
                    volume = solid.Volume
                    centroid = solid.ComputeCentroid()
                    weighted_center = weighted_center + (centroid * volume)
                    total_volume += volume
                
                if total_volume > 0:
                    return weighted_center / total_volume
            
            # Fallback 1: Try bounding box center
            bbox = element.get_BoundingBox(None)
            if bbox:
                center = (bbox.Min + bbox.Max) / 2
                print("    (Using bounding box center)")
                return center
        
        # Fallback 2: Use location point
        location = element.Location
        if isinstance(location, LocationPoint):
            print("    (Using location point)")
            return location.Point
            
    except Exception as e:
        print("    Error getting geometry center: {}".format(str(e)))
    
    return None


def get_family_symbol_by_names(family_name, type_name):
    """Find a family symbol by family name and type name."""
    collector = FilteredElementCollector(doc)\
        .OfClass(FamilySymbol)\
        .WhereElementIsElementType()
    
    for symbol in collector:
        if (symbol.Family.Name == family_name and 
            symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == type_name):
            return symbol
    return None


def main():
    # Define source family and type names
    source_family_name = "BES_RESA_CIRC"
    source_type_name = "BES_RESA CIRC HORIZONTAL"
    
    # Define target family name (assuming any type from this family)
    target_family_name = "BES_Opening_Horizontal-circular"
    
    print("\n" + "="*60)
    print("Starting replacement process...")
    print("Source: {} : {}".format(source_family_name, source_type_name))
    print("Target family: {}".format(target_family_name))
    print("="*60 + "\n")
    
    # Get all family instances in active view
    collector = FilteredElementCollector(doc, active_view.Id)\
        .OfClass(FamilyInstance)\
        .WhereElementIsNotElementType()
    
    # Filter instances of source family and type
    source_instances = []
    
    for instance in collector:
        family_name = instance.Symbol.Family.Name
        type_name = instance.Symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        
        if family_name == source_family_name and type_name == source_type_name:
            source_instances.append(instance)
            print("Found instance: {} (Family: {}, Type: {})".format(
                instance.Id, family_name, type_name))
    
    print("\nTotal source instances found: {}".format(len(source_instances)))
    
    if not source_instances:
        print("\nNo instances of '{}:{}' found in active view.".format(
            source_family_name, source_type_name))
        print("\nListing all family instances in view for reference:")
        
        collector = FilteredElementCollector(doc, active_view.Id)\
            .OfClass(FamilyInstance)\
            .WhereElementIsNotElementType()
        
        families_found = {}
        for instance in collector:
            fam = instance.Symbol.Family.Name
            typ = instance.Symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            key = "{} : {}".format(fam, typ)
            families_found[key] = families_found.get(key, 0) + 1
        
        for key in sorted(families_found.keys()):
            print("  {} ({} instances)".format(key, families_found[key]))
        
        return
    
    # Find target family symbol (any type from the target family)
    target_symbol = None
    all_symbols = FilteredElementCollector(doc)\
        .OfClass(FamilySymbol)\
        .WhereElementIsElementType()
    
    print("\nSearching for target family '{}'...".format(target_family_name))
    
    for symbol in all_symbols:
        family_name = symbol.Family.Name
        if family_name == target_family_name:
            type_name = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            target_symbol = symbol
            print("Found target: {} : {}".format(family_name, type_name))
            break
    
    if not target_symbol:
        print("\nERROR: Target family '{}' not found!".format(target_family_name))
        print("\nListing all families in project containing 'Opening' or 'BES':")
        
        for symbol in all_symbols:
            fam = symbol.Family.Name
            if 'opening' in fam.lower() or 'bes' in fam.lower():
                typ = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                print("  {} : {}".format(fam, typ))
        return
    
    # Start transaction
    t = Transaction(doc, "Replace and Align Family Instances")
    t.Start()
    
    try:
        # Activate the symbol if it's not already
        if not target_symbol.IsActive:
            target_symbol.Activate()
            print("\nActivated target family symbol")
        
        replaced_count = 0
        failed_count = 0
        
        for i, instance in enumerate(source_instances):
            try:
                print("\n--- Processing instance {} of {} (ID: {}) ---".format(
                    i+1, len(source_instances), instance.Id))
                
                # Get the geometry center of the source instance
                old_center = get_geometry_center(instance)
                
                if old_center is None:
                    print("  ERROR: Could not get geometry center")
                    failed_count += 1
                    continue
                
                print("  Old center: ({:.3f}, {:.3f}, {:.3f})".format(
                    old_center.X, old_center.Y, old_center.Z))
                
                # Get instance properties
                old_location = instance.Location
                
                if not isinstance(old_location, LocationPoint):
                    print("  ERROR: Instance doesn't have LocationPoint")
                    failed_count += 1
                    continue
                
                old_point = old_location.Point
                old_rotation = old_location.Rotation
                
                print("  Old location: ({:.3f}, {:.3f}, {:.3f})".format(
                    old_point.X, old_point.Y, old_point.Z))
                print("  Old rotation: {:.3f} radians".format(old_rotation))
                
                # Get the host level
                level_id = instance.LevelId
                level = doc.GetElement(level_id) if level_id != ElementId.InvalidElementId else None
                
                print("  Level: {}".format(level.Name if level else "None"))
                
                # Create new instance at the same location
                if level:
                    new_instance = doc.Create.NewFamilyInstance(
                        old_point,
                        target_symbol,
                        level,
                        StructuralType.NonStructural
                    )
                else:
                    new_instance = doc.Create.NewFamilyInstance(
                        old_point,
                        target_symbol,
                        active_view,
                        StructuralType.NonStructural
                    )
                
                print("  New instance created (ID: {})".format(new_instance.Id))
                
                # Transfer parameter values FIRST (before calculating geometry)
                print("  Transferring parameters...")
                
                # Transfer BES_RESA Diameter -> OPE_DIAMETER
                old_diameter_param = instance.LookupParameter("BES_RESA Diameter")
                new_diameter_param = new_instance.LookupParameter("OPE_DIAMETER")
                
                if old_diameter_param and new_diameter_param:
                    if old_diameter_param.HasValue:
                        diameter_value = old_diameter_param.AsDouble()
                        new_diameter_param.Set(diameter_value)
                        print("    BES_RESA Diameter -> OPE_DIAMETER: {:.3f}".format(diameter_value))
                    else:
                        print("    BES_RESA Diameter: No value set")
                else:
                    if not old_diameter_param:
                        print("    WARNING: BES_RESA Diameter parameter not found on old instance")
                    if not new_diameter_param:
                        print("    WARNING: OPE_DIAMETER parameter not found on new instance")
                
                # Transfer BES_RESA Profondeur -> OPE_THICKNESS
                old_thickness_param = instance.LookupParameter("BES_RESA Profondeur")
                new_thickness_param = new_instance.LookupParameter("OPE_THICKNESS")
                
                if old_thickness_param and new_thickness_param:
                    if old_thickness_param.HasValue:
                        thickness_value = old_thickness_param.AsDouble()
                        new_thickness_param.Set(thickness_value)
                        print("    BES_RESA Profondeur -> OPE_THICKNESS: {:.3f}".format(thickness_value))
                    else:
                        print("    BES_RESA Profondeur: No value set")
                else:
                    if not old_thickness_param:
                        print("    WARNING: BES_RESA Profondeur parameter not found on old instance")
                    if not new_thickness_param:
                        print("    WARNING: OPE_THICKNESS parameter not found on new instance")
                
                # Transfer OPE_NUMBER (integer)
                old_number_param = instance.LookupParameter("OPE_NUMBER")
                new_number_param = new_instance.LookupParameter("OPE_NUMBER")
                
                if old_number_param and new_number_param:
                    if old_number_param.HasValue:
                        number_value = old_number_param.AsInteger()
                        new_number_param.Set(number_value)
                        print("    OPE_NUMBER: {}".format(number_value))
                    else:
                        print("    OPE_NUMBER: No value set")
                else:
                    if not old_number_param:
                        print("    WARNING: OPE_NUMBER parameter not found on old instance")
                    if not new_number_param:
                        print("    WARNING: OPE_NUMBER parameter not found on new instance")
                
                # Transfer OPE_DISCIPLINE (string)
                old_discipline_param = instance.LookupParameter("OPE_DISCIPLINE")
                new_discipline_param = new_instance.LookupParameter("OPE_DISCIPLINE")
                
                if old_discipline_param and new_discipline_param:
                    if old_discipline_param.HasValue:
                        discipline_value = old_discipline_param.AsString()
                        new_discipline_param.Set(discipline_value)
                        print("    OPE_DISCIPLINE: {}".format(discipline_value))
                    else:
                        print("    OPE_DISCIPLINE: No value set")
                else:
                    if not old_discipline_param:
                        print("    WARNING: OPE_DISCIPLINE parameter not found on old instance")
                    if not new_discipline_param:
                        print("    WARNING: OPE_DISCIPLINE parameter not found on new instance")
                
                # Transfer OPE_LEVEL (string)
                old_level_param = instance.LookupParameter("OPE_LEVEL")
                new_level_param = new_instance.LookupParameter("OPE_LEVEL")
                
                if old_level_param and new_level_param:
                    if old_level_param.HasValue:
                        level_value = old_level_param.AsString()
                        new_level_param.Set(level_value)
                        print("    OPE_LEVEL: {}".format(level_value))
                    else:
                        print("    OPE_LEVEL: No value set")
                else:
                    if not old_level_param:
                        print("    WARNING: OPE_LEVEL parameter not found on old instance")
                    if not new_level_param:
                        print("    WARNING: OPE_LEVEL parameter not found on new instance")
                
                # Transfer OPE_INDEC (string)
                old_indec_param = instance.LookupParameter("OPE_INDEC")
                new_indec_param = new_instance.LookupParameter("OPE_INDEC")
                
                if old_indec_param and new_indec_param:
                    if old_indec_param.HasValue:
                        indec_value = old_indec_param.AsString()
                        new_indec_param.Set(indec_value)
                        print("    OPE_INDEC: {}".format(indec_value))
                    else:
                        print("    OPE_INDEC: No value set")
                else:
                    if not old_indec_param:
                        print("    WARNING: OPE_INDEC parameter not found on old instance")
                    if not new_indec_param:
                        print("    WARNING: OPE_INDEC parameter not found on new instance")
                
                # Transfer OPE_DATE (string)
                old_date_param = instance.LookupParameter("OPE_DATE")
                new_date_param = new_instance.LookupParameter("OPE_DATE")
                
                if old_date_param and new_date_param:
                    if old_date_param.HasValue:
                        date_value = old_date_param.AsString()
                        new_date_param.Set(date_value)
                        print("    OPE_DATE: {}".format(date_value))
                    else:
                        print("    OPE_DATE: No value set")
                else:
                    if not old_date_param:
                        print("    WARNING: OPE_DATE parameter not found on old instance")
                    if not new_date_param:
                        print("    WARNING: OPE_DATE parameter not found on new instance")
                
                # Transfer OPE_ABSOLUTE LEVEL (could be string or double, we'll try both)
                old_abs_level_param = instance.LookupParameter("OPE_ABSOLUTE LEVEL")
                new_abs_level_param = new_instance.LookupParameter("OPE_ABSOLUTE LEVEL")
                
                if old_abs_level_param and new_abs_level_param:
                    if old_abs_level_param.HasValue:
                        # Try as double first (elevation value)
                        if old_abs_level_param.StorageType == StorageType.Double:
                            abs_level_value = old_abs_level_param.AsDouble()
                            new_abs_level_param.Set(abs_level_value)
                            print("    OPE_ABSOLUTE LEVEL: {:.3f}".format(abs_level_value))
                        # Try as string
                        elif old_abs_level_param.StorageType == StorageType.String:
                            abs_level_value = old_abs_level_param.AsString()
                            new_abs_level_param.Set(abs_level_value)
                            print("    OPE_ABSOLUTE LEVEL: {}".format(abs_level_value))
                        # Try as integer
                        elif old_abs_level_param.StorageType == StorageType.Integer:
                            abs_level_value = old_abs_level_param.AsInteger()
                            new_abs_level_param.Set(abs_level_value)
                            print("    OPE_ABSOLUTE LEVEL: {}".format(abs_level_value))
                    else:
                        print("    OPE_ABSOLUTE LEVEL: No value set")
                else:
                    if not old_abs_level_param:
                        print("    WARNING: OPE_ABSOLUTE LEVEL parameter not found on old instance")
                    if not new_abs_level_param:
                        print("    WARNING: OPE_ABSOLUTE LEVEL parameter not found on new instance")
                
                # Set rotation
                new_location = new_instance.Location
                if isinstance(new_location, LocationPoint) and old_rotation != 0:
                    axis = Line.CreateBound(old_point, old_point + XYZ.BasisZ)
                    new_location.Rotate(axis, old_rotation)
                    print("  Rotation applied")
                
                # Force multiple regenerations to ensure geometry is fully loaded with new parameters
                doc.Regenerate()
                doc.Regenerate()
                doc.Regenerate()
                
                # Get the geometry center of the new instance
                print("  Calculating new instance centroid...")
                new_center = get_geometry_center(new_instance)
                
                if new_center and old_center:
                    print("  New center: ({:.3f}, {:.3f}, {:.3f})".format(
                        new_center.X, new_center.Y, new_center.Z))
                    
                    # Calculate translation vector
                    translation = old_center - new_center
                    
                    # Check if translation is significant (more than 0.001 feet)
                    translation_length = translation.GetLength()
                    
                    print("  Translation vector: ({:.3f}, {:.3f}, {:.3f})".format(
                        translation.X, translation.Y, translation.Z))
                    print("  Translation distance: {:.3f}".format(translation_length))
                    
                    if translation_length > 0.001:
                        # Move the new instance to align centers
                        ElementTransformUtils.MoveElement(doc, new_instance.Id, translation)
                        print("  ✓ Instance moved to align centers")
                        
                        # Verify alignment after move
                        doc.Regenerate()
                        final_center = get_geometry_center(new_instance)
                        if final_center:
                            verification_distance = (old_center - final_center).GetLength()
                            print("  Verification: centers are now {:.4f} apart".format(verification_distance))
                    else:
                        print("  ✓ Centers already aligned (no move needed)")
                else:
                    if not new_center:
                        print("  WARNING: Could not get new center, skipping alignment")
                    if not old_center:
                        print("  WARNING: Could not get old center")
                
                # Delete the old instance
                doc.Delete(instance.Id)
                print("  ✓ Old instance deleted")
                
                replaced_count += 1
                
            except Exception as e:
                print("  ERROR: Failed - {}".format(str(e)))
                failed_count += 1
                continue
        
        t.Commit()
        
        # Show results
        print("\n" + "="*60)
        print("REPLACEMENT COMPLETED")
        print("="*60)
        print("Successfully replaced: {} instances".format(replaced_count))
        if failed_count > 0:
            print("Failed: {} instances".format(failed_count))
        print("="*60 + "\n")
        
    except Exception as e:
        t.RollBack()
        print("\nERROR: Transaction failed: {}".format(str(e)))
        import traceback
        traceback.print_exc()


if __name__ == '__main__':
    main()