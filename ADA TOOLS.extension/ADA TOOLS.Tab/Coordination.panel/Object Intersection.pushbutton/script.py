# -*- coding: utf-8 -*-
"""
Places a "VAL_M_Round Face Opening Solid_RVT2023" family 
at the base of the intersection between structural and MEP elements,
hosted on the structural element's face
"""
__title__ = 'Intersection\nOpening'
__author__ = 'Assistant'

import clr
import math
import System
from System.Collections.Generic import List

from pyrevit import forms, script
from pyrevit import revit, DB, UI

from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import (
    ElementId, Wall, Floor, FamilyInstance, BuiltInCategory, 
    Options, Solid, GeometryElement, BooleanOperationsUtils, 
    BooleanOperationsType, Transaction, Line, XYZ,
    FilteredElementCollector, FamilySymbol, Level, Face,
    CurveLoop, SolidOptions, GeometryCreationUtilities,
    BuiltInParameter, UnitUtils, Transform, SolidUtils,
    Reference, PlanarFace, UV, GeometryCreationUtilities, 
    BRepBuilderEdgeGeometry
)

# Check Revit API version and choose appropriate unit type
try:
    # For newer Revit versions (2022+)
    from Autodesk.Revit.DB import UnitTypeId
    length_unit = UnitTypeId.Millimeters
    use_unit_type_id = True
except ImportError:
    # For older Revit versions
    from Autodesk.Revit.DB import DisplayUnitType
    length_unit = DisplayUnitType.DUT_MILLIMETERS
    use_unit_type_id = False

# Set up the output
output = script.get_output()
output.close_others()

# Get document and UI document
doc = revit.doc
uidoc = revit.uidoc

# Create selection filter for structural elements
class StructuralElementsFilter(ISelectionFilter):
    def AllowElement(self, element):
        if isinstance(element, (Wall, Floor)):
            return True
        
        if isinstance(element, FamilyInstance):
            category = element.Category
            if category and (category.Id.IntegerValue == 
                            int(BuiltInCategory.OST_StructuralColumns) or 
                            category.Id.IntegerValue == 
                            int(BuiltInCategory.OST_StructuralFraming)):
                return True
        
        return False
    
    def AllowReference(self, reference, position):
        return False

# Create selection filter for MEP elements
class MEPElementsFilter(ISelectionFilter):
    def AllowElement(self, element):
        category = element.Category
        if not category:
            return False
            
        # OST_DuctCurves, OST_PipeCurves, OST_Conduit, OST_MechanicalEquipment, OST_ElectricalEquipment
        valid_category_ids = [
            int(BuiltInCategory.OST_DuctCurves),
            int(BuiltInCategory.OST_PipeCurves),
            int(BuiltInCategory.OST_Conduit),
            int(BuiltInCategory.OST_MechanicalEquipment),
            int(BuiltInCategory.OST_ElectricalEquipment),
            int(BuiltInCategory.OST_CableTray),
            int(BuiltInCategory.OST_FlexDuctCurves),
            int(BuiltInCategory.OST_FlexPipeCurves)
        ]
        
        return category.Id.IntegerValue in valid_category_ids
    
    def AllowReference(self, reference, position):
        return False

def get_element_geometry(element, options=None):
    """Gets the geometry of an element."""
    if not options:
        options = Options()
        options.ComputeReferences = True
        options.DetailLevel = DB.ViewDetailLevel.Fine
        
    geometry = element.get_Geometry(options)
    if not geometry:
        return None
        
    # Flatten the geometry ino a list of solids
    solids = []
    
    # First level of iteration gets GeometryObjects
    for geo_obj in geometry:
        # If we have an instance, we need to get its symbol geometry
        if isinstance(geo_obj, DB.GeometryInstance):
            # Get the instance geometry in model coordinates
            instance_geo = geo_obj.GetInstanceGeometry()
            for instance_obj in instance_geo:
                if isinstance(instance_obj, DB.Solid) and instance_obj.Volume > 0:
                    solids.append(instance_obj)
        # Direct solid
        elif isinstance(geo_obj, DB.Solid) and geo_obj.Volume > 0:
            solids.append(geo_obj)
            
    # If we found solids, return them
    if solids:
        return solids
    
    return None

def get_largest_solid(solids):
    """Returns the largest solid from a list of solids."""
    largest_volume = 0
    largest_solid = None
    
    for solid in solids:
        if solid.Volume > largest_volume:
            largest_volume = solid.Volume
            largest_solid = solid
            
    return largest_solid

def get_structural_face_at_intersection(struct_element, intersection_solid):
    """Find the face of the structural element that should host the opening"""
    if not struct_element or not intersection_solid:
        return None
    
    # Get the geometry of the structural element
    options = Options()
    options.ComputeReferences = True
    options.DetailLevel = DB.ViewDetailLevel.Fine
    
    struct_geom = struct_element.get_Geometry(options)
    if not struct_geom:
        return None
    
    # Get the bounding box of the intersection solid to find its center
    bbox = intersection_solid.GetBoundingBox()
    if not bbox:
        return None
    
    # Calculate the center of the intersection
    intersection_center = XYZ(
        (bbox.Min.X + bbox.Max.X) / 2,
        (bbox.Min.Y + bbox.Max.Y) / 2,
        (bbox.Min.Z + bbox.Max.Z) / 2
    )
    
    # Find the closest face on the structural element to the intersection center
    min_distance = float('inf')
    closest_face_ref = None
    
    for geom_obj in struct_geom:
        if isinstance(geom_obj, Solid):
            for face in geom_obj.Faces:
                # For each face, find the closest point to the intersection center
                try:
                    uv_point = face.Project(intersection_center).UVPoint
                    face_point = face.Evaluate(uv_point)
                    distance = face_point.DistanceTo(intersection_center)
                    
                    if distance < min_distance:
                        min_distance = distance
                        closest_face_ref = face.Reference
                except:
                    continue
    
    return closest_face_ref

def get_bottom_face_center(solid):
    """Get the center point of the bottom face of a solid"""
    if not solid:
        return None
    
    # Get the bounding box to determine Z min
    bbox = solid.GetBoundingBox()
    if not bbox:
        return None
    
    min_z = bbox.Min.Z
    bottom_face = None
    
    # Find all faces where Z is approximately equal to min_z
    tolerance = 0.001  # Small tolerance for floating point comparison
    for face in solid.Faces:
        bbox_face = face.GetBoundingBox()
        if abs(bbox_face.Min.Z - min_z) < tolerance:
            bottom_face = face
            break
    
    if not bottom_face:
        # Fallback: find the face with lowest Z coordinate
        lowest_z = float('inf')
        for face in solid.Faces:
            bbox_face = face.GetBoundingBox()
            if bbox_face.Min.Z < lowest_z:
                lowest_z = bbox_face.Min.Z
                bottom_face = face
    
    if bottom_face:
        # Get face centroid by computing bbox center
        face_bbox = bottom_face.GetBoundingBox()
        center_pt = XYZ(
            (face_bbox.Min.X + face_bbox.Max.X) / 2,
            (face_bbox.Min.Y + face_bbox.Max.Y) / 2,
            (face_bbox.Min.Z + face_bbox.Max.Z) / 2
        )
        return center_pt
    
    # If we couldn't find a suitable face, fall back to bbox center at min Z
    return XYZ(
        (bbox.Min.X + bbox.Max.X) / 2,
        (bbox.Min.Y + bbox.Max.Y) / 2,
        bbox.Min.Z
    )

def get_structural_face_at_intersection(struct_element, intersection_solid):
    """Find the face of the structural element that should host the opening"""
    if not struct_element or not intersection_solid:
        return None
    
    # Get the geometry of the structural element
    options = Options()
    options.ComputeReferences = True
    options.DetailLevel = DB.ViewDetailLevel.Fine
    
    struct_geom = struct_element.get_Geometry(options)
    if not struct_geom:
        return None
    
    # Get the bounding box of the intersection solid to find its center
    bbox = intersection_solid.GetBoundingBox()
    if not bbox:
        return None
    
    # Calculate the center of the intersection
    intersection_center = XYZ(
        (bbox.Min.X + bbox.Max.X) / 2,
        (bbox.Min.Y + bbox.Max.Y) / 2,
        (bbox.Min.Z + bbox.Max.Z) / 2
    )
    
    # Find the closest face on the structural element to the intersection center
    min_distance = float('inf')
    closest_face_ref = None
    
    for geom_obj in struct_geom:
        if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
            for face in geom_obj.Faces:
                # Try to get closest point on face to intersection center
                try:
                    # Get face bounding box center as approximation
                    face_bbox = face.GetBoundingBox()
                    face_center = XYZ(
                        (face_bbox.Min.X + face_bbox.Max.X) / 2,
                        (face_bbox.Min.Y + face_bbox.Max.Y) / 2,
                        (face_bbox.Min.Z + face_bbox.Max.Z) / 2
                    )
                    
                    # Calculate distance
                    distance = face_center.DistanceTo(intersection_center)
                    
                    if distance < min_distance:
                        min_distance = distance
                        closest_face_ref = face.Reference
                except Exception:
                    continue
    
    return closest_face_ref


def get_solid_dimensions(solid):
    """Get the width, depth, height of a solid's bounding box."""
    if not solid:
        return None, None, None
    
    try:
        bb = solid.GetBoundingBox()
        min_pt = bb.Min
        max_pt = bb.Max
        width = abs(max_pt.X - min_pt.X)
        depth = abs(max_pt.Y - min_pt.Y)
        height = abs(max_pt.Z - min_pt.Z)
        return width, depth, height
    except:
        return None, None, None

def create_face_based_family_instance(doc, family_symbol, face_reference, location, diameter=None):
    """Creates a face-based family instance at the specified location on a face."""
    try:
        with Transaction(doc, "Place Opening Family") as trans:
            trans.Start()
            
            # Make sure the family symbol is active
            if not family_symbol.IsActive:
                family_symbol.Activate()
                doc.Regenerate()
            
            # Find the parent element
            host_element = doc.GetElement(face_reference.ElementId)
            
            # Create the instance using face-based placement
            instance = doc.Create.NewFamilyInstance(
                face_reference,
                location,
                XYZ(0, 0, 1),  # Up direction
                family_symbol
            )
                
            # Set the diameter parameter if needed
            if diameter:
                try:
                    diameter_param = instance.LookupParameter("Diameter")
                    if diameter_param:
                        diameter_param.Set(diameter)
                except Exception as e:
                    output.print_md("\nWarning: Failed to set diameter parameter: " + str(e))
                
            trans.Commit()
            
        return instance.Id
    except Exception as e:
        output.print_md("\nError creating family instance: " + str(e))
        return None

def format_length(length_value):
    """Format a length value to reasonable precision."""
    if use_unit_type_id:
        formatted = UnitUtils.ConvertFromInternalUnits(length_value, length_unit)
    else:
        formatted = UnitUtils.ConvertFromInternalUnits(length_value, length_unit)
    return round(formatted, 1)

def get_element_name_id(element):
    """Returns a string with element name and ID."""
    if not element:
        return "None"
    
    try:
        element_id = element.Id.IntegerValue
        element_name = element.Name
        category_name = element.Category.Name
        return "{}: {} (ID: {})".format(category_name, element_name, element_id)
    except:
        try:
            return "Element (ID: {})".format(element.Id.IntegerValue)
        except:
            return "Unknown Element"


# Main script execution
try:
    output.print_md("# Place Opening at Intersection")
    
    # Select the structural element
    output.print_md("\n## Selecting Structural Element")
    output.print_md("\nPlease select a structural element (wall, floor, column, beam)...")
    
    try:
        structural_ref = uidoc.Selection.PickObject(
            ObjectType.Element, StructuralElementsFilter(), 
            "Select a structural element (wall, floor, column, beam)"
        )
        structural_element = doc.GetElement(structural_ref.ElementId)
    except Exception as e:
        if 'canceled' in str(e).lower():
            raise Exception("Selection canceled")
        else:
            raise e
    
    if structural_element:
        output.print_md("\n### Selected Structural Element")
        output.print_md("\n{}".format(get_element_name_id(structural_element)))
        
        # Select the MEP element
        output.print_md("\n## Selecting MEP Element")
        output.print_md("\nPlease select an MEP element (duct, pipe, conduit, equipment)...")
        
        try:
            mep_ref = uidoc.Selection.PickObject(
                ObjectType.Element, MEPElementsFilter(), 
                "Select an MEP element (duct, pipe, conduit, equipment)"
            )
            mep_element = doc.GetElement(mep_ref.ElementId)
        except Exception as e:
            if 'canceled' in str(e).lower():
                raise Exception("Selection canceled")
            else:
                raise e
        
        if mep_element:
            output.print_md("\n### Selected MEP Element")
            output.print_md("\n{}".format(get_element_name_id(mep_element)))
            
            # Get the geometry of both elements
            output.print_md("\n## Analyzing Intersection")
            
            struct_solids = get_element_geometry(structural_element)
            mep_solids = get_element_geometry(mep_element)
            
            if struct_solids and mep_solids:
                try:
                    # Get the largest solid for each element
                    struct_solid = get_largest_solid(struct_solids)
                    mep_solid = get_largest_solid(mep_solids)
                    
                    if len(struct_solids) > 1:
                        output.print_md("\nNote: Multiple solids found in structural element, using the largest one.")
                    
                    if len(mep_solids) > 1:
                        output.print_md("\nNote: Multiple solids found in MEP element, using the largest one.")
                    
                    # Try to compute the intersection
                    try:
                        # Perform Boolean intersection
                        result_solid = BooleanOperationsUtils.ExecuteBooleanOperation(
                            struct_solid, mep_solid, BooleanOperationsType.Intersect
                        )
                        
                        if result_solid and result_solid.Volume > 0:
                            # Format the volume for display
                            if use_unit_type_id:
                                volume_mm3 = UnitUtils.ConvertFromInternalUnits(
                                    result_solid.Volume, UnitTypeId.CubicMillimeters
                                )
                            else:
                                volume_mm3 = UnitUtils.ConvertFromInternalUnits(
                                    result_solid.Volume, DisplayUnitType.DUT_CUBIC_MILLIMETERS
                                )
                                
                            # Display intersection info
                            output.print_md("\nIntersection volume: {:,.0f} mm³".format(volume_mm3))
                            
                            # Find the bottom face center of the intersection
                            bottom_center = get_bottom_face_center(result_solid)
                            
                            if not bottom_center:
                                output.print_md("\nWarning: Could not determine the base of the intersection solid.")
                                output.print_md("Falling back to the bounding box center...")
                                
                                # Fallback to bounding box center
                                bbox = result_solid.GetBoundingBox()
                                bottom_center = XYZ(
                                    (bbox.Min.X + bbox.Max.X) / 2,
                                    (bbox.Min.Y + bbox.Max.Y) / 2,
                                    bbox.Min.Z
                                )
                            
                            # Find the appropriate face of the structural element for hosting
                            # Find the appropriate face of the structural element for hosting
                            face_ref = get_structural_face_at_intersection(structural_element, result_solid)

                            if not face_ref:
                                output.print_md("\nWarning: Could not find a suitable host face on the structural element.")
                                output.print_md("Will try to place directly using the structural element...")
                                
                                # Try to use a reference from the structural element
                                try:
                                    # Create a transaction to get a reference
                                    with Transaction(doc, "Get Reference") as t:
                                        t.Start()
                                        ref_list = []
                                        for face in structural_element.GetGeometryObjectFromReference(structural_ref).Faces:
                                            ref_list.append(face.Reference)
                                        if ref_list:
                                            face_ref = ref_list[0]
                                        t.RollBack()
                                except Exception as e:
                                    output.print_md("\nError getting face reference: " + str(e))
                                    
                                # If we still couldn't get a face reference, try to get the user to select one
                                if not face_ref:
                                    output.print_md("\nPlease manually select a face on the structural element to place the opening...")
                                    try:
                                        selected_face_ref = uidoc.Selection.PickObject(ObjectType.Face, "Select a face on the structural element")
                                        face_ref = selected_face_ref
                                    except Exception as e:
                                        if 'canceled' in str(e).lower():
                                            output.print_md("\nFace selection canceled by user.")
                                        else:
                                            output.print_md("\nError selecting face: " + str(e))
                            
                            # Find the "VAL_M_Round Face Opening Solid_RVT2023" family
                            output.print_md("\n## Placing Opening Family")
                            
                            family_symbols = FilteredElementCollector(doc).OfClass(FamilySymbol)
                            opening_symbol = None
                            
                            for symbol in family_symbols:
                                if "VAL_M_Round Face Opening Solid_RVT2023" in symbol.Family.Name:
                                    opening_symbol = symbol
                                    break
                            
                            if not opening_symbol:
                                output.print_md("\n**Error:** Could not find the 'VAL_M_Round Face Opening Solid_RVT2023' family. Please load it into the project.")
                            else:
                                # Get the dimensions of the intersection
                                width, depth, height = get_solid_dimensions(result_solid)
                                
                                # Calculate diameter from the intersection dimensions
                                # Use the maximum of width and depth to ensure the opening covers the MEP element
                                diameter = max(width, depth, height) * 1.1  # Add 10% safety margin
                                
                                # Place the family on the structural face
                                output.print_md("\nPlacing opening at: X={:.1f} mm, Y={:.1f} mm, Z={:.1f} mm".format(
                                    format_length(bottom_center.X),
                                    format_length(bottom_center.Y),
                                    format_length(bottom_center.Z)
                                ))
                                
                                # Place the family on the face at the intersection position
                                instance_id = create_face_based_family_instance(
                                    doc, opening_symbol, face_ref, bottom_center, diameter
                                )
                                
                                if instance_id:
                                    output.print_md("\n**Success!** Placed opening family on structural face.")
                                    
                                    # Try to zoom to the new family
                                    try:
                                        id_list = List[ElementId]()
                                        id_list.Add(instance_id)
                                        uidoc.Selection.SetElementIds(id_list)
                                        uidoc.ShowElements(instance_id)
                                    except Exception as e:
                                        output.print_md("Could not zoom to new element: " + str(e))
                                else:
                                    output.print_md("\n**Failed to place family.**")
                        else:
                            output.print_md("\n**No valid intersection** found between the elements.")
                            
                            # Use a face selection instead
                            output.print_md("\nNo intersection found. Please select a face on the structural element to place the opening...")
                            
                            try:
                                face_ref = uidoc.Selection.PickObject(
                                    ObjectType.Face, 
                                    "Select a face on the structural element to place the opening"
                                )
                                
                                # Find the location on the selected face
                                # Use the MEP element's location as basis
                                location = mep_element.Location
                                if location:
                                    location_point = None
                                    if hasattr(location, 'Point'):
                                        location_point = location.Point
                                    elif hasattr(location, 'Curve'):
                                        curve = location.Curve
                                        location_point = curve.Evaluate(0.5, True)  # Midpoint
                                    
                                    if location_point:
                                        # Find the MEP diameter
                                        mep_width, mep_depth, mep_height = get_solid_dimensions(mep_solid)
                                        diameter = max(mep_width, mep_depth, mep_height) * 1.1
                                        
                                        # Find the family
                                        family_symbols = FilteredElementCollector(doc).OfClass(FamilySymbol)
                                        opening_symbol = None
                                        
                                        for symbol in family_symbols:
                                            if "VAL_M_Round Face Opening Solid_RVT2023" in symbol.Family.Name:
                                                opening_symbol = symbol
                                                break
                                        
                                        if opening_symbol:
                                            # Place the family on the selected face
                                            instance_id = create_face_based_family_instance(
                                                doc, opening_symbol, face_ref, location_point, diameter
                                            )
                                            
                                            if instance_id:
                                                output.print_md("\n**Success!** Placed opening family on selected face.")
                                                
                                                # Try to zoom to the new family
                                                try:
                                                    id_list = List[ElementId]()
                                                    id_list.Add(instance_id)
                                                    uidoc.Selection.SetElementIds(id_list)
                                                    uidoc.ShowElements(instance_id)
                                                except Exception as e:
                                                    output.print_md("Could not zoom to new element: " + str(e))
                                            else:
                                                output.print_md("\n**Failed to place family.**")
                                        else:
                                            output.print_md("\n**Error:** Could not find the 'VAL_M_Round Face Opening Solid_RVT2023' family.")
                                    else:
                                        output.print_md("\n**Could not determine MEP element location.**")
                                else:
                                    output.print_md("\n**MEP element has no location.**")
                            except Exception as e:
                                if 'canceled' in str(e).lower():
                                    output.print_md("\n**Face selection canceled by user.**")
                                else:
                                    output.print_md("\n**Error selecting face:** " + str(e))
                                    
                    except Exception as e:
                        output.print_md("\n**Error performing boolean operation:** {}".format(str(e)))
                        output.print_md("\nSome geometry configurations can cause Boolean operations to fail. Simplifying the model, using different elements, or performing a sequence of Boolean operations in an order that avoids such conditions, may solve the problem.")
                
                except Exception as e:
                    output.print_md("\n**Error analyzing intersection:** {}".format(str(e)))
                    output.print_md("This is likely due to complex geometry or an invalid Boolean operation.")
            else:
                output.print_md("\n**Could not extract geometry from one or both elements.**")
        else:
            output.print_md("\n**Process canceled. No MEP element was selected.**")
    else:
        output.print_md("\n**Process canceled. No structural element was selected.**")
except Exception as e:
    if 'Selection canceled' in str(e):
        output.print_md("\n**Process canceled by user.**")
    else:
        output.print_md("\n**Error occurred:** {}".format(str(e)))

output.print_md("\n**Script completed.**")
