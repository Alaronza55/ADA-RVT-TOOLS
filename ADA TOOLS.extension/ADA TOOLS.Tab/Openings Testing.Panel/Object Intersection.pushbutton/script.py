# -*- coding: utf-8 -*-
"""
Creates an in-place mass at the intersection volume between structural and MEP elements
and prints detailed geometry information about the mass in millimeters
"""
__title__ = 'Intersection\nMass'
__author__ = 'Assistant'

import clr
from pyrevit import forms, script
from pyrevit import revit, DB, UI
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.DB import (ElementId, Wall, Floor, FamilyInstance, BuiltInCategory, 
                              Options, Solid, GeometryElement, BooleanOperationsUtils, 
                              BooleanOperationsType, Transaction, Line, XYZ,
                              CurveLoop, SolidOptions, GeometryCreationUtilities,
                              BuiltInParameter, UnitUtils)
import System
from System.Collections.Generic import List
import math

# Check Revit API version and choose appropriate unit type
try:
    # For newer Revit versions (2022+)
    from Autodesk.Revit.DB import UnitTypeId
    use_unit_type_id = True
except ImportError:
    # For older Revit versions
    from Autodesk.Revit.DB import DisplayUnitType
    use_unit_type_id = False

# Create selection filter for structural elements
class StructuralElementsFilter(ISelectionFilter):
    def AllowElement(self, element):
        # Allow Walls
        if isinstance(element, Wall):
            return True
        
        # Allow Floors
        if isinstance(element, Floor):
            return True
        
        # Check for Structural Framing and Structural Columns
        if isinstance(element, FamilyInstance):
            category_id = element.Category.Id
            if (category_id == ElementId(BuiltInCategory.OST_StructuralFraming) or
                category_id == ElementId(BuiltInCategory.OST_StructuralColumns)):
                return True
        
        return False

    def AllowReference(self, reference, position):
        return False

# Create selection filter for MEP elements
class MEPElementsFilter(ISelectionFilter):
    def AllowElement(self, element):
        # Initialize category_id
        category_id = None
        
        # Get category id if element has a category
        if element.Category:
            category_id = element.Category.Id
            
        # Check for various MEP categories
        mep_categories = [
            BuiltInCategory.OST_DuctCurves,
            BuiltInCategory.OST_DuctFitting,
            BuiltInCategory.OST_DuctTerminal, 
            BuiltInCategory.OST_PipeCurves,
            BuiltInCategory.OST_PipeFitting,
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_PlumbingFixtures,
            BuiltInCategory.OST_CableTray,
            BuiltInCategory.OST_Conduit,
            BuiltInCategory.OST_ElectricalEquipment,
            BuiltInCategory.OST_ElectricalFixtures
        ]
        
        if category_id:
            for category in mep_categories:
                if category_id == ElementId(category):
                    return True
                
        return False

    def AllowReference(self, reference, position):
        return False

# Helper function to get solid geometry
def get_element_solid(element, doc):
    try:
        # Create options for geometry extraction
        options = Options()
        options.DetailLevel = DB.ViewDetailLevel.Fine
        options.ComputeReferences = True
        options.IncludeNonVisibleObjects = True
        
        # Get geometry element
        geom_elem = element.get_Geometry(options)
        if geom_elem:
            solids = []
            
            # Iterate through geometry elements
            for geom in geom_elem:
                # Check if it's a solid
                if isinstance(geom, Solid) and geom.Volume > 0:
                    solids.append(geom)
                # Check if it's an instance
                elif isinstance(geom, DB.GeometryInstance):
                    inst_geom = geom.GetInstanceGeometry()
                    for inst_obj in inst_geom:
                        if isinstance(inst_obj, Solid) and inst_obj.Volume > 0:
                            solids.append(inst_obj)
            
            # Combine all solids if there are multiple
            if len(solids) > 1:
                combined = solids[0]
                for i in range(1, len(solids)):
                    combined = BooleanOperationsUtils.ExecuteBooleanOperation(
                        combined, solids[i], BooleanOperationsType.Union)
                return combined
            # Return the first solid found
            elif solids:
                return solids[0]
        
        return None
    
    except Exception as e:
        output.print_md("Note: Encountered issue getting solid: " + str(e))
        return None

# Function to get element info
def get_element_info(element):
    if element.Category:
        return "{}: {} (ID: {})".format(
            element.Category.Name,
            element.Name if hasattr(element, "Name") and element.Name else "",
            element.Id.IntegerValue
        )
    else:
        return "Element (ID: {})".format(element.Id.IntegerValue)

# Function to create an in-place mass from a solid
def create_in_place_mass(doc, solid, struct_name, mep_name):
    try:
        # Start a transaction
        with Transaction(doc, "Create Intersection Mass") as t:
            t.Start()
            
            # Create a Direct Shape element
            category_id = ElementId(BuiltInCategory.OST_Mass)
            
            # Create a DirectShape
            ds = DB.DirectShape.CreateElement(doc, category_id)
            
            # Set the shape
            ds.SetShape([solid])
            
            # Set name for the mass  
            element_name = "Intersection: {} - {}".format(struct_name, mep_name)
            ds.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(element_name)
            
            t.Commit()
            
            return ds.Id, solid
            
    except Exception as e:
        output.print_md("Error creating in-place mass: " + str(e))
        return None, None

# Function for unit conversion - handles both old and new API, converting to millimeters/cubic millimeters
def convert_to_mm(length):
    # First convert from internal units to feet
    if use_unit_type_id:
        feet = UnitUtils.ConvertFromInternalUnits(length, UnitTypeId.Feet)
    else:
        feet = UnitUtils.ConvertFromInternalUnits(length, DisplayUnitType.DUT_FEET)
    # Then convert from feet to millimeters (1 foot = 304.8 mm)
    return feet * 304.8

def convert_to_sq_mm(area):
    # First convert from internal units to square feet
    if use_unit_type_id:
        sq_feet = UnitUtils.ConvertFromInternalUnits(area, UnitTypeId.SquareFeet)
    else:
        sq_feet = UnitUtils.ConvertFromInternalUnits(area, DisplayUnitType.DUT_SQUARE_FEET)
    # Then convert from square feet to square millimeters (1 sq foot = 92903.04 sq mm)
    return sq_feet * 92903.04

def convert_to_cu_mm(volume):
    # First convert from internal units to cubic feet
    if use_unit_type_id:
        cu_feet = UnitUtils.ConvertFromInternalUnits(volume, UnitTypeId.CubicFeet)
    else:
        cu_feet = UnitUtils.ConvertFromInternalUnits(volume, DisplayUnitType.DUT_CUBIC_FEET)
    # Then convert from cubic feet to cubic millimeters (1 cu foot = 28316846.6 cu mm)
    return cu_feet * 28316846.6

# Function to print geometric details of a solid
def print_solid_geometry(solid, doc):
    if not solid:
        return
    
    def format_point(point):
        x = convert_to_mm(point.X)
        y = convert_to_mm(point.Y)
        z = convert_to_mm(point.Z)
        return "[{:.1f}, {:.1f}, {:.1f}]".format(x, y, z)
    
    # Print volume and surface area
    volume_cu_mm = convert_to_cu_mm(solid.Volume)
    surface_area_sq_mm = convert_to_sq_mm(solid.SurfaceArea)
    
    # Format large numbers with thousand separators
    volume_formatted = "{:,.0f}".format(volume_cu_mm)
    area_formatted = "{:,.0f}".format(surface_area_sq_mm)
    
    output.print_md("### Geometric Properties\n")
    output.print_md("- **Volume**: {} cubic millimeters (mm³)".format(volume_formatted))
    output.print_md("- **Surface Area**: {} square millimeters (mm²)".format(area_formatted))
    
    # Get bounding box
    bbox = solid.GetBoundingBox()
    if bbox:
        min_pt = bbox.Min
        max_pt = bbox.Max
        
        # Calculate dimensions in millimeters
        width_mm = convert_to_mm(max_pt.X - min_pt.X)
        depth_mm = convert_to_mm(max_pt.Y - min_pt.Y)
        height_mm = convert_to_mm(max_pt.Z - min_pt.Z)
        
        diagonal_mm = math.sqrt(width_mm**2 + depth_mm**2 + height_mm**2)
        
        output.print_md("\n### Bounding Box\n")
        output.print_md("- **Min Point (mm)**: {}".format(format_point(min_pt)))
        output.print_md("- **Max Point (mm)**: {}".format(format_point(max_pt)))
        output.print_md("- **Width × Depth × Height**: {:.1f} × {:.1f} × {:.1f} mm".format(
            width_mm, depth_mm, height_mm))
        output.print_md("- **Diagonal Length**: {:.1f} mm".format(diagonal_mm))
    
    # Count edges and faces
    edge_count = 0
    face_count = 0
    
    faces = solid.Faces
    edges = solid.Edges
    
    for _ in faces:
        face_count += 1
        
    for _ in edges:
        edge_count += 1
    
    output.print_md("\n### Topology\n")
    output.print_md("- **Face Count**: {}".format(face_count))
    output.print_md("- **Edge Count**: {}".format(edge_count))
    
    # Get centroid
    try:
        # Approximate centroid by averaging vertices
        vertex_points = []
        for edge in edges:
            curve = edge.AsCurve()
            if curve:
                vertex_points.append(curve.GetEndPoint(0))
                vertex_points.append(curve.GetEndPoint(1))
        
        if vertex_points:
            # Remove duplicate points
            unique_points = []
            for pt in vertex_points:
                is_duplicate = False
                for existing_pt in unique_points:
                    if pt.DistanceTo(existing_pt) < 0.01:  # threshold for considering points identical
                        is_duplicate = True
                        break
                if not is_duplicate:
                    unique_points.append(pt)
            
            # Calculate centroid
            if unique_points:
                sum_x = sum([pt.X for pt in unique_points])
                sum_y = sum([pt.Y for pt in unique_points])
                sum_z = sum([pt.Z for pt in unique_points])
                
                centroid = XYZ(sum_x / len(unique_points), 
                              sum_y / len(unique_points), 
                              sum_z / len(unique_points))
                
                output.print_md("\n### Location\n")
                output.print_md("- **Centroid (mm)**: {}".format(format_point(centroid)))
    except:
        pass

# Set up display output
output = script.get_output()
output.close_others(all_open_outputs=True)
output.print_md("# Create Mass at Intersection Volume\n")

# Get active document and UI document
doc = revit.doc
uidoc = revit.uidoc

# Create selection filters
structural_filter = StructuralElementsFilter()
mep_filter = MEPElementsFilter()

try:
    # PHASE 1: Select structural element
    struct_element = None
    if forms.alert("Step 1: Select a structural element (wall, floor, column, beam)", 
                yes=True, no=False, ok=False, cancel=False, 
                title="Intersection Mass Creator"):
        try:
            # Prompt user to select structural element
            output.print_md("## Selecting Structural Element\n")
            output.print_md("Please select a structural element (wall, floor, column, beam)...")
            
            selected_struct_ref = uidoc.Selection.PickObject(
                ObjectType.Element, structural_filter, "Select a structural element"
            )
            
            # Get the selected structural element
            struct_element = doc.GetElement(selected_struct_ref.ElementId)
            
            # Output selected structural element
            output.print_md("\n## Selected Structural Element\n")
            struct_info = get_element_info(struct_element)
            output.print_md(struct_info)
            
        except Exception as e:
            if 'Canceled' in str(e):
                output.print_md("**Structural selection was canceled by user**")
                raise Exception("Selection canceled")
            else:
                output.print_md("**Error during structural selection:** {}".format(str(e)))
                raise

    # PHASE 2: Select MEP element only if structural element was selected
    mep_element = None
    if struct_element and forms.alert("Step 2: Select an MEP element (duct, pipe, conduit, equipment)",
                                    yes=True, no=False, ok=False, cancel=False,
                                    title="Intersection Mass Creator"):
        try:
            # Prompt user to select MEP element
            output.print_md("\n## Selecting MEP Element\n")
            output.print_md("Please select an MEP element (duct, pipe, conduit, equipment)...")
            
            selected_mep_ref = uidoc.Selection.PickObject(
                ObjectType.Element, mep_filter, "Select a MEP element"
            )
            
            # Get the selected MEP element
            mep_element = doc.GetElement(selected_mep_ref.ElementId)
            
            # Output selected MEP element
            output.print_md("\n## Selected MEP Element\n")
            mep_info = get_element_info(mep_element)
            output.print_md(mep_info)
            
        except Exception as e:
            if 'Canceled' in str(e):
                output.print_md("\n**MEP selection canceled by user**")
                raise Exception("Selection canceled")
            else:
                output.print_md("\n**Error during MEP selection:** {}".format(str(e)))
                raise

    # PHASE 3: Create intersection mass if both elements were selected
    if struct_element and mep_element:
        output.print_md("\n## Creating Intersection Mass\n")
        
        # Get solids
        struct_solid = get_element_solid(struct_element, doc)
        mep_solid = get_element_solid(mep_element, doc)
        
        if struct_solid and mep_solid:
            try:
                # Calculate intersection solid
                intersection_solid = BooleanOperationsUtils.ExecuteBooleanOperation(
                    struct_solid, mep_solid, BooleanOperationsType.Intersect)
                
                if intersection_solid and intersection_solid.Volume > 0:
                    # Create the in-place mass
                    struct_name = struct_element.Category.Name
                    mep_name = mep_element.Category.Name
                    
                    mass_id, intersection_solid = create_in_place_mass(doc, intersection_solid, struct_name, mep_name)
                    
                    if mass_id:
                        # Success message
                        output.print_md("**Success!** Created mass at intersection location.")
                        
                        # Display volume in cubic millimeters
                        volume_cubic_mm = convert_to_cu_mm(intersection_solid.Volume)
                        volume_formatted = "{:,.0f}".format(volume_cubic_mm)
                        output.print_md("Intersection volume: {} mm³".format(volume_formatted))
                        
                        # Print detailed geometry information
                        output.print_md("\n## Intersection Geometry Details\n")
                        print_solid_geometry(intersection_solid, doc)
                        
                        # Try to zoom to the new mass
                        try:
                            # Create a .NET List of ElementId
                            id_list = List[ElementId]()
                            id_list.Add(mass_id)
                            
                            # Now pass the proper .NET collection to SetElementIds
                            uidoc.Selection.SetElementIds(id_list)
                            uidoc.ShowElements(mass_id)
                        except Exception as zoom_error:
                            # Non-critical - if the zoom fails, just continue
                            output.print_md("\n_Note: Could not zoom to mass. You'll need to find it manually._")
                    else:
                        output.print_md("**Failed to create mass.**")
                else:
                    output.print_md("**No intersection found** between the selected elements.")
            except Exception as e:
                output.print_md("**Error creating intersection:** {}".format(str(e)))
                output.print_md("This is likely due to complex geometry or an invalid Boolean operation.")
        else:
            output.print_md("**Could not extract geometry** from one or both elements.")
    else:
        if struct_element:
            output.print_md("\n**Process canceled. No MEP element was selected.**")
        elif mep_element:
            output.print_md("\n**Process canceled. No structural element was selected.**")
        else:
            output.print_md("\n**Process canceled. No elements were selected.**")

except Exception as e:
    if 'Selection canceled' in str(e):
        output.print_md("\n**Process canceled by user.**")
    else:
        output.print_md("\n**Error occurred:** {}".format(str(e)))

output.print_md("\n**Script completed.**")
