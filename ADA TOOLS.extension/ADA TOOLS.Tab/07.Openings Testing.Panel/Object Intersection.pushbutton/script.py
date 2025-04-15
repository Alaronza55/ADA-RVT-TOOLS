# -*- coding: utf-8 -*-
"""
Creates an in-place mass at the intersection volume between structural and MEP elements
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
                              CurveLoop, SolidOptions, GeometryCreationUtilities)
import System

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
        print("Error getting solid: " + str(e))
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
            ds.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS).Set(element_name)
            
            t.Commit()
            
            return ds.Id
            
    except Exception as e:
        print("Error creating in-place mass: " + str(e))
        import traceback
        print(traceback.format_exc())
        return None

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
    # Prompt user to select structural element
    output.print_md("Select a structural element (wall, floor, column, beam)...\n")
    selected_struct_ref = uidoc.Selection.PickObject(
        ObjectType.Element, structural_filter, "Select a structural element"
    )
    
    # Get the selected structural element
    struct_element = doc.GetElement(selected_struct_ref.ElementId)
    
    # Output selected structural element
    output.print_md("## Selected Structural Element\n")
    struct_info = get_element_info(struct_element)
    output.print_md(struct_info)
    
    try:
        # Prompt user to select MEP element
        output.print_md("\nSelect a MEP element (duct, pipe, conduit, equipment)...\n")
        selected_mep_ref = uidoc.Selection.PickObject(
            ObjectType.Element, mep_filter, "Select a MEP element"
        )
        
        # Get the selected MEP element
        mep_element = doc.GetElement(selected_mep_ref.ElementId)
        
        # Output selected MEP element
        output.print_md("\n## Selected MEP Element\n")
        mep_info = get_element_info(mep_element)
        output.print_md(mep_info)
        
        # Check for intersection
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
                    
                    mass_id = create_in_place_mass(doc, intersection_solid, struct_name, mep_name)
                    
                    if mass_id:
                        # Success message
                        output.print_md("**Success!** Created mass at intersection location.")
                        output.print_md("Intersection volume: {:.3f} cubic feet".format(
                            intersection_solid.Volume))
                        
                        # Zoom to the new mass
                        uidoc.Selection.SetElementIds([mass_id])
                        uidoc.ShowElements(mass_id)
                    else:
                        output.print_md("**Failed to create mass.**")
                else:
                    output.print_md("**No intersection found** between the selected elements.")
            except Exception as e:
                output.print_md("**Error creating intersection:** {}".format(str(e)))
                import traceback
                output.print_md("```\n{}\n```".format(traceback.format_exc()))
        else:
            output.print_md("**Could not extract geometry** from one or both elements.")
            
    except Exception as e:
        if 'Canceled' in str(e):
            output.print_md("\n**MEP selection canceled by user**")
        else:
            output.print_md("\n**Error during MEP selection:** {}".format(str(e)))
            
except Exception as e:
    if 'Canceled' in str(e):
        output.print_md("**Structural selection was canceled by user**")
    else:
        output.print_md("**Error during structural selection:** {}".format(str(e)))
