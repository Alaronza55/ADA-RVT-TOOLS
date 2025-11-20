# -*- coding: utf-8 -*-
"""Select Element in Linked Model
This script allows you to select an element inside a linked Revit model.
"""

__title__ = 'Select Element\nin Link'
__author__ = 'Your Name'

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import revit, DB, UI, forms
import math
import traceback

# Get the current document and UI document
doc = revit.doc
uidoc = revit.uidoc


class LinkedElementSelectionFilter(ISelectionFilter):
    """Selection filter to allow only linked elements"""
    
    def AllowElement(self, element):
        """Allow selection of RevitLinkInstance"""
        if isinstance(element, RevitLinkInstance):
            return True
        return False
    
    def AllowReference(self, reference, point):
        """Allow selection of elements within links"""
        return True


class StructuralLinkedElementFilter(ISelectionFilter):
    """Selection filter to allow only structural elements in linked models"""
    
    def __init__(self, linked_doc):
        self.doc = linked_doc
    
    def AllowElement(self, element):
        """Allow selection of RevitLinkInstance"""
        if isinstance(element, RevitLinkInstance):
            return True
        return False
    
    def AllowReference(self, reference, point):
        """Check if the referenced element is a structural element"""
        try:
            # Get the link instance
            link_instance = self.doc.GetElement(reference.ElementId)
            if isinstance(link_instance, RevitLinkInstance):
                link_doc = link_instance.GetLinkDocument()
                if link_doc:
                    # Get the element in the link
                    element = link_doc.GetElement(reference.LinkedElementId)
                    if element and element.Category:
                        # Check if it's a structural category
                        structural_categories = [
                            'Structural Framing',
                            'Structural Columns',
                            'Structural Foundations',
                            'Floors',
                            'Walls',
                            'Structural Beam Systems',
                            'Structural Trusses',
                            'Structural Stiffeners',
                            'Structural Connections'
                        ]
                        if element.Category.Name in structural_categories:
                            return True
            return False
        except:
            return False


class MEPLinkedElementFilter(ISelectionFilter):
    """Selection filter to allow only MEP elements in linked models"""
    
    def __init__(self, linked_doc):
        self.doc = linked_doc
    
    def AllowElement(self, element):
        """Allow selection of RevitLinkInstance"""
        if isinstance(element, RevitLinkInstance):
            return True
        return False
    
    def AllowReference(self, reference, point):
        """Check if the referenced element is an MEP element"""
        try:
            # Get the link instance
            link_instance = self.doc.GetElement(reference.ElementId)
            if isinstance(link_instance, RevitLinkInstance):
                link_doc = link_instance.GetLinkDocument()
                if link_doc:
                    # Get the element in the link
                    element = link_doc.GetElement(reference.LinkedElementId)
                    if element and element.Category:
                        # Check if it's an MEP category
                        mep_categories = [
                            'Ducts',
                            'Duct Fittings',
                            'Duct Accessories',
                            'Flex Ducts',
                            'Pipes',
                            'Pipe Fittings',
                            'Pipe Accessories',
                            'Flex Pipes',
                            'Mechanical Equipment',
                            'Electrical Equipment',
                            'Electrical Fixtures',
                            'Conduits',
                            'Conduit Fittings',
                            'Cable Trays',
                            'Cable Tray Fittings',
                            'Lighting Fixtures',
                            'Air Terminals',
                            'Sprinklers',
                            'Plumbing Fixtures'
                        ]
                        if element.Category.Name in mep_categories:
                            return True
            return False
        except:
            return False


def get_element_from_link(reference):
    """Extract the linked element from a reference"""
    linked_elem = doc.GetElement(reference.ElementId)
    
    if isinstance(linked_elem, RevitLinkInstance):
        link_doc = linked_elem.GetLinkDocument()
        if link_doc:
            linked_elem_id = reference.LinkedElementId
            element_in_link = link_doc.GetElement(linked_elem_id)
            return element_in_link, link_doc, linked_elem
    
    return None, None, None


def get_mep_insulation_thickness(mep_element):
    """Get the insulation thickness of an MEP element"""
    try:
        insulation_thickness = 0.0
        
        # Method 1: Try to get InsulationThickness property
        if hasattr(mep_element, 'InsulationThickness'):
            try:
                thickness = mep_element.InsulationThickness
                if thickness and thickness > 0:
                    insulation_thickness = thickness
                    return insulation_thickness
            except:
                pass
        
        # Method 2: Try RBS_INSULATION_THICKNESS parameter
        if hasattr(mep_element, 'get_Parameter'):
            try:
                insulation_param = mep_element.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS)
                if insulation_param and insulation_param.AsDouble() > 0:
                    insulation_thickness = insulation_param.AsDouble()
                    return insulation_thickness
            except:
                pass
        
        # Method 3: Look through all parameters for "insulation" in name
        if hasattr(mep_element, 'Parameters'):
            try:
                for param in mep_element.Parameters:
                    param_name = param.Definition.Name.lower()
                    if 'insulation' in param_name and 'thickness' in param_name:
                        if param.StorageType == StorageType.Double:
                            thickness = param.AsDouble()
                            if thickness and thickness > 0:
                                insulation_thickness = thickness
                                return insulation_thickness
            except:
                pass
        
        # Method 4: Check if MEP element has insulation type set
        if hasattr(mep_element, 'InsulationType'):
            try:
                insulation_type = mep_element.InsulationType
                if insulation_type and insulation_type.Id != ElementId.InvalidElementId:
                    # If there's an insulation type, try to get thickness from type
                    thickness_param = insulation_type.get_Parameter(BuiltInParameter.RBS_INSULATION_THICKNESS)
                    if thickness_param and thickness_param.AsDouble() > 0:
                        insulation_thickness = thickness_param.AsDouble()
                        return insulation_thickness
            except:
                pass
        
        return insulation_thickness
        
    except Exception as e:
        return 0.0


def get_mep_diameter(mep_element):
    """Get the diameter of an MEP element"""
    try:
        diameter = 0.0
        
        # Try different diameter parameters
        # For ducts: Width or Height (use the smaller one for round ducts, or width for rectangular)
        # For pipes and conduits: Diameter parameter
        
        if mep_element.Category:
            cat_name = mep_element.Category.Name
            
            # For Pipes, Conduits, and Flex Pipes - look for Diameter
            if cat_name in ['Pipes', 'Pipe Fittings', 'Pipe Accessories', 'Flex Pipes', 
                           'Conduits', 'Conduit Fittings']:
                # Try Diameter parameter
                diameter_param = mep_element.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM)
                if not diameter_param:
                    diameter_param = mep_element.get_Parameter(BuiltInParameter.RBS_CONDUIT_DIAMETER_PARAM)
                
                if diameter_param:
                    diameter = diameter_param.AsDouble()
            
            # For Ducts - look for Width/Height
            elif cat_name in ['Ducts', 'Duct Fittings', 'Duct Accessories', 'Flex Ducts']:
                width_param = mep_element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
                height_param = mep_element.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
                
                if width_param and height_param:
                    width = width_param.AsDouble()
                    height = height_param.AsDouble()
                    # For round ducts, width and height should be equal
                    # For rectangular, use the larger dimension as diameter
                    diameter = max(width, height)
                elif width_param:
                    diameter = width_param.AsDouble()
                elif height_param:
                    diameter = height_param.AsDouble()
            
            # For Cable Trays
            elif cat_name in ['Cable Trays', 'Cable Tray Fittings']:
                width_param = mep_element.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
                if width_param:
                    diameter = width_param.AsDouble()
        
        # Return raw diameter (no margin added)
        return diameter
        
    except Exception as e:
        forms.alert(
            'Error getting MEP diameter:\n{}\nUsing default value.'.format(str(e)),
            title='Warning',
            warn_icon=True
        )
        return 0.0


def get_mep_diameter_with_insulation(mep_element):
    """Get the diameter of an MEP element including insulation thickness and 30mm offset"""
    try:
        # Get base diameter
        diameter = get_mep_diameter(mep_element)
        
        # Get insulation thickness
        insulation_thickness = get_mep_insulation_thickness(mep_element)
        
        # Add insulation thickness * 2 (thickness on both sides) + 30mm offset
        total_diameter = diameter + (insulation_thickness * 2.0) + (30.0 / 304.8)
        
        return total_diameter
        
    except Exception as e:
        return get_mep_diameter(mep_element)


def get_structural_depth(struct_element):
    """Get the depth/width/thickness of a structural element"""
    try:
        depth = 0.0
        found_param_name = "Unknown"
        cat_name = "Unknown"
        
        if struct_element.Category:
            cat_name = struct_element.Category.Name
            
            # For Structural Framing (beams)
            if cat_name == 'Structural Framing':
                # Try to get the depth/height parameter
                try:
                    depth_param = struct_element.get_Parameter(BuiltInParameter.STRUCTURAL_SECTION_COMMON_DEPTH)
                    if depth_param and depth_param.AsDouble() > 0:
                        depth = depth_param.AsDouble()
                        found_param_name = "STRUCTURAL_SECTION_COMMON_DEPTH (instance)"
                except:
                    pass
                
                # Try to get from type
                if depth == 0:
                    try:
                        if hasattr(struct_element, 'Symbol') and struct_element.Symbol:
                            depth_param = struct_element.Symbol.get_Parameter(BuiltInParameter.STRUCTURAL_SECTION_COMMON_DEPTH)
                            if depth_param and depth_param.AsDouble() > 0:
                                depth = depth_param.AsDouble()
                                found_param_name = "STRUCTURAL_SECTION_COMMON_DEPTH (type)"
                    except:
                        pass
                
                # If still not found, try height
                if depth == 0:
                    try:
                        if hasattr(struct_element, 'Symbol') and struct_element.Symbol:
                            depth_param = struct_element.Symbol.get_Parameter(BuiltInParameter.STRUCTURAL_SECTION_COMMON_HEIGHT)
                            if depth_param and depth_param.AsDouble() > 0:
                                depth = depth_param.AsDouble()
                                found_param_name = "STRUCTURAL_SECTION_COMMON_HEIGHT"
                    except:
                        pass
                
                # If still not found, search for parameters by name (b, h, width, depth)
                if depth == 0:
                    try:
                        # Try 'b' parameter on instance (width)
                        b_param = struct_element.LookupParameter("b")
                        if b_param and b_param.AsDouble() > 0:
                            depth = b_param.AsDouble()
                            found_param_name = "b (width parameter - instance)"
                        else:
                            # Try on type
                            if hasattr(struct_element, 'Symbol') and struct_element.Symbol:
                                b_param = struct_element.Symbol.LookupParameter("b")
                                if b_param and b_param.AsDouble() > 0:
                                    depth = b_param.AsDouble()
                                    found_param_name = "b (width parameter - type)"
                    except:
                        pass
                
                if depth == 0:
                    try:
                        # Try 'h' parameter on instance (height)
                        h_param = struct_element.LookupParameter("h")
                        if h_param and h_param.AsDouble() > 0:
                            depth = h_param.AsDouble()
                            found_param_name = "h (height parameter - instance)"
                        else:
                            # Try on type
                            if hasattr(struct_element, 'Symbol') and struct_element.Symbol:
                                h_param = struct_element.Symbol.LookupParameter("h")
                                if h_param and h_param.AsDouble() > 0:
                                    depth = h_param.AsDouble()
                                    found_param_name = "h (height parameter - type)"
                    except:
                        pass
                
                # If still zero, search ALL parameters for b or h
                if depth == 0:
                    try:
                        for param in struct_element.Parameters:
                            param_name = param.Definition.Name
                            if param_name == "b" or param_name == "h":
                                if param.StorageType == StorageType.Double and param.AsDouble() > 0:
                                    depth = param.AsDouble()
                                    found_param_name = "{} (found in all parameters - instance)".format(param_name)
                                    break
                    except:
                        pass
                
                # Check type parameters if still zero
                if depth == 0:
                    try:
                        if hasattr(struct_element, 'Symbol') and struct_element.Symbol:
                            for param in struct_element.Symbol.Parameters:
                                param_name = param.Definition.Name
                                if param_name == "b" or param_name == "h":
                                    if param.StorageType == StorageType.Double and param.AsDouble() > 0:
                                        depth = param.AsDouble()
                                        found_param_name = "{} (found in all parameters - type)".format(param_name)
                                        break
                    except:
                        pass
            
            # For Structural Columns
            elif cat_name == 'Structural Columns':
                # Try to get the width parameter (b dimension)
                width_param = struct_element.get_Parameter(BuiltInParameter.STRUCTURAL_SECTION_COMMON_WIDTH)
                if width_param and width_param.AsDouble() > 0:
                    depth = width_param.AsDouble()
                    found_param_name = "STRUCTURAL_SECTION_COMMON_WIDTH (instance)"
                else:
                    if hasattr(struct_element, 'Symbol') and struct_element.Symbol:
                        width_param = struct_element.Symbol.get_Parameter(BuiltInParameter.STRUCTURAL_SECTION_COMMON_WIDTH)
                        if width_param and width_param.AsDouble() > 0:
                            depth = width_param.AsDouble()
                            found_param_name = "STRUCTURAL_SECTION_COMMON_WIDTH (type)"
                
                # Try depth if width is 0
                if depth == 0:
                    if hasattr(struct_element, 'Symbol') and struct_element.Symbol:
                        depth_param = struct_element.Symbol.get_Parameter(BuiltInParameter.STRUCTURAL_SECTION_COMMON_DEPTH)
                        if depth_param and depth_param.AsDouble() > 0:
                            depth = depth_param.AsDouble()
                            found_param_name = "STRUCTURAL_SECTION_COMMON_DEPTH (type)"
            
            # For Structural Foundations
            elif cat_name == 'Structural Foundations':
                # Try multiple different parameters that foundations might use
                
                # Try STRUCTURAL_FOUNDATION_THICKNESS first
                thickness_param = struct_element.get_Parameter(BuiltInParameter.STRUCTURAL_FOUNDATION_THICKNESS)
                if thickness_param and thickness_param.AsDouble() > 0:
                    depth = thickness_param.AsDouble()
                    found_param_name = "STRUCTURAL_FOUNDATION_THICKNESS (instance)"
                
                # Try from type if not found on instance
                if depth == 0 and hasattr(struct_element, 'Symbol') and struct_element.Symbol:
                    thickness_param = struct_element.Symbol.get_Parameter(BuiltInParameter.STRUCTURAL_FOUNDATION_THICKNESS)
                    if thickness_param and thickness_param.AsDouble() > 0:
                        depth = thickness_param.AsDouble()
                        found_param_name = "STRUCTURAL_FOUNDATION_THICKNESS (type)"
                
                # Try FLOOR_ATTR_THICKNESS_PARAM (foundations might behave like floors)
                if depth == 0:
                    thickness_param = struct_element.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)
                    if thickness_param and thickness_param.AsDouble() > 0:
                        depth = thickness_param.AsDouble()
                        found_param_name = "FLOOR_ATTR_THICKNESS_PARAM (instance)"
                
                # Try getting from FoundationType
                if depth == 0 and hasattr(struct_element, 'FoundationType') and struct_element.FoundationType:
                    thickness_param = struct_element.FoundationType.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)
                    if thickness_param and thickness_param.AsDouble() > 0:
                        depth = thickness_param.AsDouble()
                        found_param_name = "FLOOR_ATTR_THICKNESS_PARAM (FoundationType)"
                
                # Try width/depth as last resort
                if depth == 0:
                    width_param = struct_element.get_Parameter(BuiltInParameter.STRUCTURAL_SECTION_COMMON_WIDTH)
                    if width_param and width_param.AsDouble() > 0:
                        depth = width_param.AsDouble()
                        found_param_name = "STRUCTURAL_SECTION_COMMON_WIDTH (instance)"
                    elif hasattr(struct_element, 'Symbol') and struct_element.Symbol:
                        width_param = struct_element.Symbol.get_Parameter(BuiltInParameter.STRUCTURAL_SECTION_COMMON_WIDTH)
                        if width_param and width_param.AsDouble() > 0:
                            depth = width_param.AsDouble()
                            found_param_name = "STRUCTURAL_SECTION_COMMON_WIDTH (type)"
                
                # If still nothing, try looking at ALL parameters by name
                if depth == 0:
                    for param in struct_element.Parameters:
                        param_name = param.Definition.Name.lower()
                        if 'thickness' in param_name or 'depth' in param_name or 'height' in param_name:
                            if param.AsDouble() and param.AsDouble() > 0:
                                depth = param.AsDouble()
                                found_param_name = "Found by name search: {}".format(param.Definition.Name)
                                break
            
            # For Floors and Walls
            elif cat_name in ['Floors', 'Walls']:
                # Get the width/thickness
                if cat_name == 'Walls':
                    width_param = struct_element.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
                    if width_param and width_param.AsDouble() > 0:
                        depth = width_param.AsDouble()
                        found_param_name = "WALL_ATTR_WIDTH_PARAM (instance)"
                    elif hasattr(struct_element, 'WallType') and struct_element.WallType:
                        width_param = struct_element.WallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM)
                        if width_param and width_param.AsDouble() > 0:
                            depth = width_param.AsDouble()
                            found_param_name = "WALL_ATTR_WIDTH_PARAM (type)"
                else:  # Floors
                    width_param = struct_element.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)
                    if width_param and width_param.AsDouble() > 0:
                        depth = width_param.AsDouble()
                        found_param_name = "FLOOR_ATTR_THICKNESS_PARAM (instance)"
                    elif hasattr(struct_element, 'FloorType') and struct_element.FloorType:
                        width_param = struct_element.FloorType.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM)
                        if width_param and width_param.AsDouble() > 0:
                            depth = width_param.AsDouble()
                            found_param_name = "FLOOR_ATTR_THICKNESS_PARAM (type)"
        
        # Show debug info
        debug_info = 'Structural Depth Calculation:\n'
        debug_info += 'Category: {}\n'.format(cat_name if struct_element.Category else 'None')
        debug_info += 'Parameter Found: {}\n'.format(found_param_name)
        debug_info += 'Depth Value: {:.4f} ft ({:.2f} mm)'.format(depth, depth * 304.8)
        
        if depth == 0:
            forms.alert(
                'WARNING: Structural depth is 0!\n\n' + debug_info + 
                '\n\nPlease verify the element has dimensional parameters.',
                title='Structural Depth Warning',
                warn_icon=True
            )
        else:
            # Show success info for debugging
            forms.alert(
                debug_info,
                title='Structural Depth Found',
                warn_icon=False
            )
        
        # Return the raw depth (no margin added)
        return depth
        
    except Exception as e:
        forms.alert(
            'Error getting structural depth:\n{}\nUsing default value.'.format(str(e)),
            title='Warning',
            warn_icon=True
        )
        return 0.0


def get_mep_direction(mep_element):
    """Get the direction vector of an MEP element"""
    try:
        # For linear MEP elements (pipes, ducts, conduits), get the curve direction
        if hasattr(mep_element, 'Location'):
            location = mep_element.Location
            if location and isinstance(location, LocationCurve):
                try:
                    curve = location.Curve
                    
                    # Get the direction vector at the start of the curve
                    if hasattr(curve, 'Direction'):
                        return curve.Direction
                    else:
                        # For curves that don't have a Direction property, calculate it
                        start_point = curve.GetEndPoint(0)
                        end_point = curve.GetEndPoint(1)
                        
                        # Create direction vector
                        direction = XYZ(
                            end_point.X - start_point.X,
                            end_point.Y - start_point.Y,
                            end_point.Z - start_point.Z
                        )
                        
                        # Normalize the vector
                        length = (direction.X**2 + direction.Y**2 + direction.Z**2)**0.5
                        if length > 0:
                            return XYZ(direction.X / length, direction.Y / length, direction.Z / length)
                except:
                    pass
        
        # For fittings, accessories, and other elements without LocationCurve
        # Try to get connectors and use the primary connector direction
        if hasattr(mep_element, 'MEPModel') and mep_element.MEPModel:
            try:
                mep_model = mep_element.MEPModel
                if hasattr(mep_model, 'ConnectorManager') and mep_model.ConnectorManager:
                    conn_mgr = mep_model.ConnectorManager
                    if hasattr(conn_mgr, 'Connectors') and conn_mgr.Connectors:
                        connectors = conn_mgr.Connectors
                        # Get the first connector's direction
                        for connector in connectors:
                            try:
                                if hasattr(connector, 'CoordinateSystem') and connector.CoordinateSystem:
                                    # Use the Z-axis of the connector (flow direction)
                                    return connector.CoordinateSystem.BasisZ
                            except:
                                continue
            except:
                pass
        
        # Default to X-axis if can't determine
        return XYZ(1, 0, 0)
        
    except Exception as e:
        # Default to X-axis on any error
        return XYZ(1, 0, 0)


def is_horizontal_element(struct_element):
    """Check if structural element is horizontal or vertical based on its orientation"""
    try:
        if struct_element.Category:
            cat_name = struct_element.Category.Name
            
            # For Structural Framing (beams) - check the orientation
            if cat_name == 'Structural Framing':
                # Try to get orientation from transform
                try:
                    if hasattr(struct_element, 'Location') and isinstance(struct_element.Location, LocationCurve):
                        curve = struct_element.Location.Curve
                        
                        # Get the transform to check orientation
                        if hasattr(struct_element.Location, 'Curve'):
                            # Try to get the hand orientation (cross product direction)
                            try:
                                transform = struct_element.GetTransform()
                                if transform:
                                    # The BasisZ of the transform indicates the "up" direction of the element
                                    basis_z = transform.BasisZ
                                    
                                    # If Z component is large, beam's "up" is vertical, so beam is VERTICAL
                                    # If Z component is small, beam's "up" is horizontal, so beam is HORIZONTAL
                                    z_component = abs(basis_z.Z)
                                    
                                    # If Z component is large, beam's "up" is vertical, so beam is VERTICAL
                                    # If Z component is small, beam's "up" is horizontal, so beam is HORIZONTAL
                                    if z_component > 0.7:
                                        return False  # Vertical (BasisZ points up)
                                    elif z_component < 0.3:
                                        return True  # Horizontal (BasisZ points sideways)
                                    else:
                                        # Diagonal - check which is dominant
                                        return z_component < 0.5  # If Z < 0.5, more horizontal
                            except:
                                pass
                except:
                    pass
            
            # For Structural Columns - always vertical
            elif cat_name == 'Structural Columns':
                return False  # Vertical
            
            # For Floors - always horizontal
            elif cat_name == 'Floors':
                return True  # Horizontal
            
            # For Walls - check if wall is vertical or horizontal
            elif cat_name == 'Walls':
                # Most walls are vertical, but we can check
                # For now, assume walls are vertical
                return False  # Vertical
        
        # Default to horizontal if can't determine
        return True
        
    except Exception as e:
        # Default to horizontal on error
        return True


def is_rectangular_mep(mep_element):
    """Check if MEP element has rectangular geometry (rectangular ducts)"""
    try:
        if mep_element.Category:
            cat_name = mep_element.Category.Name
            
            # For ducts, check if rectangular (width != height)
            if cat_name in ['Ducts', 'Duct Fittings', 'Duct Accessories', 'Flex Ducts']:
                width_param = mep_element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
                height_param = mep_element.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
                
                if width_param and height_param:
                    width = width_param.AsDouble()
                    height = height_param.AsDouble()
                    # Check if dimensions are different (within tolerance) - indicates rectangular duct
                    tolerance = 0.01  # 0.01 feet tolerance
                    if abs(width - height) >= tolerance:
                        return True
        
        return False
        
    except Exception as e:
        return False


def get_mep_width(mep_element):
    """Get the width of an MEP element"""
    try:
        width = 0.0
        
        if mep_element.Category:
            cat_name = mep_element.Category.Name
            
            # For Ducts - look for Width
            if cat_name in ['Ducts', 'Duct Fittings', 'Duct Accessories', 'Flex Ducts']:
                width_param = mep_element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
                if width_param:
                    width = width_param.AsDouble()
            
            # For Cable Trays
            elif cat_name in ['Cable Trays', 'Cable Tray Fittings']:
                width_param = mep_element.get_Parameter(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
                if width_param:
                    width = width_param.AsDouble()
        
        return width
        
    except Exception as e:
        return 0.0


def get_mep_height(mep_element):
    """Get the height of an MEP element"""
    try:
        height = 0.0
        
        if mep_element.Category:
            cat_name = mep_element.Category.Name
            
            # For Ducts - look for Height
            if cat_name in ['Ducts', 'Duct Fittings', 'Duct Accessories', 'Flex Ducts']:
                height_param = mep_element.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
                if height_param:
                    height = height_param.AsDouble()
            
            # For Cable Trays - use height parameter
            elif cat_name in ['Cable Trays', 'Cable Tray Fittings']:
                height_param = mep_element.get_Parameter(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
                if height_param:
                    height = height_param.AsDouble()
        
        return height
        
    except Exception as e:
        return 0.0


def is_cylindrical_mep(mep_element):
    """Check if MEP element has cylindrical geometry (pipes, round ducts, conduits)"""
    try:
        if mep_element.Category:
            cat_name = mep_element.Category.Name
            
            # Pipes and conduits are always cylindrical
            if cat_name in ['Pipes', 'Pipe Fittings', 'Pipe Accessories', 'Flex Pipes', 
                           'Conduits', 'Conduit Fittings']:
                return True
            
            # For ducts, check if round (width == height)
            elif cat_name in ['Ducts', 'Duct Fittings', 'Duct Accessories', 'Flex Ducts']:
                width_param = mep_element.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
                height_param = mep_element.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
                
                if width_param and height_param:
                    width = width_param.AsDouble()
                    height = height_param.AsDouble()
                    # Check if dimensions are equal (within tolerance) - indicates round duct
                    tolerance = 0.01  # 0.01 feet tolerance
                    if abs(width - height) < tolerance:
                        return True
        
        return False
        
    except Exception as e:
        return False


def place_family_at_point(point, mep_diameter, structural_depth, mep_element, struct_element):
    """Place a family instance at the specified point with custom parameters"""
    try:
        # Check if MEP is cylindrical or rectangular
        is_cylindrical = is_cylindrical_mep(mep_element)
        is_rectangular = is_rectangular_mep(mep_element)
        is_horizontal = is_horizontal_element(struct_element)
        
        # Determine which family to place based on MEP geometry and structural orientation
        if is_rectangular:
            if is_horizontal:
                bes_resa_type_name = "BES_RESA RECT HORIZONTAL"
            else:
                bes_resa_type_name = "BES_RESA RECT VERTICAL"
            use_rectangular_params = True
        elif is_cylindrical:
            if is_horizontal:
                bes_resa_type_name = "BES_RESA CIRC HORIZONTAL"
            else:
                bes_resa_type_name = "BES_RESA CIRC VERTICAL"
            use_rectangular_params = False
        else:
            # Ask user to select family for other cases
            use_rectangular_params = False
            bes_resa_type_name = None
        
        if bes_resa_type_name:
            # Automatically use BES_RESA family
            collector = FilteredElementCollector(doc)
            family_symbols = collector.OfClass(FamilySymbol).ToElements()
            
            selected_symbol = None
            for symbol in family_symbols:
                if symbol.Family:
                    family_name = symbol.Family.Name
                    # Get the type name
                    type_name_param = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    if type_name_param:
                        type_name = type_name_param.AsString()
                        # Create full name: Family Name + Type Name
                        full_name = "{} {}".format(family_name, type_name)
                        
                        # Check if it matches what we're looking for
                        if full_name == bes_resa_type_name or type_name == bes_resa_type_name:
                            selected_symbol = symbol
                            break
            
            if not selected_symbol:
                forms.alert(
                    '{} type not found in the project.\n'.format(bes_resa_type_name) +
                    'Please load the BES_RESA family with this type.',
                    title='Error',
                    warn_icon=True
                )
                return None
        else:
            # Not cylindrical - ask user to select family
            collector = FilteredElementCollector(doc)
            family_symbols = collector.OfClass(FamilySymbol).ToElements()
            
            if not family_symbols:
                forms.alert(
                    'No family symbols found in the project. Please load a family first.',
                    title='Error',
                    warn_icon=True
                )
                return None
            
            # Create a dictionary of family symbols for selection
            family_dict = {}
            for symbol in family_symbols:
                if symbol.Family:
                    family_name = symbol.Family.Name
                    symbol_name = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                    display_name = "{}: {}".format(family_name, symbol_name)
                    family_dict[display_name] = symbol
            
            # Let user select a family symbol
            selected_family_name = forms.SelectFromList.show(
                sorted(family_dict.keys()),
                title='Select Family Type to Place',
                width=500,
                height=600,
                button_name='Select'
            )
            
            if not selected_family_name:
                # forms.alert('No family selected. Placement cancelled.', title='Cancelled', warn_icon=False)
                return None
            
            selected_symbol = family_dict[selected_family_name]
        
        # Start a transaction - moved here before ANY Revit modifications
        t = Transaction(doc, "Place Family at Intersection")
        t.Start()
        
        try:
            # Activate the symbol if it's not already active
            if not selected_symbol.IsActive:
                selected_symbol.Activate()
                doc.Regenerate()
            
            # Create the family instance at the intersection point
            new_instance = doc.Create.NewFamilyInstance(
                point,
                selected_symbol,
                Structure.StructuralType.NonStructural
            )
            
            # Try to rotate the instance to align with MEP direction
            # Skip rotation for fittings and accessories as they may not have proper direction info
            try:
                if mep_element.Category:
                    cat_name = mep_element.Category.Name
                    # Rotate for both cylindrical and rectangular linear elements (not fittings or accessories)
                    if cat_name in ['Pipes', 'Ducts', 'Conduits', 'Cable Trays', 'Flex Pipes', 'Flex Ducts']:
                        # Get MEP direction vector
                        mep_direction = get_mep_direction(mep_element)
                        
                        # Project to XY plane (horizontal direction)
                        mep_direction_xy = XYZ(mep_direction.X, mep_direction.Y, 0)
                        length_xy = (mep_direction_xy.X**2 + mep_direction_xy.Y**2)**0.5
                        
                        if length_xy > 0.001:
                            # Normalize
                            mep_direction_xy = XYZ(mep_direction_xy.X / length_xy, mep_direction_xy.Y / length_xy, 0)
                            
                            # Calculate rotation angle from X-axis
                            angle = math.atan2(mep_direction_xy.Y, mep_direction_xy.X)
                            
                            # Create rotation axis (Z-axis through the point)
                            rotation_axis = Line.CreateBound(point, XYZ(point.X, point.Y, point.Z + 10))
                            
                            # Rotate the instance to align with MEP direction
                            # This works for both cylindrical and rectangular MEP
                            ElementTransformUtils.RotateElement(doc, new_instance.Id, rotation_axis, angle)
            except Exception as rotation_error:
                # If rotation fails, just continue without rotation
                # The family will be placed but not rotated
                pass
            
            # Set parameters based on MEP type
            param_warnings = []
            
            if use_rectangular_params:
                # For rectangular MEP - set width, height, and depth
                
                # Get MEP width and height
                mep_width = get_mep_width(mep_element)
                mep_height = get_mep_height(mep_element)
                
                # Set the width parameter (BES_RESA Largeur)
                width_param = new_instance.LookupParameter("BES_RESA Largeur")
                if width_param and not width_param.IsReadOnly:
                    width_param.Set(mep_width)
                else:
                    # Try without space
                    width_param = new_instance.LookupParameter("BES_RESA_Largeur")
                    if width_param and not width_param.IsReadOnly:
                        width_param.Set(mep_width)
                
                if not width_param or width_param.IsReadOnly:
                    param_warnings.append('Could not find or set "BES_RESA Largeur" parameter.')
                
                # Set the height parameter (BES_RESA Hauteur)
                height_param = new_instance.LookupParameter("BES_RESA Hauteur")
                if height_param and not height_param.IsReadOnly:
                    height_param.Set(mep_height)
                else:
                    # Try without space
                    height_param = new_instance.LookupParameter("BES_RESA_Hauteur")
                    if height_param and not height_param.IsReadOnly:
                        height_param.Set(mep_height)
                
                if not height_param or height_param.IsReadOnly:
                    param_warnings.append('Could not find or set "BES_RESA Hauteur" parameter.')
                
                # Set the depth parameter (BES_RESA Profondeur) - structural thickness + 55mm
                depth_param = new_instance.LookupParameter("BES_RESA Profondeur")
                if depth_param and not depth_param.IsReadOnly:
                    depth_param.Set(structural_depth)
                else:
                    # Try without space
                    depth_param = new_instance.LookupParameter("BES_RESA_Profondeur")
                    if depth_param and not depth_param.IsReadOnly:
                        depth_param.Set(structural_depth)
                
                if not depth_param or depth_param.IsReadOnly:
                    param_warnings.append('Could not find or set "BES_RESA Profondeur" parameter.')
            else:
                # For cylindrical MEP - set diameter and depth
                
                # Set the diameter parameter (BES_RESA Diameter)
                diameter_param = new_instance.LookupParameter("BES_RESA Diameter")
                if diameter_param and not diameter_param.IsReadOnly:
                    diameter_param.Set(mep_diameter)
                else:
                    # Try without space
                    diameter_param = new_instance.LookupParameter("BES_RESA_Diameter")
                    if diameter_param and not diameter_param.IsReadOnly:
                        diameter_param.Set(mep_diameter)
                
                if not diameter_param or diameter_param.IsReadOnly:
                    param_warnings.append('Could not find or set "BES_RESA Diameter" parameter.')
                
                # Set the depth parameter (BES_RESA Profondeur)
                depth_param = new_instance.LookupParameter("BES_RESA Profondeur")
                if depth_param and not depth_param.IsReadOnly:
                    depth_param.Set(structural_depth)
                else:
                    # Try without space
                    depth_param = new_instance.LookupParameter("BES_RESA_Profondeur")
                    if depth_param and not depth_param.IsReadOnly:
                        depth_param.Set(structural_depth)
                
                if not depth_param or depth_param.IsReadOnly:
                    param_warnings.append('Could not find or set "BES_RESA Profondeur" parameter.')
            
            # Check if parameters were set successfully
            # param_warnings list is already populated above
            
            # Commit the transaction
            t.Commit()
            
            # Show warnings after commit if any
            # if param_warnings:
            #     forms.alert(
            #         'Warning:\n' + '\n'.join(param_warnings) + 
            #         '\n\nMake sure the family has these parameters.',
            #         title='Parameter Warning',
            #         warn_icon=True
            #     )
            
            # forms.alert(
            #     'Family instance placed successfully!\n\n' +
            #     'Family: {}\n'.format(selected_symbol.Family.Name) +
            #     'Type: {}\n'.format(selected_symbol.Name) +
            #     'Location: ({:.4f}, {:.4f}, {:.4f})\n\n'.format(point.X, point.Y, point.Z) +
            #     'BES_RESA Diameter: {:.4f} ft ({:.2f} mm)\n'.format(mep_diameter, mep_diameter * 304.8) +
            #     'BES_RESA Profondeur: {:.4f} ft ({:.2f} mm)'.format(structural_depth, structural_depth * 304.8),
            #     title='Success',
            #     warn_icon=False
            # )
            
            # Select the newly placed instance (after commit)
            from System.Collections.Generic import List
            element_ids = List[ElementId]()
            element_ids.Add(new_instance.Id)
            uidoc.Selection.SetElementIds(element_ids)
            
            return new_instance
            
        except Exception as e:
            t.RollBack()
            forms.alert(
                'Error placing family instance:\n{}\n\n'.format(str(e)) +
                'Make sure the selected family can be placed as a point-based instance.',
                title='Error',
                warn_icon=True
            )
            return None
            
    except Exception as e:
        forms.alert(
            'Error in place_family_at_point:\n{}'.format(str(e)),
            title='Error',
            warn_icon=True
        )
        return None


def check_bounding_box_intersection(struct_solids, mep_solids, struct_element, mep_element, structural_ref, mep_reference, struct_link_instance):
    """Simplified intersection check using bounding boxes for structural framing"""
    global uidoc
    try:
        # Initialize variables
        struct_bbox = None
        mep_bbox = None
        intersection_point = None
        
        # Get bounding boxes and check for overlap
        try:
            for solid in struct_solids:
                try:
                    bbox = solid.GetBoundingBox()
                    if bbox:
                        if struct_bbox is None:
                            struct_bbox = bbox
                        else:
                            # Expand to include this bbox
                            struct_bbox.Min = XYZ(
                                min(struct_bbox.Min.X, bbox.Min.X),
                                min(struct_bbox.Min.Y, bbox.Min.Y),
                                min(struct_bbox.Min.Z, bbox.Min.Z)
                            )
                            struct_bbox.Max = XYZ(
                                max(struct_bbox.Max.X, bbox.Max.X),
                                max(struct_bbox.Max.Y, bbox.Max.Y),
                                max(struct_bbox.Max.Z, bbox.Max.Z)
                            )
                except Exception as e:
                    continue
        except Exception as e:
            forms.alert(
                'Error getting structural bounding box:\n{}'.format(str(e)),
                title='Error',
                warn_icon=True
            )
            return False
        
        try:
            for solid in mep_solids:
                try:
                    bbox = solid.GetBoundingBox()
                    if bbox:
                        if mep_bbox is None:
                            mep_bbox = bbox
                        else:
                            # Expand to include this bbox
                            mep_bbox.Min = XYZ(
                                min(mep_bbox.Min.X, bbox.Min.X),
                                min(mep_bbox.Min.Y, bbox.Min.Y),
                                min(mep_bbox.Min.Z, bbox.Min.Z)
                            )
                            mep_bbox.Max = XYZ(
                                max(mep_bbox.Max.X, bbox.Max.X),
                                max(mep_bbox.Max.Y, bbox.Max.Y),
                                max(mep_bbox.Max.Z, bbox.Max.Z)
                            )
                except Exception as e:
                    continue
        except Exception as e:
            forms.alert(
                'Error getting MEP bounding box:\n{}'.format(str(e)),
                title='Error',
                warn_icon=True
            )
            return False
        
        if not struct_bbox or not mep_bbox:
            forms.alert(
                'Could not get bounding boxes for intersection check.',
                title='Warning',
                warn_icon=True
            )
            return False
        
        # Check if bounding boxes overlap
        try:
            overlap_x = not (struct_bbox.Max.X < mep_bbox.Min.X or mep_bbox.Max.X < struct_bbox.Min.X)
            overlap_y = not (struct_bbox.Max.Y < mep_bbox.Min.Y or mep_bbox.Max.Y < struct_bbox.Min.Y)
            overlap_z = not (struct_bbox.Max.Z < mep_bbox.Min.Z or mep_bbox.Max.Z < struct_bbox.Min.Z)
            
            intersection_found = overlap_x and overlap_y and overlap_z
        except Exception as e:
            forms.alert(
                'Error checking overlap:\n{}'.format(str(e)),
                title='Error',
                warn_icon=True
            )
            return False
        
        if intersection_found:
            try:
                # Calculate approximate intersection point (center of overlap region)
                intersection_point = XYZ(
                    (max(struct_bbox.Min.X, mep_bbox.Min.X) + min(struct_bbox.Max.X, mep_bbox.Max.X)) / 2.0,
                    (max(struct_bbox.Min.Y, mep_bbox.Min.Y) + min(struct_bbox.Max.Y, mep_bbox.Max.Y)) / 2.0,
                    (max(struct_bbox.Min.Z, mep_bbox.Min.Z) + min(struct_bbox.Max.Z, mep_bbox.Max.Z)) / 2.0
                )
            except Exception as e:
                forms.alert(
                    'Error calculating intersection point:\n{}'.format(str(e)),
                    title='Error',
                    warn_icon=True
                )
                return False
            
            # Ask user if they want to place a family at the intersection point
            try:
                if forms.alert(
                    'Bounding boxes overlap - intersection likely!\n\n' +
                    'Do you want to place a family instance at the approximate intersection point?',
                    title='Place Family',
                    yes=True,
                    no=True
                ):
                    # Get MEP diameter with insulation
                    mep_diameter = get_mep_diameter_with_insulation(mep_element)
                    
                    # Get structural depth (no margin)
                    structural_depth = get_structural_depth(struct_element)
                    
                    # Place family at intersection point with parameters
                    place_family_at_point(intersection_point, mep_diameter, structural_depth, mep_element, struct_element)
                
                # Highlight both elements
                uidoc.Selection.SetReferences([structural_ref, mep_reference])
            except Exception as e:
                forms.alert(
                    'Error placing family:\n{}'.format(str(e)),
                    title='Error',
                    warn_icon=True
                )
            
            return True
        else:
            return False
            
    except Exception as e:
        forms.alert(
            'Error during bounding box intersection check:\n{}\nType: {}'.format(str(e), type(e).__name__),
            title='Error',
            warn_icon=True
        )
        return False


def check_geometry_intersection(struct_element, struct_link_instance, mep_element, mep_link_instance, structural_ref, mep_reference):
    """Check if two elements' geometries intersect and find intersection points"""
    global uidoc, doc
    try:
        # Get geometry options
        options = Options()
        options.ComputeReferences = True
        options.DetailLevel = ViewDetailLevel.Fine
        options.IncludeNonVisibleObjects = False
        
        # Get the transforms for both link instances
        struct_transform = struct_link_instance.GetTotalTransform()
        mep_transform = mep_link_instance.GetTotalTransform()
        
        # Check if structural element is framing - use simplified approach
        is_structural_framing = False
        if struct_element.Category:
            is_structural_framing = (struct_element.Category.Name == 'Structural Framing')
        
        # Get geometry for structural element
        struct_geometry = struct_element.get_Geometry(options)
        if not struct_geometry:
            forms.alert(
                'Could not retrieve geometry for structural element.\n' +
                'Element may not have visible geometry or may be a symbolic representation.',
                title='Geometry Error',
                warn_icon=True
            )
            return False
        
        # Get geometry for MEP element
        mep_geometry = mep_element.get_Geometry(options)
        if not mep_geometry:
            forms.alert(
                'Could not retrieve geometry for MEP element.\n' +
                'Element may not have visible geometry or may be a symbolic representation.',
                title='Geometry Error',
                warn_icon=True
            )
            return False
        
        # Collect all solids from structural element
        struct_solids = []
        try:
            if struct_geometry:
                for geom_obj in struct_geometry:
                    try:
                        if geom_obj and isinstance(geom_obj, Solid):
                            try:
                                if geom_obj.Volume > 0:
                                    # Transform the solid to the current document coordinate system
                                    try:
                                        transformed_solid = SolidUtils.CreateTransformed(geom_obj, struct_transform)
                                        if transformed_solid and transformed_solid.Volume > 0:
                                            struct_solids.append(transformed_solid)
                                    except:
                                        pass
                            except:
                                pass
                        elif geom_obj and isinstance(geom_obj, GeometryInstance):
                            try:
                                inst_geom = geom_obj.GetInstanceGeometry()
                                if inst_geom:
                                    for inst_obj in inst_geom:
                                        try:
                                            if inst_obj and isinstance(inst_obj, Solid):
                                                try:
                                                    if inst_obj.Volume > 0:
                                                        try:
                                                            transformed_solid = SolidUtils.CreateTransformed(inst_obj, struct_transform)
                                                            if transformed_solid and transformed_solid.Volume > 0:
                                                                struct_solids.append(transformed_solid)
                                                        except:
                                                            pass
                                                except:
                                                    pass
                                        except:
                                            pass
                            except:
                                pass
                    except:
                        pass
        except Exception as e:
            forms.alert(
                'Error processing structural element geometry:\n{}\nType: {}'.format(str(e), type(e).__name__),
                title='Geometry Error',
                warn_icon=True
            )
            return False
        
        # Collect all solids from MEP element
        mep_solids = []
        try:
            if mep_geometry:
                for geom_obj in mep_geometry:
                    try:
                        if geom_obj and isinstance(geom_obj, Solid):
                            try:
                                if geom_obj.Volume > 0:
                                    # Transform the solid to the current document coordinate system
                                    try:
                                        transformed_solid = SolidUtils.CreateTransformed(geom_obj, mep_transform)
                                        if transformed_solid and transformed_solid.Volume > 0:
                                            mep_solids.append(transformed_solid)
                                    except:
                                        pass
                            except:
                                pass
                        elif geom_obj and isinstance(geom_obj, GeometryInstance):
                            try:
                                inst_geom = geom_obj.GetInstanceGeometry()
                                if inst_geom:
                                    for inst_obj in inst_geom:
                                        try:
                                            if inst_obj and isinstance(inst_obj, Solid):
                                                try:
                                                    if inst_obj.Volume > 0:
                                                        try:
                                                            transformed_solid = SolidUtils.CreateTransformed(inst_obj, mep_transform)
                                                            if transformed_solid and transformed_solid.Volume > 0:
                                                                mep_solids.append(transformed_solid)
                                                        except:
                                                            pass
                                                except:
                                                    pass
                                        except:
                                            pass
                            except:
                                pass
                    except:
                        pass
        except Exception as e:
            forms.alert(
                'Error processing MEP element geometry:\n{}\nType: {}'.format(str(e), type(e).__name__),
                title='Geometry Error',
                warn_icon=True
            )
            return False
        
        if not struct_solids:
            forms.alert('No valid solids found in structural element.', title='Warning', warn_icon=True)
            return False
        
        if not mep_solids:
            forms.alert('No valid solids found in MEP element.', title='Warning', warn_icon=True)
            return False
        
        # Disabled structural framing special handling - use regular intersection for all
        # if is_structural_framing:
        #     return check_bounding_box_intersection(
        #         struct_solids, mep_solids, struct_element, mep_element,
        #         structural_ref, mep_reference, struct_link_instance
        #     )
        
        # Check for intersections using boolean operations (for all elements)
        intersection_results = []
        intersection_found = False
        
        for i, struct_solid in enumerate(struct_solids):
            for j, mep_solid in enumerate(mep_solids):
                try:
                    # Perform boolean intersection
                    intersection_solid = BooleanOperationsUtils.ExecuteBooleanOperation(
                        struct_solid, 
                        mep_solid, 
                        BooleanOperationsType.Intersect
                    )
                    
                    if intersection_solid and intersection_solid.Volume > 0.0001:  # Small threshold for floating point
                        intersection_found = True
                        
                        # Get centroid of intersection as the intersection point
                        centroid = intersection_solid.ComputeCentroid()
                        
                        # Get the bounding box to find min/max points
                        bbox = intersection_solid.GetBoundingBox()
                        
                        intersection_results.append({
                            'volume': intersection_solid.Volume,
                            'centroid': centroid,
                            'bbox_min': bbox.Min,
                            'bbox_max': bbox.Max
                        })
                        
                except Exception as e:
                    # Some geometry operations might fail, continue checking others
                    continue
        
        # Display results
        if intersection_found:
            result_text = []
            result_text.append("=== INTERSECTION DETECTED ===\n")
            result_text.append("The geometries INTERSECT!\n")
            result_text.append("Number of intersections found: {}\n".format(len(intersection_results)))
            
            for idx, result in enumerate(intersection_results):
                result_text.append("\n--- Intersection {} ---".format(idx + 1))
                result_text.append("Intersection Volume: {:.4f} cubic feet".format(result['volume']))
                result_text.append("\nIntersection Point (Centroid):")
                result_text.append("  X: {:.4f} ft".format(result['centroid'].X))
                result_text.append("  Y: {:.4f} ft".format(result['centroid'].Y))
                result_text.append("  Z: {:.4f} ft".format(result['centroid'].Z))
                result_text.append("\nBounding Box:")
                result_text.append("  Min: ({:.4f}, {:.4f}, {:.4f})".format(
                    result['bbox_min'].X, result['bbox_min'].Y, result['bbox_min'].Z
                ))
                result_text.append("  Max: ({:.4f}, {:.4f}, {:.4f})".format(
                    result['bbox_max'].X, result['bbox_max'].Y, result['bbox_max'].Z
                ))
            
            # forms.alert(
            #     '\n'.join(result_text),
            #     title='Intersection Results',
            #     warn_icon=False
            # )
            
            # Ask user if they want to place a family at the intersection point
            if forms.alert(
                'Do you want to place a family instance at the intersection point?',
                title='Place Family',
                yes=True,
                no=True
            ):
                # Get the first intersection point (centroid)
                intersection_point = intersection_results[0]['centroid']
                
                # Get MEP diameter with insulation
                mep_diameter = get_mep_diameter_with_insulation(mep_element)
                
                # Get structural depth (no margin)
                structural_depth = get_structural_depth(struct_element)
                
                # Show calculated values for debugging
                # forms.alert(
                #     'Calculated Parameters:\n\n' +
                #     'MEP Diameter + 30mm: {:.4f} ft ({:.2f} mm)\n'.format(mep_diameter, mep_diameter * 304.8) +
                #     'Structural Depth: {:.4f} ft ({:.2f} mm)\n\n'.format(structural_depth, structural_depth * 304.8) +
                #     'These values will be set to:\n' +
                #     '- BES_RESA Diameter\n' +
                #     '- BES_RESA Profondeur',
                #     title='Parameter Values',
                #     warn_icon=False
                # )
                
                # Place family at intersection point with parameters
                place_family_at_point(intersection_point, mep_diameter, structural_depth, mep_element, struct_element)
            
            # Highlight both elements
            uidoc.Selection.SetReferences([structural_ref, mep_reference])
            
        else:
            # forms.alert(
            #     '=== NO INTERSECTION ===\n\nThe geometries DO NOT intersect.',
            #     title='Intersection Results',
            #     warn_icon=False
            # )
            pass
        
        return intersection_found
        
    except Exception as e:
        tb = traceback.format_exc()
        forms.alert(
            'Error during geometry intersection check:\n{}\n\nFull traceback:\n{}'.format(str(e), tb),
            title='Error',
            warn_icon=True
        )
        return False


def main():
    try:
        # STEP 1: Select a structural element in a linked model
        forms.alert(
            'STEP 1: Select a STRUCTURAL element inside a linked model.',
            title='Select Structural Element in Link',
            warn_icon=False
        )
        
        # Create structural selection filter for linked elements
        structural_filter = StructuralLinkedElementFilter(doc)
        
        structural_ref = uidoc.Selection.PickObject(
            ObjectType.LinkedElement,
            structural_filter,
            "Select a structural element (beam, column, wall, floor, etc.) in a linked model"
        )
        
        # Get the structural element from the link
        structural_element, structural_link_doc, structural_link_instance = get_element_from_link(structural_ref)
        
        if not structural_element or not structural_link_doc:
            forms.alert('Could not retrieve structural element from linked model.', title='Error', warn_icon=True)
            return
        
        # Verify it's a structural element
        is_structural = False
        if structural_element.Category:
            structural_categories = [
                'Structural Framing', 'Structural Columns', 'Structural Foundations',
                'Floors', 'Walls', 'Structural Beam Systems', 'Structural Trusses',
                'Structural Stiffeners', 'Structural Connections'
            ]
            is_structural = structural_element.Category.Name in structural_categories
        
        if not is_structural:
            forms.alert(
                'The selected element is not a structural element.\n' +
                'Category: {}\n\n'.format(structural_element.Category.Name if structural_element.Category else 'None') +
                'Please select a structural element (beam, column, wall, floor, etc.)',
                title='Invalid Selection',
                warn_icon=True
            )
            return
        
        # Store information about the structural element
        structural_info = []
        structural_info.append("=== STRUCTURAL ELEMENT (STORED) ===\n")
        structural_info.append("Link Name: {}".format(structural_link_instance.Name))
        structural_info.append("Link Document: {}".format(structural_link_doc.Title))
        structural_info.append("Element ID: {}".format(structural_element.Id))
        
        try:
            if structural_element.Category:
                structural_info.append("Category: {}".format(structural_element.Category.Name))
            else:
                structural_info.append("Category: None")
        except:
            structural_info.append("Category: Unable to retrieve")
        
        try:
            if hasattr(structural_element, 'Symbol') and structural_element.Symbol:
                structural_info.append("Family: {}".format(structural_element.Symbol.Family.Name))
                structural_info.append("Type: {}".format(structural_element.Symbol.Name))
        except Exception as e:
            # Skip if family/type info can't be retrieved
            pass
        
        # forms.alert(
        #     '\n'.join(structural_info) + '\n\nStructural element stored successfully!',
        #     title='Structural Element Selected',
        #     warn_icon=False
        # )
        
        # STEP 2: Select an MEP element in a linked model
        forms.alert(
            'STEP 2: Now select an MEP element inside a linked model.',
            title='Select MEP Element in Link',
            warn_icon=False
        )
        
        # Create MEP selection filter for linked elements
        mep_filter = MEPLinkedElementFilter(doc)
        
        # Pick object with MEP filter
        mep_reference = uidoc.Selection.PickObject(
            ObjectType.LinkedElement,
            mep_filter,
            "Select an MEP element (duct, pipe, conduit, etc.) in a linked model"
        )
        
        if mep_reference:
            # Get the element from the link
            mep_element, link_doc, link_instance = get_element_from_link(mep_reference)
            
            if mep_element and link_doc:
                # Verify it's an MEP element
                is_mep = False
                if mep_element.Category:
                    mep_categories = [
                        'Ducts', 'Duct Fittings', 'Duct Accessories', 'Flex Ducts',
                        'Pipes', 'Pipe Fittings', 'Pipe Accessories', 'Flex Pipes',
                        'Mechanical Equipment', 'Electrical Equipment', 'Electrical Fixtures',
                        'Conduits', 'Conduit Fittings', 'Cable Trays', 'Cable Tray Fittings',
                        'Lighting Fixtures', 'Air Terminals', 'Sprinklers', 'Plumbing Fixtures'
                    ]
                    is_mep = mep_element.Category.Name in mep_categories
                
                if not is_mep:
                    forms.alert(
                        'The selected element is not an MEP element.\n' +
                        'Category: {}\n\n'.format(mep_element.Category.Name if mep_element.Category else 'None') +
                        'Please select an MEP element (duct, pipe, conduit, mechanical equipment, etc.)',
                        title='Invalid Selection',
                        warn_icon=True
                    )
                    return
                
                # Display information about both elements
                info = []
                info.append("=== SELECTION SUMMARY ===\n")
                
                # Structural element info
                info.append("STRUCTURAL ELEMENT (IN LINK):")
                info.append("  Link Name: {}".format(structural_link_instance.Name))
                info.append("  Link Document: {}".format(structural_link_doc.Title))
                info.append("  Element ID: {}".format(structural_element.Id))
                info.append("  Category: {}".format(
                    structural_element.Category.Name if structural_element.Category else "None"
                ))
                
                if hasattr(structural_element, 'Symbol') and structural_element.Symbol:
                    try:
                        info.append("  Family: {}".format(structural_element.Symbol.Family.Name))
                    except:
                        pass
                    try:
                        type_name_param = structural_element.Symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                        if type_name_param:
                            info.append("  Type: {}".format(type_name_param.AsString()))
                    except:
                        pass
                
                # MEP element info
                info.append("\nMEP ELEMENT (IN LINK):")
                info.append("  Link Name: {}".format(link_instance.Name))
                info.append("  Link Document: {}".format(link_doc.Title))
                info.append("  Element ID: {}".format(mep_element.Id))
                info.append("  Category: {}".format(
                    mep_element.Category.Name if mep_element.Category else "None"
                ))
                info.append("  Type: {}".format(mep_element.GetType().Name))
                
                # Get family and type name if applicable
                if hasattr(mep_element, 'Symbol') and mep_element.Symbol:
                    try:
                        info.append("  Family: {}".format(mep_element.Symbol.Family.Name))
                    except:
                        pass
                    try:
                        type_name_param = mep_element.Symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                        if type_name_param:
                            info.append("  Type: {}".format(type_name_param.AsString()))
                    except:
                        pass
                elif hasattr(mep_element, 'Name'):
                    try:
                        info.append("  Name: {}".format(mep_element.Name))
                    except:
                        pass
                
                # Show the information
                # forms.alert(
                #     '\n'.join(info),
                #     title='Both Elements Selected',
                #     warn_icon=False
                # )
                
                # Check for geometry intersection
                # forms.alert(
                #     'Checking for geometry intersection...',
                #     title='Analyzing Geometry',
                #     warn_icon=False
                # )
                
                try:
                    intersection_found = check_geometry_intersection(
                        structural_element, 
                        structural_link_instance,
                        mep_element, 
                        link_instance,
                        structural_ref,
                        mep_reference
                    )
                except Exception as intersection_error:
                    tb = traceback.format_exc()
                    forms.alert(
                        'Error during intersection check:\n{}\n\n'.format(str(intersection_error)) +
                        'Type: {}\n\n'.format(type(intersection_error).__name__) +
                        'Full Traceback:\n{}'.format(tb),
                        title='Intersection Error',
                        warn_icon=True
                    )
                    return
                
            else:
                forms.alert(
                    'Could not retrieve MEP element from linked model.',
                    title='Error',
                    warn_icon=True
                )
    
    except Exception as e:
        if 'cancel' not in str(e).lower():
            tb = traceback.format_exc()
            forms.alert(
                'Error: {}\n\nFull Traceback:\n{}'.format(str(e), tb),
                title='Error',
                warn_icon=True
            )


if __name__ == '__main__':
    main()