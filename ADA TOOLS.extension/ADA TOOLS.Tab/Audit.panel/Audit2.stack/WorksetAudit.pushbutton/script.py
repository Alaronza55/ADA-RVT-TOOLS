# -*- coding: utf-8 -*-
"""List All Worksets - Simple Table
Lists all worksets in the current project in a simple table format.
"""

__title__ = "Worksets Audit"
__author__ = "Almog Davidson"

from pyrevit import revit, DB, forms, script
import datetime
import os
import re
import csv

# Get current document
doc = revit.doc

folder_name = doc.Title

# Check if the model is workshared
if not doc.IsWorkshared:
    forms.alert("This model is not workshared. No worksets available.", 
                title="No Worksets", 
                exitscript=True)

# Get all worksets
worksets = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)

# Prepare output
output = script.get_output()

def get_workset_element_details(workset_id):

    """Get detailed information about elements in a specific workset - ULTRA ROBUST VERSION"""
    try:
        # Create a filter for elements in the specific workset
        workset_filter = DB.ElementWorksetFilter(workset_id)

        # Get all elements in the workset
        collector = DB.FilteredElementCollector(doc).WherePasses(workset_filter)

        # Convert to list
        elements = list(collector)

        # Filter out unwanted element types and collect details
        element_details = []
        for elem in elements:
            # Skip view-specific elements if desired
            if isinstance(elem, (DB.View, DB.ViewSheet)):
                continue

            # Skip Revit Link Types (keep only instances)
            if isinstance(elem, DB.RevitLinkType):
                continue

            # Skip Legend Components
            if isinstance(elem, DB.FamilySymbol) and elem.Category and elem.Category.Name == "Legend Components":
                continue

            # Skip Mass element types (keep only instances)
            if hasattr(elem, 'Category') and elem.Category and elem.Category.Name == "Mass":
                # Only include if it's an instance (not a type)
                if isinstance(elem, DB.ElementType):
                    continue

            # Skip all ElementType objects for families that have both type and instance
            # (This helps avoid double counting for Mass and other families)
            if isinstance(elem, DB.ElementType):
                # Allow system families types but skip loadable family types
                if hasattr(elem, 'Family') and elem.Family:
                    continue

            try:
                # Get category name with multiple fallbacks
                category_name = get_robust_category_name(elem)

                # Skip if category is Legend Components (additional check)
                if category_name == "Legend Components":
                    continue

                # Get family and type names using ultra-robust method
                family_name, type_name = get_ultra_robust_family_and_type(elem, category_name)

                element_details.append({
                    'category': category_name,
                    'family': family_name,
                    'type': type_name,
                    'element_id': elem.Id.IntegerValue
                })

            except Exception as e:
                # Ultimate fallback
                element_info = get_emergency_element_info(elem)
                # Skip if emergency info indicates legend component
                if element_info['category'] != "Legend Components":
                    element_details.append(element_info)

        return element_details, len(element_details)

    except Exception as e:
        print("Error getting element details for workset {}: {}".format(workset_id, str(e)))
        return [], 0

    """Get detailed information about elements in a specific workset - ULTRA ROBUST VERSION"""
    try:
        # Create a filter for elements in the specific workset
        workset_filter = DB.ElementWorksetFilter(workset_id)

        # Get all elements in the workset
        collector = DB.FilteredElementCollector(doc).WherePasses(workset_filter)

        # Convert to list
        elements = list(collector)

        # Filter out unwanted element types and collect details
        element_details = []
        for elem in elements:
            # Skip view-specific elements if desired
            if isinstance(elem, (DB.View, DB.ViewSheet)):
                continue

            # Skip Revit Link Types (keep only instances)
            if isinstance(elem, DB.RevitLinkType):
                continue

            try:
                # Get category name with multiple fallbacks
                category_name = get_robust_category_name(elem)

                # Get family and type names using ultra-robust method
                family_name, type_name = get_ultra_robust_family_and_type(elem, category_name)

                element_details.append({
                    'category': category_name,
                    'family': family_name,
                    'type': type_name,
                    'element_id': elem.Id.IntegerValue
                })

            except Exception as e:
                # Ultimate fallback
                element_details.append(get_emergency_element_info(elem))

        return element_details, len(element_details)

    except Exception as e:
        print("Error getting element details for workset {}: {}".format(workset_id, str(e)))
        return [], 0

def get_robust_category_name(elem):
    """Get category name with multiple fallback methods"""
    try:
        if elem.Category and elem.Category.Name:
            category_name = elem.Category.Name
            # Additional check for legend components
            if category_name == "Legend Components":
                return category_name
            return category_name
    except:
        pass

    try:
        # Try built-in parameter
        cat_param = elem.get_Parameter(DB.BuiltInParameter.ELEM_CATEGORY_PARAM)
        if cat_param and cat_param.AsString():
            return cat_param.AsString()
    except:
        pass

    try:
        # Try from element type
        elem_type = doc.GetElement(elem.GetTypeId())
        if elem_type and elem_type.Category:
            return elem_type.Category.Name
    except:
        pass

    return "Unknown Category"

def get_ultra_robust_family_and_type(elem, category_name):
    """Ultra-robust method to get family and type names"""
    family_name = "N/A"
    type_name = "N/A"

    # METHOD 1: Family Instance approach
    try:
        if isinstance(elem, DB.FamilyInstance):
            if elem.Symbol:
                type_name = elem.Symbol.Name or type_name
                if elem.Symbol.Family:
                    family_name = elem.Symbol.Family.Name or family_name
    except:
        pass

    # METHOD 2: Element Type approach
    try:
        element_type = doc.GetElement(elem.GetTypeId())
        if element_type:
            if type_name == "N/A" and element_type.Name:
                type_name = element_type.Name

            # Try all possible family name sources
            if family_name == "N/A":
                # Try FamilyName property
                if hasattr(element_type, 'FamilyName') and element_type.FamilyName:
                    family_name = element_type.FamilyName
                # Try Family property
                elif hasattr(element_type, 'Family') and element_type.Family and element_type.Family.Name:
                    family_name = element_type.Family.Name
    except:
        pass

    # METHOD 3: Built-in Parameters (comprehensive list)
    parameter_attempts = [
        (DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM, 'family'),
        (DB.BuiltInParameter.ELEM_FAMILY_PARAM, 'family'),
        (DB.BuiltInParameter.SYMBOL_NAME_PARAM, 'type'),
        (DB.BuiltInParameter.ELEM_TYPE_PARAM, 'type'),
        (DB.BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM, 'both'),
        (DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME, 'family'),
        (DB.BuiltInParameter.ALL_MODEL_TYPE_NAME, 'type'),
        (DB.BuiltInParameter.KEYNOTE_PARAM, 'info'),
    ]

    for param_id, param_type in parameter_attempts:
        try:
            param = elem.get_Parameter(param_id)
            if param and param.AsString():
                value = param.AsString().strip()
                if value:
                    if param_type == 'family' and family_name == "N/A":
                        family_name = value
                    elif param_type == 'type' and type_name == "N/A":
                        type_name = value
                    elif param_type == 'both':
                        if ":" in value:
                            parts = value.split(":", 1)
                            if family_name == "N/A":
                                family_name = parts[0].strip()
                            if type_name == "N/A" and len(parts) > 1:
                                type_name = parts[1].strip()
                        elif family_name == "N/A":
                            family_name = value
        except:
            continue

    # Also try on element type
    try:
        element_type = doc.GetElement(elem.GetTypeId())
        if element_type:
            for param_id, param_type in parameter_attempts:
                try:
                    param = element_type.get_Parameter(param_id)
                    if param and param.AsString():
                        value = param.AsString().strip()
                        if value:
                            if param_type == 'family' and family_name == "N/A":
                                family_name = value
                            elif param_type == 'type' and type_name == "N/A":
                                type_name = value
                            elif param_type == 'both':
                                if ":" in value:
                                    parts = value.split(":", 1)
                                    if family_name == "N/A":
                                        family_name = parts[0].strip()
                                    if type_name == "N/A" and len(parts) > 1:
                                        type_name = parts[1].strip()
                except:
                    continue
    except:
        pass

    # METHOD 4: Specific Element Type Handling
    try:
        element_class = elem.GetType().Name

        # System families with known patterns
        system_family_map = {
            'Wall': 'Basic Wall',
            'Floor': 'Floor',
            'RoofBase': 'Basic Roof',
            'ExtrusionRoof': 'Basic Roof',
            'FootPrintRoof': 'Basic Roof',
            'Ceiling': 'Compound Ceiling',
            'Stairs': 'Assembled Stair',
            'Railing': 'Railing',
            'CurtainSystem': 'Curtain System',
            'Grid': 'Grid',
            'Level': 'Level',
            'ReferencePlane': 'Reference Plane',
            'Dimension': 'Linear Dimension',
            'TextNote': 'Text',
            'IndependentTag': 'Multi-Category Tag',
            'SpatialElementTag': 'Room Tag',
            'AreaTag': 'Area Tag',
            'SpotDimension': 'Spot Dimension',
            'SpotElevation': 'Spot Elevation',
            'SpotCoordinate': 'Spot Coordinate',
            'DetailLine': 'Detail Line',
            'ModelLine': 'Model Line',
            'SymbolicCurve': 'Symbolic Line',
            'CurveElement': 'Line',
            'FilledRegion': 'Filled Region',
            'Group': 'Group',
            'AssemblyInstance': 'Assembly',
        }

        for class_key, default_family in system_family_map.items():
            if class_key in element_class:
                if family_name == "N/A":
                    family_name = default_family
                if type_name == "N/A":
                    # Try to get more specific type
                    elem_type = doc.GetElement(elem.GetTypeId())
                    if elem_type and elem_type.Name:
                        type_name = elem_type.Name
                    else:
                        type_name = default_family
                break
    except:
        pass

    # METHOD 5: Host/Hosted relationship analysis
    try:
        if family_name == "N/A" or type_name == "N/A":
            if isinstance(elem, DB.FamilyInstance):
                # Check if it's hosted
                if elem.Host:
                    host_category = elem.Host.Category.Name if elem.Host.Category else "Unknown Host"
                    if family_name == "N/A":
                        family_name = "Hosted on " + host_category
    except:
        pass

    # METHOD 6: Material and compound structure analysis
    try:
        if isinstance(elem, (DB.Wall, DB.Floor, DB.RoofBase, DB.Ceiling)) and type_name == "N/A":
            elem_type = doc.GetElement(elem.GetTypeId())
            if elem_type:
                # Try to get compound structure info
                if hasattr(elem_type, 'GetCompoundStructure'):
                    compound = elem_type.GetCompoundStructure()
                    if compound:
                        layer_count = compound.LayerCount
                        if layer_count > 1:
                            type_name = elem_type.Name or "Compound " + category_name
                        else:
                            type_name = elem_type.Name or "Simple " + category_name
    except:
        pass

    # METHOD 7: Workset and phase information as last resort
    try:
        if family_name == "N/A":
            # Use category-based family assignment
            category_family_map = {
                'Doors': 'Door Family',
                'Windows': 'Window Family',
                'Furniture': 'Furniture Family',
                'Lighting Fixtures': 'Lighting Family',
                'Plumbing Fixtures': 'Plumbing Family',
                'Mechanical Equipment': 'Mechanical Family',
                'Electrical Equipment': 'Electrical Family',
                'Electrical Fixtures': 'Electrical Family',
                'Specialty Equipment': 'Equipment Family',
                'Casework': 'Casework Family',
                'Entourage': 'Entourage Family',
                'Parking': 'Parking Family',
                'Planting': 'Planting Family',
                'Site': 'Site Family',
                'Structural Framing': 'Structural Framing',
                'Structural Columns': 'Structural Column',
                'Structural Foundations': 'Structural Foundation',
                'Pipes': 'Pipe Family',
                'Ducts': 'Duct Family',
                'Cable Trays': 'Cable Tray Family',
                'Conduits': 'Conduit Family',
            }

            family_name = category_family_map.get(category_name, "System - " + category_name)
    except:
        pass

    # METHOD 8: Element properties deep dive
    try:
        if type_name == "N/A":
            # Try element's Name property
            if hasattr(elem, 'Name') and elem.Name:
                type_name = elem.Name
            # Try LookupParameter for common type names
            elif hasattr(elem, 'LookupParameter'):
                type_params = ['Type', 'Style', 'Mark', 'Model', 'Size']
                for param_name in type_params:
                    try:
                        param = elem.LookupParameter(param_name)
                        if param and param.AsString():
                            type_name = param.AsString()
                            break
                    except:
                        continue
    except:
        pass

    # METHOD 9: Geometry-based classification
    try:
        if family_name == "N/A" and hasattr(elem, 'Geometry'):
            # This is a very advanced approach - analyzing geometry
            geom_options = DB.Options()
            geom = elem.get_Geometry(geom_options)
            if geom:
                solid_count = 0
                curve_count = 0
                for geom_obj in geom:
                    if isinstance(geom_obj, DB.Solid):
                        solid_count += 1
                    elif isinstance(geom_obj, DB.Curve):
                        curve_count += 1

                if solid_count > 0:
                    family_name = "3D - " + category_name
                elif curve_count > 0:
                    family_name = "Linear - " + category_name
    except:
        pass

    # METHOD 10: Final cleanup and validation
    if family_name == "N/A":
        family_name = "Unknown - " + category_name

    if type_name == "N/A":
        # Last resort: use element class name cleaned up
        try:
            element_class = elem.GetType().Name
            # Clean up .NET class names
            if "Autodesk.Revit.DB." in element_class:
                element_class = element_class.replace("Autodesk.Revit.DB.", "")
            # Remove common suffixes
            element_class = element_class.replace("Element", "").replace("Instance", "")
            type_name = element_class or "Unknown Type"
        except:
            type_name = "Unknown Type"

    # Final validation - ensure strings are not empty
    family_name = family_name.strip() if family_name and family_name.strip() else "Unknown Family"
    type_name = type_name.strip() if type_name and type_name.strip() else "Unknown Type"

    return family_name, type_name

def get_emergency_element_info(elem):
    """Emergency fallback when everything else fails"""
    try:
        category_name = elem.Category.Name if elem.Category else "Unknown Category"
    except:
        category_name = "Error Category"

    try:
        element_id = elem.Id.IntegerValue
    except:
        element_id = "Unknown ID"

    try:
        element_class = elem.GetType().Name
        if "Autodesk.Revit.DB." in element_class:
            element_class = element_class.replace("Autodesk.Revit.DB.", "")
    except:
        element_class = "Unknown Element"

    return {
        'category': category_name,
        'family': "Emergency - " + element_class,
        'type': "Emergency - " + element_class,
        'element_id': element_id
    }

def Workset_Audit():
    results_workset = []
    detailed_elements = {}  # Dictionary to store detailed element info for each workset

    print("=== WORKSETS IN CURRENT PROJECT ===")

    if worksets:
        workset_list = list(worksets)

        # Sort worksets by name
        workset_list.sort(key=lambda x: x.Name)

        # Create table header
        print("{:<30} {:<8} {:<20} {:<10} {:<8} {:<10} {:<12}".format(
            "Workset Name|", "ID|", "Owner|", "Editable|", "Open|", "Visible|", "Element Count|"
        ))

        # Create table rows
        for workset in workset_list:
            workset_name = workset.Name
            workset_id = str(workset.Id.IntegerValue)
            workset_owner = workset.Owner if workset.Owner else "No Owner"
            is_editable = "Yes" if workset.IsEditable else "No"
            is_open = "Yes" if workset.IsOpen else "No"
            is_visible = "Yes" if workset.IsVisibleByDefault else "No"

            # Get element details and count for this workset
            element_details, element_count = get_workset_element_details(workset.Id)
            detailed_elements[workset_name] = element_details

            print("{:<30}| {:<8}| {:<20}| {:<10}| {:<8}| {:<10}| {:<12}|".format(
                workset_name[:29],
                workset_id,
                workset_owner[:19],
                is_editable,
                is_open,
                is_visible,
                str(element_count)
            ))

            # Add to results for CSV export (without truncation)
            results_workset.append([
                workset_name,
                workset_id,
                workset_owner,
                is_editable,
                is_open,
                is_visible,
                element_count
            ])

    else:
        print("No user worksets found in this project.")
        results_workset.append(["No user worksets found in this project.", "", "", "", "", "", ""])

    return results_workset, detailed_elements

def save_to_csv(detailed_elements):
    """Save the detailed audit results to CSV file"""

    # Define the output folder - change this path as needed
    output_folder = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker\{}".format(folder_name)

    # Create the folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Clean document title for filename (remove invalid characters)
    doc_title = doc.Title or "Unknown"
    clean_title = re.sub(r'[<>:"/\\|?*]', '_', doc_title)

    # Save detailed element breakdown
    filename_detailed = "Workset_Audit_Detailed.csv"
    filepath_detailed = os.path.join(output_folder, filename_detailed)

    try:
        # Save detailed file
        with open(filepath_detailed, 'w') as csvfile:
            writer = csv.writer(csvfile)

            # Write timestamp header
            writer.writerow(["Workset Audit Report - Detailed Element Breakdown"])
            writer.writerow(["Generated on: {}".format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))])
            writer.writerow(["Document: {}".format(doc.Title or "Unknown")])
            writer.writerow([])  # Empty row for spacing

            # Write column headers
            writer.writerow(["Workset Name", "Category", "Family Name", "Type Name", "Element ID"])

            # Write detailed data
            for workset_name, elements in detailed_elements.items():
                if not elements:
                    writer.writerow([workset_name, "No elements found", "", "", ""])
                else:
                    for element in elements:
                        writer.writerow([
                            workset_name,
                            element['category'],
                            element['family'],
                            element['type'],
                            element['element_id']
                        ])

        print("Detailed CSV report saved to: {}".format(filepath_detailed))
        return filepath_detailed

    except Exception as e:
        print("Error saving CSV file: {}".format(str(e)))
        return None

if __name__ == '__main__':
    # Run the audit and get results
    results_workset, detailed_elements = Workset_Audit()

    # Save results to CSV
    save_to_csv(detailed_elements)
