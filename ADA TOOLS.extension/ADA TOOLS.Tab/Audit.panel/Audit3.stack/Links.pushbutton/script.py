# -*- coding: utf-8 -*-

from pyrevit import script, revit, DB, DOCS
import csv
import os

doc = DOCS.doc
folder_name = doc.Title

def collect_cadinstances():
    """Collect ImportInstance class from whole model"""
    collector = DB.FilteredElementCollector(doc)
    cadinstances = collector.OfClass(DB.ImportInstance).WhereElementIsNotElementType().ToElements()
    return cadinstances if cadinstances else []

def collect_revitlinks():
    """Collect RevitLinkInstance class from whole model"""
    collector = DB.FilteredElementCollector(doc)
    revitlinks = collector.OfClass(DB.RevitLinkInstance).WhereElementIsNotElementType().ToElements()
    return revitlinks if revitlinks else []

def get_load_stat(cad, is_link):
    """Loaded status from the import instance's CADLinkType"""
    cad_type = doc.GetElement(cad.GetTypeId())

    if not is_link:
        return "IMPORTED"
    exfs = cad_type.GetExternalFileReference()
    status = exfs.GetLinkedFileStatus().ToString()
    if status == "Loaded":
        return "Loaded"
    if status == "NotFound":
        return "NotFound"
    if status == "Unloaded":
        return "Unloaded"
    return status

def get_views_containing_cad(cad):
    """Get all views where the CAD instance appears"""
    views_list = []

    # If CAD has an owner view, it's only in that specific view
    if cad.OwnerViewId != DB.ElementId.InvalidElementId:
        owner_view = doc.GetElement(cad.OwnerViewId)
        if owner_view:
            views_list.append(owner_view.Name)
        return views_list

    # If placed on workplane/level, check all views that can display it
    collector = DB.FilteredElementCollector(doc)
    all_views = collector.OfClass(DB.View).WhereElementIsNotElementType().ToElements()

    for view in all_views:
        try:
            # Skip template views and 3D views that might not work well
            if view.IsTemplate:
                continue

            # Get all CAD instances visible in this view
            view_collector = DB.FilteredElementCollector(doc, view.Id)
            view_cads = view_collector.OfClass(DB.ImportInstance).WhereElementIsNotElementType().ToElements()

            # Check if our CAD is in this view
            for view_cad in view_cads:
                if view_cad.Id == cad.Id:
                    views_list.append(view.Name)
                    break

        except Exception as e:
            # Some views might not support this, continue with others
            continue

    return views_list

def get_revit_link_filename(link):
    """Get the actual file name of the Revit link with improved fallback methods"""
    try:
        link_type = doc.GetElement(link.GetTypeId())
        
        if not link_type:
            return "Unknown_Type_{}.rvt".format(link.Id.IntegerValue)

        # Method 1: Try to get from loaded document (most reliable)
        try:
            linked_doc = link.GetLinkDocument()
            if linked_doc and linked_doc.Title:
                title = linked_doc.Title.strip()
                if title and not title.startswith("Error"):
                    if not title.lower().endswith('.rvt'):
                        title += '.rvt'
                    return title
        except:
            pass

        # Method 2: Try ExternalFileReference paths
        try:
            exfs = link_type.GetExternalFileReference()
            if exfs:
                # Try GetPath with ModelPathUtils
                try:
                    model_path = exfs.GetPath()
                    if model_path:
                        user_path = DB.ModelPathUtils.ConvertModelPathToUserVisiblePath(model_path)
                        if user_path:
                            filename = os.path.basename(str(user_path))
                            if filename and len(filename) > 4 and not filename.startswith("Error"):
                                return filename
                except:
                    pass
                
                # Try GetAbsolutePath
                try:
                    abs_path = exfs.GetAbsolutePath()
                    if abs_path:
                        filename = os.path.basename(str(abs_path))
                        if filename and len(filename) > 4:
                            return filename
                except:
                    pass
        except:
            pass

        # Method 3: Parse the type name intelligently
        try:
            type_name = link_type.Name
            if type_name:
                parsed_name = parse_revit_link_type_name(type_name)
                if parsed_name and len(parsed_name) > 4:
                    return parsed_name
        except:
            pass

        # Method 4: Try using Parameter to get the link name
        try:
            param = link_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            if param and param.HasValue:
                param_value = param.AsString()
                if param_value and len(param_value) > 0:
                    clean_name = clean_revit_type_name(param_value)
                    if clean_name and len(clean_name) > 4:
                        return clean_name
        except:
            pass

        # Method 5: Check if there's a saved path parameter
        try:
            saved_path_param = link_type.get_Parameter(DB.BuiltInParameter.RVT_LINK_INSTANCE_PATH)
            if saved_path_param and saved_path_param.HasValue:
                saved_path = saved_path_param.AsString()
                if saved_path:
                    filename = os.path.basename(saved_path)
                    if filename and len(filename) > 4:
                        return filename
        except:
            pass

        # Method 6: Try Family Name parameter
        try:
            family_name_param = link_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_FAMILY_NAME)
            if family_name_param and family_name_param.HasValue:
                family_name = family_name_param.AsString()
                if family_name and len(family_name) > 0:
                    clean_name = clean_revit_type_name(family_name)
                    if clean_name and len(clean_name) > 4:
                        return clean_name
        except:
            pass

        # Last resort: Use type name as-is if available
        try:
            if link_type.Name and len(link_type.Name) > 0:
                clean_name = clean_revit_type_name(link_type.Name)
                if clean_name and len(clean_name) > 4:
                    return clean_name
        except:
            pass

        return "UnknownLink_{}.rvt".format(link.Id.IntegerValue)

    except Exception as e:
        return "ErrorLink_{}.rvt".format(link.Id.IntegerValue)

def parse_revit_link_type_name(type_name):
    """Parse Revit link type name to extract filename"""
    if not type_name:
        return None

    original_name = type_name.strip()

    # If the type name already contains .rvt, extract it properly
    if '.rvt' in original_name.lower():
        lower_name = original_name.lower()
        rvt_index = lower_name.find('.rvt')

        # Look backwards from .rvt to find the start of filename
        start_index = 0
        for i in range(rvt_index - 1, -1, -1):
            char = original_name[i]
            if char in ['\\', '/', ':', ' ', '\t']:
                start_index = i + 1
                break

        filename = original_name[start_index:rvt_index + 4]
        if len(filename) > 4:  # More than just ".rvt"
            return filename.strip()

    # Try to clean up the name and make it a valid filename
    clean_name = original_name

    # Remove common Revit prefixes
    prefixes = [
        "Revit Link - ",
        "Revit Link: ",
        "Link - ",
        "Link: ",
        "Type: ",
        "RVT - ",
        "RVT: ",
        "RevitLinkType ",
        "Type ",
    ]

    for prefix in prefixes:
        if clean_name.startswith(prefix):
            clean_name = clean_name[len(prefix):].strip()
            break

    # Remove file paths if present
    clean_name = clean_name.replace('\\', '/').split('/')[-1]

    # If we have a reasonable name, use it
    if clean_name and len(clean_name.strip()) > 0:
        clean_name = clean_name.strip()

        # Remove any remaining invalid characters for display
        invalid_chars = ['<', '>', '|', '"', '*', '?']
        for char in invalid_chars:
            clean_name = clean_name.replace(char, '_')

        # Add .rvt if not present
        if not clean_name.lower().endswith('.rvt'):
            clean_name += '.rvt'

        return clean_name

    return None

def clean_revit_type_name(type_name):
    """Clean up Revit link type name to get a proper filename"""
    if not type_name:
        return "Unknown.rvt"

    # Remove common prefixes
    prefixes = ["Revit Link ", "Link ", "Type ", "RVT ", "RevitLinkType "]
    clean_name = type_name.strip()
    
    for prefix in prefixes:
        if clean_name.startswith(prefix):
            clean_name = clean_name[len(prefix):].strip()

    # If it contains path separators, get just the filename
    clean_name = clean_name.replace('\\', '/').split('/')[-1]

    # Remove trailing/leading whitespace
    clean_name = clean_name.strip()

    # Ensure it ends with .rvt
    if clean_name and not clean_name.lower().endswith('.rvt'):
        clean_name += '.rvt'

    return clean_name if clean_name else "Unknown.rvt"

def get_revit_link_status(link):
    """Get the load status of a Revit link"""
    try:
        link_type = doc.GetElement(link.GetTypeId())
        if not link_type:
            return "Unknown"
        
        try:
            exfs = link_type.GetExternalFileReference()
            if exfs:
                status = exfs.GetLinkedFileStatus()
                status_str = status.ToString()
                
                # Clean up the status string for better readability
                if status_str == "Loaded":
                    return "Loaded"
                elif status_str == "NotFound":
                    return "Not Found"
                elif status_str == "Unloaded":
                    return "Unloaded"
                elif status_str == "Invalid":
                    return "Invalid"
                else:
                    return status_str
            else:
                # No external file reference - might be an in-place link
                return "No Reference"
        except:
            # If we can't get external file reference, check if doc is loaded
            try:
                linked_doc = link.GetLinkDocument()
                if linked_doc:
                    return "Loaded"
                else:
                    return "Not Loaded"
            except:
                return "Unknown"
    except Exception as e:
        return "Unknown"
def check_model():
    output = script.get_output()
    output.close_others()
    output.set_title("CAD and Revit Links audit of model '{}'".format(doc.Title))
    output.set_width(1800)

    # =========================
    # CAD INSTANCES AUDIT
    # =========================

    # Collect CAD instances
    cad_instances = collect_cadinstances()

    # Prepare data for table and CSV
    cad_table_data = []
    cad_csv_data = []
    cad_row_head = ["DWG Name", "Element ID", "Loaded Status", "Workplane or Single View", "Workset", "Count", "Views"]
    cad_csv_data.append(cad_row_head)

    if not cad_instances:
        no_cad_row = ["No CAD instances found", "-", "-", "-", "-", "-", "-"]
        cad_table_data.append(no_cad_row)
        cad_csv_data.append(no_cad_row)
    else:
        # Group CADs by name and collect all instances for each name
        cad_groups = {}

        for cad in cad_instances:
            try:
                cad_name = cad.Parameter[DB.BuiltInParameter.IMPORT_SYMBOL_NAME].AsString()
            except:
                cad_name = "Unknown"

            if cad_name not in cad_groups:
                cad_groups[cad_name] = []
            cad_groups[cad_name].append(cad)

        # Process each unique CAD name
        for cad_name, cad_list in cad_groups.items():
            # Use the first instance for getting common properties
            first_cad = cad_list[0]

            # Get element IDs of all instances (comma separated)
            element_ids = [str(cad.Id.IntegerValue) for cad in cad_list]
            cad_id = ", ".join(element_ids)

            # Get loaded status from first instance
            loaded_status = get_load_stat(first_cad, first_cad.IsLinked)

            # Check workplane/view info from first instance
            cad_own_view_id = first_cad.OwnerViewId
            if cad_own_view_id == DB.ElementId.InvalidElementId:
                # Placed on workplane/level
                try:
                    workplane_info = doc.GetElement(first_cad.LevelId).Name
                except:
                    workplane_info = "Workplane"
            else:
                # Placed on single view (warning)
                cad_own_view_name = doc.GetElement(cad_own_view_id).Name
                workplane_info = "View: {}".format(cad_own_view_name)

            # Get workset from first instance
            try:
                workset_name = revit.query.get_element_workset(first_cad).Name
            except:
                workset_name = "Unknown"

            # Count is the number of instances
            count = len(cad_list)

            # Collect views from all instances and remove duplicates
            all_views = set()
            for cad in cad_list:
                views_with_cad = get_views_containing_cad(cad)
                all_views.update(views_with_cad)

            views_string = "; ".join(sorted(all_views)) if all_views else "None found"

            # Create row data
            row_data = [cad_name, cad_id, loaded_status, workplane_info, workset_name, str(count), views_string]
            cad_table_data.append(row_data)
            cad_csv_data.append(row_data)

    # =========================
    # REVIT LINKS AUDIT
    # =========================

    # Collect Revit links
    revit_links = collect_revitlinks()

    # Prepare data for Revit links table and CSV
    link_table_data = []
    link_csv_data = []
    link_row_head = ["Revit Link File Name", "Link Instance ID", "Load Status", "Instance Workset", "Link Type Workset", "Link Type ID"]
    link_csv_data.append(link_row_head)

    if not revit_links:
        no_link_row = ["No Revit links found", "-", "-", "-", "-", "-"]
        link_table_data.append(no_link_row)
        link_csv_data.append(no_link_row)
    else:
        # Process each Revit link instance (no grouping to show all instances)
        for link in revit_links:
            # Get the actual file name
            link_filename = get_revit_link_filename(link)
            
            # Get load status
            load_status = get_revit_link_status(link)

            # Get link instance ID
            link_instance_id = str(link.Id.IntegerValue)

            # Get link type ID
            link_type_id = str(link.GetTypeId().IntegerValue)

            # Get workset of the instance
            try:
                instance_workset_name = revit.query.get_element_workset(link).Name
            except:
                instance_workset_name = "Unknown"

            # Get workset of the link type
            try:
                link_type = doc.GetElement(link.GetTypeId())
                type_workset_name = revit.query.get_element_workset(link_type).Name
            except:
                type_workset_name = "Unknown"

            row_data = [link_filename, link_instance_id, load_status, instance_workset_name, type_workset_name, link_type_id]
            link_table_data.append(row_data)
            link_csv_data.append(row_data)

    # Display tables
    output.print_md("## CAD Instances Audit")
    output.print_table(table_data=cad_table_data,
                      title="",
                      columns=cad_row_head,
                      formats=['', '', '', '', '', '', ''])

    output.print_md("## Revit Links Audit")
    output.print_table(table_data=link_table_data,
                      title="",
                      columns=link_row_head,
                      formats=['', '', '', '', '', ''])

    # Export to CSV files
    output_folder = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker\{}".format(folder_name)

    # CAD CSV
    cad_csv_filename = "CAD_Audit.csv"
    cad_csv_filepath = os.path.join(output_folder, cad_csv_filename)

    # Revit Links CSV
    link_csv_filename = "Revit_Links_Audit.csv"
    link_csv_filepath = os.path.join(output_folder, link_csv_filename)

    try:
        # Create directory if it doesn't exist
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Export CAD CSV
        with open(cad_csv_filepath, 'wb') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerows(cad_csv_data)

        # Export Revit Links CSV
        with open(link_csv_filepath, 'wb') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerows(link_csv_data)

        print("\nAudit completed!")
        print("CAD instances: Found {} unique CAD files with {} total instances".format(len(cad_groups) if cad_instances else 0, len(cad_instances)))
        print("Revit links: Found {} instances".format(len(revit_links)))

        # Print CAD duplicate summary
        if cad_instances:
            duplicates = [name for name, cads in cad_groups.items() if len(cads) > 1]
            if duplicates:
                print("CAD duplicates found:")
                for dup_name in duplicates:
                    print("  '{}' appears {} times".format(dup_name, len(cad_groups[dup_name])))
            else:
                print("No CAD duplicates found.")

        # Print unknown links warning
        unknown_links = [row[0] for row in link_table_data if row[0].startswith("UnknownLink_")]
        if unknown_links:
            print("\nWARNING: {} links could not be identified (showing as UnknownLink_xxx)".format(len(unknown_links)))
            print("These links may be unloaded, not found, or have corrupted data.")
            print("Consider reloading these links or checking their status in the model.")

        print("\nCAD CSV exported to: {}".format(cad_csv_filepath))
        print("Revit Links CSV exported to: {}".format(link_csv_filepath))

    except Exception as e:
        print("Error exporting CSV: {}".format(str(e)))
        print("Data collected but could not save to file.")

# Run the check
if __name__ == "__main__":
    check_model()
