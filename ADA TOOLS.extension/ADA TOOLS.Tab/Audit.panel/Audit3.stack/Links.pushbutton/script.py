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
    """Get the actual file name of the Revit link"""
    try:
        # Get the link type
        link_type = doc.GetElement(link.GetTypeId())

        # Method 1: Check if we can get the linked document name directly
        try:
            linked_doc = link.GetLinkDocument()
            if linked_doc:
                return os.path.basename(linked_doc.PathName) if linked_doc.PathName else linked_doc.Title + ".rvt"
        except:
            pass

        # Method 2: Try to get filename from type name (most reliable for loaded links)
        type_name = link_type.Name
        if type_name:
            # Often the type name IS the filename or contains it
            if type_name.endswith('.rvt'):
                return type_name
            elif '.rvt' in type_name:
                # Extract just the .rvt part
                parts = type_name.split('.rvt')
                return parts[0] + '.rvt'
            else:
                # If no .rvt extension, add it (assuming it's a Revit file)
                return type_name + '.rvt'

        # Method 3: Try external file reference
        try:
            if hasattr(link_type, 'GetExternalFileReference'):
                exfs = link_type.GetExternalFileReference()
                if exfs:
                    try:
                        path = exfs.GetAbsolutePath()
                        if path:
                            return os.path.basename(path)
                    except:
                        pass

                    try:
                        path = exfs.GetPath()
                        if path:
                            path_str = str(path)
                            if path_str and path_str != "":
                                return os.path.basename(path_str)
                    except:
                        pass
        except:
            pass

        # Method 4: Last resort - use element name
        element_name = link.Name
        if element_name and element_name != type_name:
            if '.rvt' in element_name:
                return os.path.basename(element_name)

        return "Unknown (ID: {})".format(link.Id.IntegerValue)

    except Exception as e:
        return "Error (ID: {})".format(link.Id.IntegerValue)

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
    link_row_head = ["Revit Link File Name", "Link Instance ID", "Instance Workset", "Link Type Workset", "Link Type ID"]
    link_csv_data.append(link_row_head)

    if not revit_links:
        no_link_row = ["No Revit links found", "-", "-", "-", "-"]
        link_table_data.append(no_link_row)
        link_csv_data.append(no_link_row)
    else:
        # Process each Revit link instance (no grouping to show all instances)
        for link in revit_links:
            # Get the actual file name
            link_filename = get_revit_link_filename(link)
            
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

            row_data = [link_filename, link_instance_id, instance_workset_name, type_workset_name, link_type_id]
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
                      formats=['', '', '', '', ''])

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

        print("CAD CSV exported to: {}".format(cad_csv_filepath))
        print("Revit Links CSV exported to: {}".format(link_csv_filepath))

    except Exception as e:
        print("Error exporting CSV: {}".format(str(e)))
        print("Data collected but could not save to file.")

# Run the check
if __name__ == "__main__":
    check_model()
