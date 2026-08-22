# -*- coding: utf-8 -*-
__title__ = "Export View\nTemplates Usage"
__author__ = "Assistant"
__doc__ = "Export a CSV file showing view templates and where they are used"

from Autodesk.Revit.DB import FilteredElementCollector, View, ViewType
from pyrevit import revit, forms
import csv
import os
import re

# Get the current document
doc = revit.doc
folder_name = doc.Title

def get_all_view_templates():
    """Get all view templates in the document"""
    collector = FilteredElementCollector(doc).OfClass(View)
    view_templates = [v for v in collector if v.IsTemplate]
    return view_templates

def get_all_views():
    """Get all views in the document (non-templates)"""
    collector = FilteredElementCollector(doc).OfClass(View)
    views = [v for v in collector if not v.IsTemplate]
    return views

def get_view_template_usage():
    """Create a dictionary mapping view templates to the views that use them"""
    templates = get_all_view_templates()
    views = get_all_views()
    
    # Initialize dictionary with all templates
    usage_dict = {}
    for template in templates:
        template_name = template.Name
        template_id = template.Id
        usage_dict[template_name] = []
    
    # Check each view for its template
    for view in views:
        try:
            # Get the view template ID
            template_id = view.ViewTemplateId
            
            # If view has a template assigned
            if template_id and template_id.IntegerValue != -1:
                template = doc.GetElement(template_id)
                if template and template.Name in usage_dict:
                    usage_dict[template.Name].append(view.Name)
        except:
            continue
    
    return usage_dict

def export_to_csv(usage_dict):
    # Define the output folder - change this path as needed
    output_folder = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker\{}".format(folder_name)

    # Create the folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Save detailed element breakdown
    filename_detailed = "ViewTemplate_Audit_Detailed.csv"
    filepath_detailed = os.path.join(output_folder, filename_detailed)

    # Prepare data for CSV
    csv_data = []

    for template_name in sorted(usage_dict.keys()):
        views_using_template = usage_dict[template_name]

        if not views_using_template:
            # Template not used
            csv_data.append([template_name, "Not used"])
        else:
            # Add one row per view using this template
            for view_name in sorted(views_using_template):
                csv_data.append([template_name, view_name])

    # Write to CSV
    try:
        import codecs
        with codecs.open(filepath_detailed, 'w', encoding='utf-8-sig') as csvfile:  # utf-8-sig adds BOM for Excel
            writer = csv.writer(csvfile)
            # Write header
            writer.writerow(['View Template Name', 'View Name'])
            # Write data
            writer.writerows(csv_data)

        forms.alert(
            "Export successful!\n\n"
            "File saved to:\n{}".format(filepath_detailed),
            title="Success"
        )

    except Exception as e:
        forms.alert(
            "Error writing CSV file:\n{}".format(str(e)),
            title="Error"
        )

# Main execution
if __name__ == '__main__':
    try:
        # Get usage data
        usage_dict = get_view_template_usage()
        
        # Check if any templates found
        if not usage_dict:
            forms.alert(
                "No view templates found in the document.",
                title="No Templates"
            )
        else:
            # Export to CSV
            export_to_csv(usage_dict)
            
    except Exception as e:
        forms.alert(
            "Error running script:\n{}".format(str(e)),
            title="Error"
        )
