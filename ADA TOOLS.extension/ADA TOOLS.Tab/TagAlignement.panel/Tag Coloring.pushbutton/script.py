"""
Color MEP Tags by Linked Model
Colors duct tags green if their host element is from a linked model containing "HVC"
Colors pipe tags magenta/pink if their host element is from a linked model containing "BLU"
Colors pipe tags orange if their host element is from a linked model containing "CLU"
Colors pipe tags red if their host element is from a linked model containing "CLU" (alternative color)
Colors electrical tags blue if their host element is from a linked model containing "ELE"
"""

__title__ = "Color MEP Tags\nby Linked Model"
__author__ = "Your Name"

from Autodesk.Revit.DB import (
    FilteredElementCollector, 
    IndependentTag, 
    Transaction,
    Color,
    OverrideGraphicSettings,
    ElementId,
    RevitLinkInstance
)
from Autodesk.Revit.UI import TaskDialog

# Get the active document and view
doc = __revit__.ActiveUIDocument.Document
active_view = doc.ActiveView

def get_linked_model_name(element):
    """
    Get the name of the linked model that contains the element
    Returns None if element is not from a linked model
    """
    try:
        # Get all RevitLinkInstances in the document
        link_instances = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
        
        # Check if element is from a linked document
        if element.Document.Title != doc.Title:
            # Find the corresponding link instance
            for link in link_instances:
                link_doc = link.GetLinkDocument()
                if link_doc and link_doc.Title == element.Document.Title:
                    return link.Name
        return None
    except:
        return None

def get_tag_categories():
    """
    Returns the category names for duct, pipe, and electrical-related tags
    """
    return [
        "Duct Tags",
        "Duct Accessory Tags",
        "Duct Fitting Tags",
        "Pipe Tags",
        "Pipe Accessory Tags",
        "Pipe Fitting Tags",
        "Cable Tray Tags",
        "Electrical Fixture Tags",
        "Cable Tray Fitting Tags"
    ]

# Collect all tags in the active view
all_tags = FilteredElementCollector(doc, active_view.Id)\
    .OfClass(IndependentTag)\
    .ToElements()

# Filter for duct and pipe-related tags
tag_categories = get_tag_categories()
mep_tags = []

for tag in all_tags:
    try:
        category_name = tag.Category.Name
        if category_name in tag_categories:
            mep_tags.append(tag)
    except:
        continue

if not mep_tags:
    TaskDialog.Show("Info", "No duct, pipe, or electrical related tags found in the active view.")
else:
    # Create green color for HVC (RGB: 0, 77, 1)
    green_color = Color(0, 77, 1)
    
    # Create magenta/pink color for BLU (RGB: 255, 0, 232)
    magenta_color = Color(255, 0, 232)
    
    # Create orange color for CLU (RGB: 255, 127, 1)
    orange_color = Color(255, 127, 1)
    
    # Create blue color for ELE (RGB: 0, 67, 255)
    blue_color = Color(0, 67, 255)
    
    # Create override settings for green
    green_override = OverrideGraphicSettings()
    green_override.SetProjectionLineColor(green_color)
    
    # Create override settings for magenta
    magenta_override = OverrideGraphicSettings()
    magenta_override.SetProjectionLineColor(magenta_color)
    
    # Create override settings for orange
    orange_override = OverrideGraphicSettings()
    orange_override.SetProjectionLineColor(orange_color)
    
    # Create override settings for blue
    blue_override = OverrideGraphicSettings()
    blue_override.SetProjectionLineColor(blue_color)
    
    # Start transaction
    t = Transaction(doc, "Color MEP Tags by Linked Model")
    t.Start()
    
    hvc_colored_count = 0
    blu_colored_count = 0
    clu_colored_count = 0
    ele_colored_count = 0
    skipped_count = 0
    not_linked_count = 0
    debug_info = []
    
    try:
        for tag in mep_tags:
            try:
                # Get the reference to the tagged element
                tagged_ref = tag.GetTaggedReferences()
                
                if tagged_ref.Count == 0:
                    skipped_count += 1
                    continue
                
                # Get the first reference (tags typically have one reference)
                ref = list(tagged_ref)[0]
                
                # Check if element is from a linked model
                linked_element_id = ref.LinkedElementId
                link_instance_id = ref.ElementId
                
                # If LinkedElementId is valid, this is a linked element
                if linked_element_id != ElementId.InvalidElementId:
                    # Get the link instance
                    link_instance = doc.GetElement(link_instance_id)
                    
                    if link_instance:
                        link_name = link_instance.Name
                        debug_info.append("Tag ID {}: Link = {}".format(tag.Id, link_name))
                        
                        # Check if link name contains "HVC", "BLU", "CLU", or "ELE"
                        if "HVC" in link_name.upper():
                            active_view.SetElementOverrides(tag.Id, green_override)
                            hvc_colored_count += 1
                        elif "BLU" in link_name.upper():
                            active_view.SetElementOverrides(tag.Id, magenta_override)
                            blu_colored_count += 1
                        elif "CLU" in link_name.upper():
                            active_view.SetElementOverrides(tag.Id, orange_override)
                            clu_colored_count += 1
                        elif "ELE" in link_name.upper():
                            active_view.SetElementOverrides(tag.Id, blue_override)
                            ele_colored_count += 1
                        else:
                            skipped_count += 1
                    else:
                        skipped_count += 1
                else:
                    # Element is in the current document (not linked)
                    not_linked_count += 1
                    debug_info.append("Tag ID {}: Not from linked model".format(tag.Id))
                        
            except Exception as e:
                debug_info.append("Tag ID {}: Error - {}".format(tag.Id, str(e)))
                skipped_count += 1
                continue
        
        t.Commit()
        
        # # Show results
        # message = "Results:\n\n"
        # message += "Total MEP tags found: {}\n".format(len(mep_tags))
        # message += "Tags colored GREEN (HVC linked model): {}\n".format(hvc_colored_count)
        # message += "Tags colored MAGENTA (BLU linked model): {}\n".format(blu_colored_count)
        # message += "Tags colored ORANGE (CLU linked model): {}\n".format(clu_colored_count)
        # message += "Tags colored BLUE (ELE linked model): {}\n".format(ele_colored_count)
        # message += "Tags from other links or current doc: {}\n".format(skipped_count)
        # message += "Tags from current document (not linked): {}\n\n".format(not_linked_count)
        
        # # Add debug info (first 10 items)
        # if debug_info:
        #     message += "Debug Info (first 10):\n"
        #     for info in debug_info[:10]:
        #         message += "{}\n".format(info)
        
        # TaskDialog.Show("Color MEP Tags Complete", message)
        
    except Exception as e:
        t.RollBack()
        TaskDialog.Show("Error", "An error occurred: {}".format(str(e)))