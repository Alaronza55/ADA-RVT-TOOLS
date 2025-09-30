"""
Retrieve Model Elements with ID and Level
"""
__title__ = "Export Elements\nwith Level"
__author__ = "Your Name"

# Import required modules
from pyrevit import revit, DB, forms
import os

# Get current document
doc = revit.doc
folder_name = doc.Title

def clean_text(text):
    """Clean text to remove problematic characters"""
    try:
        # Convert to string and replace problematic characters
        text = str(text)
        # Replace common problematic characters
        text = text.replace('\u25B2', 'triangle')  # Triangle symbol
        text = text.replace('\u00B0', 'deg')       # Degree symbol
        text = text.replace('\u00B2', '2')         # Superscript 2
        text = text.replace('\u00B3', '3')         # Superscript 3
        text = text.replace('\u2013', '-')         # En dash
        text = text.replace('\u2014', '-')         # Em dash
        text = text.replace('\u2019', "'")         # Right single quotation mark
        text = text.replace('\u201C', '"')         # Left double quotation mark
        text = text.replace('\u201D', '"')         # Right double quotation mark
        
        # Remove any remaining non-ASCII characters
        text = ''.join(char if ord(char) < 128 else '?' for char in text)
        return text
    except:
        return "Unknown"

def get_element_level(element):
    """Get the level associated with an element"""
    try:
        # Try to get Level parameter
        level_param = element.get_Parameter(DB.BuiltInParameter.FAMILY_LEVEL_PARAM)
        if level_param and level_param.AsElementId() != DB.ElementId.InvalidElementId:
            level = doc.GetElement(level_param.AsElementId())
            return clean_text(level.Name) if level else "No Level"

        # Try to get Reference Level parameter
        ref_level_param = element.get_Parameter(DB.BuiltInParameter.FAMILY_BASE_LEVEL_PARAM)
        if ref_level_param and ref_level_param.AsElementId() != DB.ElementId.InvalidElementId:
            level = doc.GetElement(ref_level_param.AsElementId())
            return clean_text(level.Name) if level else "No Level"

        # Try to get Base Constraint parameter (for walls, floors, etc.)
        base_constraint_param = element.get_Parameter(DB.BuiltInParameter.WALL_BASE_CONSTRAINT)
        if base_constraint_param and base_constraint_param.AsElementId() != DB.ElementId.InvalidElementId:
            level = doc.GetElement(base_constraint_param.AsElementId())
            return clean_text(level.Name) if level else "No Level"

        # Try to get Level parameter for floors
        floor_level_param = element.get_Parameter(DB.BuiltInParameter.LEVEL_PARAM)
        if floor_level_param and floor_level_param.AsElementId() != DB.ElementId.InvalidElementId:
            level = doc.GetElement(floor_level_param.AsElementId())
            return clean_text(level.Name) if level else "No Level"

        # Try to get Base Level for stairs
        base_level_param = element.get_Parameter(DB.BuiltInParameter.STAIRS_BASE_LEVEL_PARAM)
        if base_level_param and base_level_param.AsElementId() != DB.ElementId.InvalidElementId:
            level = doc.GetElement(base_level_param.AsElementId())
            return clean_text(level.Name) if level else "No Level"

        # For elements that might have a Level property directly
        if hasattr(element, 'Level') and element.Level:
            return clean_text(element.Level.Name)

        return "No Level"

    except:
        return "No Level"

def get_element_category_name(element):
    """Get the category name of an element"""
    try:
        if element.Category:
            return clean_text(element.Category.Name)
        else:
            return "No Category"
    except:
        return "No Category"

def get_element_name(element):
    """Get the name of an element"""
    try:
        # Try to get the Name parameter
        if hasattr(element, 'Name') and element.Name:
            return clean_text(element.Name)

        # Try to get Type name
        element_type = doc.GetElement(element.GetTypeId())
        if element_type and hasattr(element_type, 'Name') and element_type.Name:
            return clean_text(element_type.Name)

        # Fallback to category name
        return get_element_category_name(element)

    except:
        return "Unknown Element"

def is_model_element(element):
    """Check if element is a model element that should be included"""
    try:
        # Exclude views, sheets, and other non-model elements
        if isinstance(element, (DB.View, DB.ViewSheet, DB.ProjectInfo, 
                              DB.PrintSetting, DB.ViewSheetSet, DB.Family,
                              DB.ElementType, DB.Material, DB.Level)):
            return False

        # Check if element has a category
        if not element.Category:
            return False

        # Exclude specific categories by name (more reliable than built-in category enum)
        category_name = element.Category.Name
        excluded_categories = [
            "Views", "Sheets", "Project Information", "Schedules", 
            "Legends", "Materials", "Levels", "Grids", "Scope Boxes"
        ]

        if category_name in excluded_categories:
            return False

        # Include elements that have geometry or location
        if element.Location is not None:
            return True

        # Check if element has geometry
        try:
            geom = element.get_Geometry(DB.Options())
            if geom is not None:
                return True
        except:
            pass

        return False

    except:
        return False

def create_directory_if_not_exists(directory):
    """Create directory if it doesn't exist"""
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
            print("Created directory: {}".format(directory))
        return True
    except Exception as e:
        print("Error creating directory: {}".format(str(e)))
        return False

def main():
    # Collect all elements (excluding element types)
    collector = DB.FilteredElementCollector(doc)
    elements = collector.WhereElementIsNotElementType().ToElements()

    # Filter model elements
    model_elements = []

    print("Processing {} total elements...".format(len(elements)))

    for element in elements:
        if is_model_element(element):
            model_elements.append(element)

    print("Found {} model elements".format(len(model_elements)))

    if not model_elements:
        forms.alert("No model elements found in the project.")
        return

    # Prepare data
    data = []
    for i, element in enumerate(model_elements):
        try:
            element_name = get_element_name(element)
            element_id = element.Id.IntegerValue
            level_name = get_element_level(element)
            category_name = get_element_category_name(element)

            data.append([element_name, element_id, level_name, category_name])

            # Progress indicator
            if (i + 1) % 100 == 0:
                print("Processed {} of {} elements".format(i + 1, len(model_elements)))

        except Exception as e:
            print("Error processing element {}: {}".format(element.Id, str(e)))
            continue

    if not data:
        forms.alert("No valid model elements found.")
        return

    # Sort data by category, then by element name
    data.sort(key=lambda x: (x[3], x[0]))

    # Create file path
    clean_folder_name = clean_text(folder_name)
    output_folder = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker\{}".format(clean_folder_name)
    csv_filename = "Model_Elements_with_Levels.csv"
    file_path = os.path.join(output_folder, csv_filename)

    # Create directory if it doesn't exist
    if not create_directory_if_not_exists(output_folder):
        forms.alert("Could not create output directory. Using desktop instead.")
        desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        file_path = os.path.join(desktop, csv_filename)

    try:
        # Create CSV file with explicit encoding
        with open(file_path, 'w') as f:
            # Write headers
            f.write("Element Name,Element ID,Level,Category\n")

            # Write data
            for row_data in data:
                # Clean and escape data
                escaped_row = []
                for item in row_data:
                    item_str = clean_text(str(item))
                    if ',' in item_str or '"' in item_str:
                        # Escape quotes and wrap in quotes
                        item_str = item_str.replace('"', '""')
                        escaped_row.append('"{}"'.format(item_str))
                    else:
                        escaped_row.append(item_str)

                f.write("{}\n".format(",".join(escaped_row)))

        # Show success message
        message = "Export completed successfully!\n\n"
        message += "File saved to: {}\n".format(file_path)
        message += "Total elements exported: {}\n\n".format(len(data))
        message += "The CSV file can be opened in Excel or any spreadsheet application."

        forms.alert(message)

        # Try to open the file
        try:
            os.startfile(file_path)
        except:
            print("Could not automatically open the file. Please navigate to: {}".format(file_path))

    except Exception as e:
        forms.alert("Error creating CSV file: {}".format(str(e)))
        print("Detailed error: {}".format(str(e)))

if __name__ == "__main__":
    main()
