"""Export Line Patterns with Import Status
Exports all line patterns from the project to a CSV file automatically.

"""
__title__ = "Export Line Patterns\nwith Import Status"
__author__ = "Almog Davidson"

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

import os
from pyrevit import forms, script
from datetime import datetime

# Get current document
doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

# Comprehensive list of built-in line pattern names (English Revit)
BUILTIN_PATTERN_NAMES = {
    # Basic patterns
    "Solid", "Dash", "Dot", "Dash dot", "Dash dot dot",
    "Long dash", "Long dash dot", "Long dash dot dot",
    "Dash short dash", "Dash space dash",
    
    # Standard architectural/engineering patterns
    "Hidden", "Center", "Phantom", "Border", "Divide",
    "Property Line", "Overhead", "Gas Line", "Hot Water Supply",
    "Steam Line", "Duct", "Pipe", "Conduit", "Communication Cable",
    "Zigzag", "Batting", "Tracks", "Fenceline 1", "Fenceline 2",
    
    # Special patterns
    "<Solid fill>", "By Category", "<Invisible lines>",
    
    # Additional patterns that may be in templates
    "Dash dot (2x)", "Dash dot dot (2x)", "Long dash (2x)",
    "Long dash dot (2x)", "Long dash dot dot (2x)"
}

def is_builtin_pattern(pattern_elem):
    """
    Determine if a line pattern is built-in using multiple checks.
    Returns: (is_builtin, detection_method)
    """
    pattern_name = pattern_elem.Name
    element_id = pattern_elem.Id.IntegerValue
    
    # Method 1: Check Element ID (most reliable)
    # Built-in elements typically have negative IDs or very low positive IDs
    if element_id < 0:
        return (False, "Built-in (Negative ID: {})".format(element_id))  # Negative = Built-in
    
    if element_id < 1000:
        return (False, "Built-in (Low ID: {})".format(element_id))  # Low positive = Built-in
    
    # Method 2: Check against known names
    if pattern_name in BUILTIN_PATTERN_NAMES:
        return (False, "Built-in (Known name)")
    
    # Method 3: If ID > 1000 and not in known list, likely imported
    if element_id >= 1000:
        return (True, "Imported (High ID: {})".format(element_id))
    
    # Default: assume imported if uncertain
    return (True, "Unknown (ID: {})".format(element_id))

def get_all_line_patterns_detailed():
    """Get all line patterns with detailed information including import status"""
    all_patterns = []

    # Get all line pattern elements
    line_pattern_collector = FilteredElementCollector(doc).OfClass(LinePatternElement)

    for pattern_elem in line_pattern_collector:
        try:
            pattern_name = pattern_elem.Name
            line_pattern = pattern_elem.GetLinePattern()

            # Determine if it's imported using improved methodology
            is_imported, detection_method = is_builtin_pattern(pattern_elem)

            pattern_info = {
                'Name': pattern_name,
                'Id': pattern_elem.Id.IntegerValue,
                'ElementId': pattern_elem.Id,
                'IsImported': is_imported,
                'DetectionMethod': detection_method,
                'Segments': [],
                'SegmentCount': 0
            }

            # Try to get pattern segments
            try:
                segments = line_pattern.GetSegments()
                if segments:
                    pattern_info['SegmentCount'] = len(segments)
                    for i, segment in enumerate(segments):
                        segment_info = {
                            'Index': i + 1,
                            'Type': str(segment.Type),
                            'Length': round(segment.Length, 6) if hasattr(segment, 'Length') else 'N/A'
                        }
                        pattern_info['Segments'].append(segment_info)
                else:
                    pattern_info['Segments'] = [{'Index': 1, 'Type': 'Solid', 'Length': 0}]
                    pattern_info['SegmentCount'] = 0
            except Exception as e:
                pattern_info['Segments'] = [{'Index': 1, 'Type': 'Error: ' + str(e), 'Length': 'N/A'}]
                pattern_info['SegmentCount'] = 0

            all_patterns.append(pattern_info)

        except Exception as e:
            # If we can't process a pattern, still add basic info
            try:
                pattern_info = {
                    'Name': pattern_elem.Name if hasattr(pattern_elem, 'Name') else 'Unknown',
                    'Id': pattern_elem.Id.IntegerValue,
                    'ElementId': pattern_elem.Id,
                    'IsImported': True,
                    'DetectionMethod': 'Error in processing',
                    'Segments': [{'Index': 1, 'Type': 'Error reading pattern', 'Length': str(e)}],
                    'SegmentCount': 0
                }
                all_patterns.append(pattern_info)
            except:
                pass

    return all_patterns

import codecs

def export_to_csv(patterns, file_path):
    """Export line patterns to a CSV file"""
    try:
        # Changed: Using codecs.open for Python 2.7 compatibility
        # utf-8-sig adds BOM for Excel compatibility
        with codecs.open(file_path, 'w', encoding='utf-8-sig') as f:
            # Updated header with Detection Method column
            f.write("Name,ID,Imported,Detection_Method,Segment_Count,Pattern_Definition\n")

            for pattern in patterns:
                segment_def = ""
                if pattern.get('Segments'):
                    segment_parts = []
                    for segment in pattern['Segments']:
                        seg_type = segment.get('Type', 'Unknown')
                        seg_length = segment.get('Length', 'N/A')
                        segment_parts.append("{}:{}".format(seg_type, seg_length))
                    segment_def = " | ".join(segment_parts)
                else:
                    segment_def = "No segments"

                # Clean the data for CSV
                name = str(pattern.get('Name', 'Unknown')).replace('"', '""')
                imported_status = 'Yes' if pattern.get('IsImported', False) else 'No'
                pattern_id = pattern.get('Id', 'Unknown')
                detection_method = str(pattern.get('DetectionMethod', 'Unknown')).replace('"', '""')
                segment_count = pattern.get('SegmentCount', 0)

                f.write('"{}",'     # Name
                       '"{}",'     # ID
                       '"{}",'     # Imported (Yes/No)
                       '"{}",'     # Detection Method
                       '"{}",'     # Segment Count
                       '"{}"\n'.format(  # Pattern Definition
                    name,
                    pattern_id,
                    imported_status,
                    detection_method,
                    segment_count,
                    segment_def.replace('"', '""')
                ))

        return True
    except Exception as e:
        forms.alert("Error writing CSV file: {}".format(str(e)))
        output.print_md("**CSV Export Error:** {}".format(str(e)))
        return False
    
def main():
    try:
        output.print_md("## Collecting All Line Patterns...")

        # Get all patterns
        patterns = get_all_line_patterns_detailed()

        if not patterns:
            forms.alert("No line patterns found in the project.")
            return

        # Count imported vs built-in
        imported_patterns = [p for p in patterns if p.get('IsImported', False)]
        builtin_patterns = [p for p in patterns if not p.get('IsImported', False)]

        # Display results
        output.print_md("### Found {} Line Patterns:".format(len(patterns)))
        output.print_md("- **Imported/Custom:** {}".format(len(imported_patterns)))
        output.print_md("- **Built-in:** {}\n".format(len(builtin_patterns)))

        for i, pattern in enumerate(patterns, 1):
            imported_status = 'IMPORTED' if pattern.get('IsImported', False) else 'BUILT-IN'
            output.print_md("**{}. {}** [{}]".format(i, pattern.get('Name', 'Unknown'), imported_status))
            output.print_md("- ID: {}".format(pattern.get('Id', 'Unknown')))
            output.print_md("- Detection: {}".format(pattern.get('DetectionMethod', 'Unknown')))
            output.print_md("- Segments: {}".format(pattern.get('SegmentCount', 0)))

            if pattern.get('Segments') and len(pattern['Segments']) > 0:
                output.print_md("- Pattern Definition:")
                for segment in pattern['Segments']:
                    output.print_md("  - {}: {} (Length: {})".format(
                        segment.get('Index', '?'), 
                        segment.get('Type', 'Unknown'), 
                        segment.get('Length', 'N/A')))
            output.print_md("")

        # Get folder name from document title (clean it for file system)
        folder_name = doc.Title
        # Remove invalid characters for folder names
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            folder_name = folder_name.replace(char, '_')

        # Build the full folder path
        base_path = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker"
        folder_path = os.path.join(base_path, folder_name)

        # Create export folder if it doesn't exist
        if not os.path.exists(folder_path):
            try:
                os.makedirs(folder_path)
                output.print_md("**Created folder:** {}".format(folder_path))
            except Exception as e:
                forms.alert("Error creating folder: {}".format(str(e)))
                output.print_md("**Folder Creation Error:** {}".format(str(e)))
                return

        # Generate filename
        filename = "LinePatterns_Audit.csv"
        file_path = os.path.join(folder_path, filename)

        # Export to CSV
        output.print_md("---")
        output.print_md("**Exporting to CSV...**")
        output.print_md("**Target path:** {}".format(file_path))

        if export_to_csv(patterns, file_path):
            output.print_md("### Export Successful!")
            output.print_md("**File saved to:**")
            output.print_md("`{}`".format(file_path))
            output.print_md("**Total patterns exported:** {}".format(len(patterns)))
            output.print_md("- Imported/Custom: {}".format(len(imported_patterns)))
            output.print_md("- Built-in: {}".format(len(builtin_patterns)))

        else:
            output.print_md("### Export Failed!")
            forms.alert("Failed to export file. Check the output window for details.")

    except Exception as e:
        forms.alert("Error: {}".format(str(e)))
        import traceback
        output.print_md("**Error Details:**\n```\n{}\n```".format(traceback.format_exc()))

if __name__ == "__main__":
    main()
