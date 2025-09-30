"""Export Imported Line Patterns
Exports all imported line patterns from the project to a text file.

"""
__title__ = "Export Imported\nLine Patterns"
__author__ = "Your Name"

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

import os
from pyrevit import forms, script

# Get current document
doc = __revit__.ActiveUIDocument.Document
output = script.get_output()

def get_imported_line_patterns():
    """Get all imported line patterns from the project"""
    imported_patterns = []
    
    # Get all line pattern elements
    line_pattern_collector = FilteredElementCollector(doc).OfClass(LinePatternElement)
    
    # Default/built-in patterns that come with Revit (these are typically not imported)
    builtin_patterns = {
        "Solid", "Dash", "Dot", "Dash dot", "Dash dot dot", "Long dash", 
        "Long dash dot", "Long dash dot dot", "Dash short dash", "Dash space dash",
        "<Solid fill>", "By Category"
    }
    
    for pattern_elem in line_pattern_collector:
        pattern_name = pattern_elem.Name
        
        # Filter out built-in patterns - imported patterns usually have different naming
        if pattern_name not in builtin_patterns:
            line_pattern = pattern_elem.GetLinePattern()
            
            pattern_info = {
                'Name': pattern_name,
                'Id': pattern_elem.Id.IntegerValue,
                'Type': 'Custom/Imported',  # Simplified since GetPatternType doesn't exist
                'Segments': [],
                'IsBuiltIn': False
            }
            
            # Get pattern segments (dashes, dots, spaces)
            try:
                segments = line_pattern.GetSegments()
                if segments:
                    for segment in segments:
                        segment_info = {
                            'Type': str(segment.Type),
                            'Length': segment.Length
                        }
                        pattern_info['Segments'].append(segment_info)
                else:
                    # If no segments, it's likely a solid pattern
                    pattern_info['Segments'] = [{'Type': 'Solid', 'Length': 0}]
            except:
                pattern_info['Segments'] = [{'Type': 'Unknown', 'Length': 'Unknown'}]
            
            imported_patterns.append(pattern_info)
    
    return imported_patterns

def get_all_line_patterns_detailed():
    """Alternative method to get more detailed pattern information"""
    all_patterns = []
    
    # Get all line pattern elements
    line_pattern_collector = FilteredElementCollector(doc).OfClass(LinePatternElement)
    
    for pattern_elem in line_pattern_collector:
        try:
            pattern_name = pattern_elem.Name
            line_pattern = pattern_elem.GetLinePattern()
            
            pattern_info = {
                'Name': pattern_name,
                'Id': pattern_elem.Id.IntegerValue,
                'ElementId': pattern_elem.Id,
                'Segments': [],
                'SegmentCount': 0,
                'IsCustom': True  # We'll determine this based on naming/segments
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
            
            # Determine if it's likely a built-in pattern
            builtin_names = ["Solid", "Dash", "Dot", "Dash dot", "Dash dot dot", 
                           "Long dash", "Long dash dot", "Long dash dot dot", 
                           "Dash short dash", "Dash space dash", "<Solid fill>"]
            
            pattern_info['IsCustom'] = pattern_name not in builtin_names
            
            all_patterns.append(pattern_info)
            
        except Exception as e:
            # If we can't process a pattern, still add basic info
            pattern_info = {
                'Name': pattern_elem.Name if hasattr(pattern_elem, 'Name') else 'Unknown',
                'Id': pattern_elem.Id.IntegerValue,
                'ElementId': pattern_elem.Id,
                'Segments': [{'Index': 1, 'Type': 'Error reading pattern', 'Length': str(e)}],
                'SegmentCount': 0,
                'IsCustom': True
            }
            all_patterns.append(pattern_info)
    
    return all_patterns

def export_to_file(patterns, file_path):
    """Export line patterns to a text file"""
    try:
        with open(file_path, 'w') as f:
            f.write("Line Patterns Export\n")
            f.write("=" * 50 + "\n")
            f.write("Project: {}\n".format(doc.Title))
            f.write("Total Line Patterns: {}\n\n".format(len(patterns)))
            
            for i, pattern in enumerate(patterns, 1):
                f.write("{}. {}\n".format(i, pattern['Name']))
                f.write("   ID: {}\n".format(pattern['Id']))
                f.write("   Type: {}\n".format('Custom/Imported' if pattern.get('IsCustom', True) else 'Built-in'))
                f.write("   Segment Count: {}\n".format(pattern.get('SegmentCount', len(pattern['Segments']))))
                
                if pattern['Segments']:
                    f.write("   Pattern Definition:\n")
                    for segment in pattern['Segments']:
                        f.write("     Segment {}: {} - Length: {}\n".format(
                            segment.get('Index', '?'),
                            segment['Type'], 
                            segment.get('Length', 'N/A')))
                f.write("\n")
        
        return True
    except Exception as e:
        forms.alert("Error writing file: {}".format(str(e)))
        return False

def export_to_csv(patterns, file_path):
    """Export line patterns to a CSV file"""
    try:
        with open(file_path, 'w') as f:
            f.write("Name,ID,Type,Segment_Count,Pattern_Definition\n")
            
            for pattern in patterns:
                segment_def = ""
                if pattern['Segments']:
                    segment_parts = []
                    for segment in pattern['Segments']:
                        segment_parts.append("{}:{}".format(segment['Type'], segment.get('Length', 'N/A')))
                    segment_def = " | ".join(segment_parts)
                else:
                    segment_def = "No segments"
                
                # Clean the data for CSV
                name = str(pattern['Name']).replace('"', '""')
                type_str = 'Custom/Imported' if pattern.get('IsCustom', True) else 'Built-in'
                
                f.write('"{}",'     # Name
                       '"{}",'     # ID
                       '"{}",'     # Type
                       '"{}",'     # Segment Count
                       '"{}"\n'.format(  # Pattern Definition
                    name,
                    pattern['Id'],
                    type_str,
                    pattern.get('SegmentCount', len(pattern['Segments'])),
                    segment_def.replace('"', '""')
                ))
        
        return True
    except Exception as e:
        forms.alert("Error writing CSV file: {}".format(str(e)))
        return False

def main():
    # Ask user which patterns to show
    filter_choice = forms.SelectFromList.show(
        ['Custom/Imported Patterns Only', 'All Line Patterns'],
        title='Pattern Filter',
        msg='Which patterns would you like to export?'
    )
    
    if not filter_choice:
        return
    
    output.print_md("## Collecting Line Patterns...")
    
    # Get all patterns first
    all_patterns = get_all_line_patterns_detailed()
    
    # Filter based on user choice
    if filter_choice == 'Custom/Imported Patterns Only':
        patterns = [p for p in all_patterns if p.get('IsCustom', True)]
        title = "Custom/Imported Line Patterns"
    else:
        patterns = all_patterns
        title = "All Line Patterns"
    
    if not patterns:
        forms.alert("No line patterns found matching your criteria.")
        return
    
    # Display results
    output.print_md("### Found {} {}:".format(len(patterns), title))
    
    for i, pattern in enumerate(patterns, 1):
        pattern_type = 'Custom/Imported' if pattern.get('IsCustom', True) else 'Built-in'
        output.print_md("**{}. {}** ({})".format(i, pattern['Name'], pattern_type))
        output.print_md("- ID: {}".format(pattern['Id']))
        output.print_md("- Segments: {}".format(pattern.get('SegmentCount', len(pattern['Segments']))))
        
        if pattern['Segments'] and len(pattern['Segments']) > 0:
            output.print_md("- Pattern Definition:")
            for segment in pattern['Segments']:
                output.print_md("  - {}: {} (Length: {})".format(
                    segment.get('Index', '?'), 
                    segment['Type'], 
                    segment.get('Length', 'N/A')))
        output.print_md("")
    
    # Export options
    export_choice = forms.SelectFromList.show(
        ['Text File (.txt)', 'CSV File (.csv)', 'Both', 'No Export'],
        title='Export Options',
        msg='Choose export format:'
    )
    
    if export_choice and export_choice != 'No Export':
        from System.Windows.Forms import SaveFileDialog, DialogResult
        
        save_dialog = SaveFileDialog()
        save_dialog.Title = "Save Line Patterns Export"
        save_dialog.DefaultExt = "txt"
        save_dialog.FileName = "{}_LinePatterns".format(doc.Title.replace(" ", "_"))
        
        if export_choice == 'Text File (.txt)':
            save_dialog.Filter = "Text Files (*.txt)|*.txt"
            if save_dialog.ShowDialog() == DialogResult.OK:
                if export_to_file(patterns, save_dialog.FileName):
                    forms.alert("Export completed successfully!\nFile saved to: {}".format(save_dialog.FileName))
                    
        elif export_choice == 'CSV File (.csv)':
            save_dialog.Filter = "CSV Files (*.csv)|*.csv"
            save_dialog.DefaultExt = "csv"
            if save_dialog.ShowDialog() == DialogResult.OK:
                if export_to_csv(patterns, save_dialog.FileName):
                    forms.alert("Export completed successfully!\nFile saved to: {}".format(save_dialog.FileName))
                    
        elif export_choice == 'Both':
            save_dialog.Filter = "All Files (*.*)|*.*"
            if save_dialog.ShowDialog() == DialogResult.OK:
                base_path = save_dialog.FileName.rsplit('.', 1)[0]
                txt_path = base_path + ".txt"
                csv_path = base_path + ".csv"
                
                txt_success = export_to_file(patterns, txt_path)
                csv_success = export_to_csv(patterns, csv_path)
                
                if txt_success and csv_success:
                    forms.alert("Export completed successfully!\nFiles saved to:\n- {}\n- {}".format(txt_path, csv_path))
                elif txt_success or csv_success:
                    forms.alert("Partial export completed. Check the output location.")
                else:
                    forms.alert("Export failed!")

if __name__ == "__main__":
    main()
