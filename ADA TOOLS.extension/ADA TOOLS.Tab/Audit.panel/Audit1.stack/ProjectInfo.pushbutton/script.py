# -*- coding: utf-8 -*-
"""Count Views and Placed Views
This script counts total views in the project and views placed on sheets.
"""

__title__ = "Project Info"
__author__ = "Almog Davidson"
__doc__ = """General Audit Information about the current Revit project."""

from pyrevit import revit, DB, script, forms
import os
from datetime import datetime
import re
import csv

# Get the current document
doc = revit.doc

folder_name = doc.Title

def format_file_size(size_bytes):
    """Convert file size from bytes to human-readable format"""
    if size_bytes == 0:
        return "0 B"
    
    size_names = ["B", "KB", "MB", "GB", "TB"]
    import math
    i = int(math.floor(math.log(size_bytes, 1024)))
    p = math.pow(1024, i)
    s = round(size_bytes / p, 2)
    return "{} {}".format(s, size_names[i])

def project_info():
    # Store results_info in a list instead of printing directly
    results_info = []

    results_info.append("=== PROJECT FILE INFORMATION ===")

    # Document Title
    results_info.append("Document Title: {}".format(doc.Title))

    # Check if file has been saved
    if doc.PathName:
        # Full path
        results_info.append("Full Path: {}".format(doc.PathName))

        # File name only
        file_name = os.path.basename(doc.PathName)
        results_info.append("File Name: {}".format(file_name))

        # File size
        try:
            file_size = os.path.getsize(doc.PathName)
            file_size_formatted = format_file_size(file_size)
            results_info.append("File Size: {} ({:,} bytes)".format(file_size_formatted, file_size))
        except Exception as e:
            results_info.append("File Size: Could not determine file size - {}".format(str(e)))

        # Check if workshared
        if doc.IsWorkshared:
            results_info.append("Is Workshared: Yes")
            try:
                # Get the central model path
                central_path = doc.GetWorksharingCentralModelPath()
                if central_path:
                    # Convert ModelPath to user visible path
                    central_path_string = central_path.CentralPath
                    if central_path_string:
                        central_file_name = os.path.basename(central_path_string)
                        results_info.append("Central Model File Name: {}".format(central_file_name))
                        
                        # Try to get central file size
                        try:
                            if os.path.exists(central_path_string):
                                central_file_size = os.path.getsize(central_path_string)
                                central_size_formatted = format_file_size(central_file_size)
                                results_info.append("Central File Size: {} ({:,} bytes)".format(central_size_formatted, central_file_size))
                            else:
                                results_info.append("Central File Size: Central file not accessible")
                        except Exception as e:
                            results_info.append("Central File Size: Could not determine - {}".format(str(e)))
                    else:
                        results_info.append("Central Model Path: {}".format(str(central_path)))
            except:
                results_info.append("Could not retrieve central model path")
        else:
            results_info.append("Is Workshared: No")

    else:
        results_info.append("File has not been saved yet - using document title: {}".format(doc.Title))
        results_info.append("File Size: N/A (file not saved)")

    # Print to PyRevit output window
    for result in results_info:
        print("{}".format(result))

    return results_info

def save_to_csv(results_info):
    """Save the audit results to a CSV file"""
    import csv

    # Define the output folder - change this path as needed
    output_folder = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker\{}".format(folder_name)

    # Create the folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Clean document title for filename (remove invalid characters)
    doc_title = doc.Title or "Unknown"
    clean_title = re.sub(r'[<>:"/\\|?*]', '_', doc_title)

    filename = "Audit_Project_Information.csv"
    filepath = os.path.join(output_folder, filename)

    try:
        # Open file without newline parameter for Python 2.7 compatibility
        with open(filepath, 'wb') as csvfile:
            writer = csv.writer(csvfile)

            # Write timestamp header
            writer.writerow(["Project Information Audit Report"])
            writer.writerow(["Generated on: {}".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))])
            writer.writerow(["Document: {}".format(doc.Title or "Unknown")])
            writer.writerow([])  # Empty row for spacing

            # Write column headers
            writer.writerow(["Property", "Value"])

            # Process and write data rows
            for result in results_info:
                if result.startswith("==="):
                    # Skip header lines or treat as section headers
                    section_name = result.replace("===", "").strip()
                    writer.writerow(["Property", section_name])
                elif ":" in result:
                    # Split property: value pairs
                    parts = result.split(":", 1)  # Split only on first colon
                    if len(parts) == 2:
                        property_name = parts[0].strip()
                        value = parts[1].strip()
                        writer.writerow([property_name, value])
                else:
                    # Handle any other format
                    writer.writerow(["Central Model Path", result])

        print("CSV report saved to: {}".format(filepath))
        return filepath

    except Exception as e:
        print("Error saving CSV file: {}".format(str(e)))
        return None

if __name__ == '__main__':
    # Run the audit and get results_info
    results_info = project_info()

    # Save results_info to CSV
    save_to_csv(results_info)
