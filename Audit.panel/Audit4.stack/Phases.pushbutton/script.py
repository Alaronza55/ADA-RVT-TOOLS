# -*- coding: utf-8 -*-
"""Count Model Elements by Phase
Lists all phases in the project and counts model elements in each phase.
Exports results to CSV file."""

__title__ = "Phases Audit"
__author__ = "Almogi Davidson"

from pyrevit import revit, DB
from pyrevit import script
from collections import defaultdict
import csv
import os

# Get current document
doc = revit.doc
output = script.get_output()

# Define model element categories to count
MODEL_CATEGORIES = [
    # Architectural
    DB.BuiltInCategory.OST_Walls,
    DB.BuiltInCategory.OST_Doors,
    DB.BuiltInCategory.OST_Windows,
    DB.BuiltInCategory.OST_Floors,
    DB.BuiltInCategory.OST_Ceilings,
    DB.BuiltInCategory.OST_Roofs,
    DB.BuiltInCategory.OST_Stairs,
    DB.BuiltInCategory.OST_StairsRailing,
    DB.BuiltInCategory.OST_Ramps,
    DB.BuiltInCategory.OST_Columns,
    DB.BuiltInCategory.OST_CurtainWallPanels,
    DB.BuiltInCategory.OST_CurtainWallMullions,
    DB.BuiltInCategory.OST_CurtaSystem,
    DB.BuiltInCategory.OST_Topography,
    DB.BuiltInCategory.OST_Site,
    DB.BuiltInCategory.OST_Parking,
    DB.BuiltInCategory.OST_Planting,

    # Structural
    DB.BuiltInCategory.OST_StructuralColumns,
    DB.BuiltInCategory.OST_StructuralFraming,
    DB.BuiltInCategory.OST_StructuralFoundation,
    DB.BuiltInCategory.OST_StructConnectionPlates,
    DB.BuiltInCategory.OST_StructConnectionBolts,
    DB.BuiltInCategory.OST_StructConnectionAnchors,
    DB.BuiltInCategory.OST_Rebar,
    DB.BuiltInCategory.OST_AreaRein,
    DB.BuiltInCategory.OST_PathRein,
    DB.BuiltInCategory.OST_FabricReinforcement,
    DB.BuiltInCategory.OST_Truss,

    # MEP - Mechanical
    DB.BuiltInCategory.OST_DuctCurves,
    DB.BuiltInCategory.OST_DuctFitting,
    DB.BuiltInCategory.OST_DuctAccessory,
    DB.BuiltInCategory.OST_DuctTerminal,
    DB.BuiltInCategory.OST_FlexDuctCurves,
    DB.BuiltInCategory.OST_MechanicalEquipment,

    # MEP - Electrical
    DB.BuiltInCategory.OST_CableTray,
    DB.BuiltInCategory.OST_CableTrayFitting,
    DB.BuiltInCategory.OST_Conduit,
    DB.BuiltInCategory.OST_ConduitFitting,
    DB.BuiltInCategory.OST_LightingFixtures,
    DB.BuiltInCategory.OST_ElectricalEquipment,
    DB.BuiltInCategory.OST_ElectricalFixtures,
    DB.BuiltInCategory.OST_DataDevices,
    DB.BuiltInCategory.OST_CommunicationDevices,
    DB.BuiltInCategory.OST_FireAlarmDevices,
    DB.BuiltInCategory.OST_LightingDevices,
    DB.BuiltInCategory.OST_NurseCallDevices,
    DB.BuiltInCategory.OST_SecurityDevices,
    DB.BuiltInCategory.OST_TelephoneDevices,

    # MEP - Plumbing
    DB.BuiltInCategory.OST_PipeCurves,
    DB.BuiltInCategory.OST_PipeFitting,
    DB.BuiltInCategory.OST_PipeAccessory,
    DB.BuiltInCategory.OST_FlexPipeCurves,
    DB.BuiltInCategory.OST_PlumbingFixtures,
    DB.BuiltInCategory.OST_PlumbingEquipment,
    DB.BuiltInCategory.OST_Sprinklers,

    # Other
    DB.BuiltInCategory.OST_Mass,
    DB.BuiltInCategory.OST_GenericModel,
    DB.BuiltInCategory.OST_Furniture,
    DB.BuiltInCategory.OST_FurnitureSystems,
    DB.BuiltInCategory.OST_Casework,
    DB.BuiltInCategory.OST_SpecialityEquipment,
    DB.BuiltInCategory.OST_Entourage,
    DB.BuiltInCategory.OST_Parts,
    DB.BuiltInCategory.OST_Assemblies,
    DB.BuiltInCategory.OST_ShaftOpening,
]

def get_category_name(category_enum):
    """Get readable category name from BuiltInCategory enum"""
    try:
        category = DB.Category.GetCategory(doc, category_enum)
        if category:
            return category.Name
        return str(category_enum).replace("OST_", "")
    except:
        return str(category_enum).replace("OST_", "")

def get_phase_name(phase_id):
    """Get phase name from phase ID"""
    if phase_id == DB.ElementId.InvalidElementId:
        return "None"
    phase = doc.GetElement(phase_id)
    if phase:
        return phase.Name
    return "Unknown"

def get_element_phase_created(element):
    """Get the phase in which the element was created"""
    try:
        phase_created_param = element.get_Parameter(DB.BuiltInParameter.PHASE_CREATED)
        if phase_created_param:
            return phase_created_param.AsElementId()
    except:
        pass
    return DB.ElementId.InvalidElementId

def get_element_phase_demolished(element):
    """Get the phase in which the element was demolished"""
    try:
        phase_demolished_param = element.get_Parameter(DB.BuiltInParameter.PHASE_DEMOLISHED)
        if phase_demolished_param:
            return phase_demolished_param.AsElementId()
    except:
        pass
    return DB.ElementId.InvalidElementId

# Get all phases in the project
phases = DB.FilteredElementCollector(doc)\
    .OfClass(DB.Phase)\
    .ToElements()

if not phases:
    output.print_md("**No phases found in the project.**")
    script.exit()

# Sort phases by sequence number
phases = sorted(phases, key=lambda p: p.get_Parameter(DB.BuiltInParameter.PHASE_SEQUENCE_NUMBER).AsInteger())

# Initialize data structure: phase_id -> category -> count
phase_created_data = defaultdict(lambda: defaultdict(int))
phase_demolished_data = defaultdict(lambda: defaultdict(int))

# Collect elements by category
output.print_md("## Collecting Elements...")
output.print_md("---")

for category_enum in MODEL_CATEGORIES:
    try:
        collector = DB.FilteredElementCollector(doc)\
            .OfCategory(category_enum)\
            .WhereElementIsNotElementType()

        elements = collector.ToElements()
        category_name = get_category_name(category_enum)

        if elements:
            for element in elements:
                # Get phase created
                phase_created_id = get_element_phase_created(element)
                if phase_created_id != DB.ElementId.InvalidElementId:
                    phase_created_data[phase_created_id][category_name] += 1

                # Get phase demolished
                phase_demolished_id = get_element_phase_demolished(element)
                if phase_demolished_id != DB.ElementId.InvalidElementId:
                    phase_demolished_data[phase_demolished_id][category_name] += 1
    except Exception as e:
        pass  # Skip categories that cause errors

# Prepare CSV data
csv_data = []

# Get all unique categories across all phases
all_categories = set()
for phase_id in phase_created_data:
    all_categories.update(phase_created_data[phase_id].keys())
for phase_id in phase_demolished_data:
    all_categories.update(phase_demolished_data[phase_id].keys())

all_categories = sorted(list(all_categories))

# Build CSV rows
for phase in phases:
    phase_id = phase.Id
    phase_name = phase.Name

    # Get all categories that have data for this phase (created or demolished)
    phase_categories = set()
    if phase_id in phase_created_data:
        phase_categories.update(phase_created_data[phase_id].keys())
    if phase_id in phase_demolished_data:
        phase_categories.update(phase_demolished_data[phase_id].keys())

    if not phase_categories:
        # Add a row even if no elements (to show the phase exists)
        csv_data.append({
            'Phase Name': phase_name,
            'Category (Created)': '',
            'Count Created': 0,
            'Category (Demolished)': '',
            'Count Demolished': 0
        })
    else:
        # Sort categories for consistent output
        phase_categories = sorted(list(phase_categories))

        for category in phase_categories:
            count_created = phase_created_data[phase_id].get(category, 0)
            count_demolished = phase_demolished_data[phase_id].get(category, 0)

            csv_data.append({
                'Phase Name': phase_name,
                'Category (Created)': category if count_created > 0 else '',
                'Count Created': count_created,
                'Category (Demolished)': category if count_demolished > 0 else '',
                'Count Demolished': count_demolished
            })

# Define the output folder - MOVED OUTSIDE THE LOOP
folder_name = doc.Title
output_folder = r"C:\Users\adavidson\OneDrive - BESIX\ADA BESIX\Audit Model\TESTING UCB\00 Model Checker\{}".format(folder_name)

# Create the folder if it doesn't exist
if not os.path.exists(output_folder):
    try:
        os.makedirs(output_folder)
        output.print_md("**Created folder:** `{}`".format(output_folder))
    except Exception as e:
        output.print_md("**Error creating folder:** {}".format(str(e)))
        output.print_md("**Attempting to save to default location...**")
        output_folder = os.path.expanduser("~\\Desktop")

# Save detailed element breakdown
filename_detailed = "Phases_Audit.csv"
filepath_detailed = os.path.join(output_folder, filename_detailed)

# Write CSV file - FIXED: Use filepath_detailed instead of filename_detailed
try:
    with open(filepath_detailed, 'wb') as csvfile: 
        fieldnames = ['Phase Name', 'Category (Created)', 'Count Created', 
                      'Category (Demolished)', 'Count Demolished']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames, lineterminator='\n')

        writer.writeheader()
        for row in csv_data:
            # Encode strings to handle special characters
            encoded_row = {}
            for key, value in row.items():
                if isinstance(value, unicode):
                    encoded_row[key] = value.encode('utf-8')
                else:
                    encoded_row[key] = value
            writer.writerow(encoded_row)

    output.print_md("## CSV Export Successful!")
    output.print_md("**File saved to:** `{}`".format(filepath_detailed))
    output.print_md("---\n")
except Exception as e:
    output.print_md("## Error Exporting CSV")
    output.print_md("**Error:** {}".format(str(e)))
    output.print_md("**Attempted path:** `{}`".format(filepath_detailed))
    output.print_md("---\n")

# Print results to output window
output.print_md("# Model Elements Count by Phase")
output.print_md("---\n")

# Print summary for each phase
for phase in phases:
    phase_id = phase.Id
    phase_name = phase.Name

    output.print_md("## Phase: **{}**".format(phase_name))

    # Elements created in this phase
    if phase_id in phase_created_data:
        output.print_md("### Elements Created:")

        # Sort categories by count (descending)
        sorted_categories = sorted(
            phase_created_data[phase_id].items(),
            key=lambda x: x[1],
            reverse=True
        )

        total_created = sum(phase_created_data[phase_id].values())

        # Create table
        output.print_md("| Category | Count |")
        output.print_md("|----------|------:|")

        for category_name, count in sorted_categories:
            output.print_md("| {} | {} |".format(category_name, count))

        output.print_md("| **TOTAL** | **{}** |".format(total_created))
    else:
        output.print_md("### Elements Created: *None*")

    output.print_md("")

    # Elements demolished in this phase
    if phase_id in phase_demolished_data:
        output.print_md("### Elements Demolished:")

        sorted_categories = sorted(
            phase_demolished_data[phase_id].items(),
            key=lambda x: x[1],
            reverse=True
        )

        total_demolished = sum(phase_demolished_data[phase_id].values())

        output.print_md("| Category | Count |")
        output.print_md("|----------|------:|")

        for category_name, count in sorted_categories:
            output.print_md("| {} | {} |".format(category_name, count))

        output.print_md("| **TOTAL** | **{}** |".format(total_demolished))
    else:
        output.print_md("### Elements Demolished: *None*")

    output.print_md("---\n")

# Print overall summary
output.print_md("## Overall Summary")
output.print_md("| Phase | Created | Demolished |")
output.print_md("|-------|--------:|-----------:|")

for phase in phases:
    phase_id = phase.Id
    phase_name = phase.Name
    created_count = sum(phase_created_data[phase_id].values()) if phase_id in phase_created_data else 0
    demolished_count = sum(phase_demolished_data[phase_id].values()) if phase_id in phase_demolished_data else 0
    output.print_md("| {} | {} | {} |".format(phase_name, created_count, demolished_count))

output.print_md("\n---")
output.print_md("**Script completed successfully!**")
output.print_md("**CSV file location:** `{}`".format(filepath_detailed))
