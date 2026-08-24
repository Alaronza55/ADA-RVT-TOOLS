# -*- coding: utf-8 -*-
"""
Select elements in the model whose OPE_NUMBER parameter matches
values listed in a user-provided CSV file.

CSV format: one OPE_NUMBER value per row, no header required.
A header row is auto-detected and skipped if the first cell is non-numeric.

Compatible with: IronPython 2 / PyRevit
"""

__title__ = "Select by\nOPE_NUMBER"
__author__ = "ADA"

import csv
import StringIO

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    ElementId,
)
from pyrevit import revit, forms, script

doc   = revit.doc
uidoc = revit.uidoc
output = script.get_output()

# ---------------------------------------------------------------------------
# 1. Ask the user for a CSV file
# ---------------------------------------------------------------------------
csv_path = forms.pick_file(
    file_ext="csv",
    title="Select CSV file containing OPE_NUMBER values"
)

if not csv_path:
    forms.alert("No file selected. Script cancelled.", exitscript=True)

# ---------------------------------------------------------------------------
# 2. Parse the CSV — no header, just a list of values
# ---------------------------------------------------------------------------
target_numbers = set()

try:
    with open(csv_path, "rb") as f:
        raw = f.read()

    # Strip UTF-8 BOM that Excel adds
    BOM = "\xef\xbb\xbf"
    if raw.startswith(BOM):
        raw = raw[len(BOM):]

    # Normalize line endings
    raw = raw.replace("\r\n", "\n").replace("\r", "\n")

    # Detect delimiter from the first line
    first_line = raw.split("\n")[0] if "\n" in raw else raw
    delimiter  = ";" if first_line.count(";") > first_line.count(",") else ","

    reader = csv.reader(StringIO.StringIO(raw), delimiter=delimiter)

    for row in reader:
        if not row:
            continue
        # Take the first cell of each row
        val = row[0].strip()
        if not val:
            continue
        # Skip a header row if the first cell is non-numeric text
        if val.upper() in ("OPE_NUMBER", "NUMBER", "NO", "NUM", "ID"):
            continue
        target_numbers.add(val)

except Exception as e:
    forms.alert(
        "Error reading CSV file:\n{}".format(str(e)),
        exitscript=True
    )

if not target_numbers:
    forms.alert(
        "The CSV file contains no values. Script cancelled.",
        exitscript=True
    )

# ---------------------------------------------------------------------------
# 3. Collect candidate elements from common BESIX opening categories
# ---------------------------------------------------------------------------
SEARCH_CATEGORIES = [
    BuiltInCategory.OST_GenericModel,
    BuiltInCategory.OST_SpecialityEquipment,
    BuiltInCategory.OST_Windows,
    BuiltInCategory.OST_Doors,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_StructuralFraming,
]

candidate_elements = []

for bic in SEARCH_CATEGORIES:
    try:
        collector = (
            FilteredElementCollector(doc)
            .OfCategory(bic)
            .WhereElementIsNotElementType()
            .ToElements()
        )
        candidate_elements.extend(collector)
    except Exception:
        pass

# ---------------------------------------------------------------------------
# 4. Filter by OPE_NUMBER
# ---------------------------------------------------------------------------
matched_ids           = []
unmatched_csv_numbers = set(target_numbers)

for elem in candidate_elements:
    param = elem.LookupParameter("OPE_NUMBER")
    if param is None:
        continue

    try:
        storage = param.StorageType.ToString()
        if storage == "Integer":
            elem_value = str(param.AsInteger())
        elif storage == "String":
            elem_value = param.AsString() or ""
        elif storage == "Double":
            elem_value = str(int(param.AsDouble()))
        else:
            elem_value = param.AsValueString() or ""
    except Exception:
        elem_value = ""

    elem_value = elem_value.strip()

    if elem_value in target_numbers:
        matched_ids.append(elem.Id)
        unmatched_csv_numbers.discard(elem_value)

# ---------------------------------------------------------------------------
# 5. Select matched elements in Revit
# ---------------------------------------------------------------------------
if not matched_ids:
    forms.alert(
        "No elements found matching the OPE_NUMBER values in the CSV.\n\n"
        "Values searched: {}\n\n"
        "Make sure the elements are in the current model (not linked) "
        "and belong to a supported category.".format(
            ", ".join(sorted(target_numbers)[:20])
            + (" ..." if len(target_numbers) > 20 else "")
        ),
        warn_icon=True
    )
else:
    from System.Collections.Generic import List
    id_list = List[ElementId]([ElementId(eid.IntegerValue) for eid in matched_ids])
    uidoc.Selection.SetElementIds(id_list)

    # ---------------------------------------------------------------------------
    # 6. Report
    # ---------------------------------------------------------------------------
    output.print_md("## Select by OPE_NUMBER -- Results")
    output.print_md(
        "**CSV values loaded:** {}  \n"
        "**Elements selected:** {}".format(
            len(target_numbers),
            len(matched_ids)
        )
    )

    if unmatched_csv_numbers:
        output.print_md(
            "\n**WARNING -- {} value(s) from the CSV had no matching element:**".format(
                len(unmatched_csv_numbers)
            )
        )
        for v in sorted(unmatched_csv_numbers):
            output.print_md("- `{}`".format(v))

    forms.alert(
        "Done!\n\n"
        "{} element(s) selected out of {} OPE_NUMBER value(s) in the CSV.{}".format(
            len(matched_ids),
            len(target_numbers),
            "\n\nSome CSV values had no match -- see the output window for details."
            if unmatched_csv_numbers else ""
        ),
        warn_icon=bool(unmatched_csv_numbers)
    )