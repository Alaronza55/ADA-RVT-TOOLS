# -*- coding: utf-8 -*-
__title__ = "Set Blockwork\nUnique Code"
__author__ = "BESIX"
__doc__ = """Sets B6_Wall_Unique_Code to 'MAS-(BaseConstraint[0:2])-(NNN)'
for all Blockwork walls in the active view that don't have the value set yet.
Walls already carrying a code are left untouched.
Numbering starts at 001 from the wall with the highest X centroid coordinate."""

from pyrevit import revit, DB, script, forms

doc         = revit.doc
active_view = doc.ActiveView
output      = script.get_output()

# --- Collect all walls visible in the active view ---
collector = DB.FilteredElementCollector(doc, active_view.Id) \
              .OfCategory(DB.BuiltInCategory.OST_Walls) \
              .WhereElementIsNotElementType()

# --- Filter walls whose family or type name contains "Blockwork" ---
blockwork_walls = []
for wall in collector:
    wall_type = doc.GetElement(wall.GetTypeId())
    if wall_type is None:
        continue

    family_name = wall_type.FamilyName or ""
    type_name   = wall_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
    type_name   = type_name.AsString() if type_name else ""

    if "blockwork" in family_name.lower() or "blockwork" in type_name.lower():
        blockwork_walls.append(wall)

if not blockwork_walls:
    forms.alert("No Blockwork walls found in the active view.", exitscript=True)

# --- Split into already-coded and needs-coding ---
already_coded = []
needs_code    = []

for wall in blockwork_walls:
    p = wall.LookupParameter("B6_Wall_Unique_Code")
    if p and p.AsString():
        already_coded.append(wall)
    else:
        needs_code.append(wall)

# --- Collect all numbers already in use across ALL coded walls ---
existing_numbers = set()
for wall in already_coded:
    p = wall.LookupParameter("B6_Wall_Unique_Code")
    parts = p.AsString().split("-")
    if len(parts) == 3:
        try:
            existing_numbers.add(int(parts[2]))
        except ValueError:
            pass

# --- Compute centroid X and sort needs_code by highest X first ---
def get_centroid_x(wall):
    bbox = wall.get_BoundingBox(active_view)
    if bbox is None:
        bbox = wall.get_BoundingBox(None)
    if bbox is None:
        return 0.0
    return (bbox.Min.X + bbox.Max.X) / 2.0

needs_code.sort(key=get_centroid_x, reverse=True)   # highest X → first

# --- Build a sequence that skips already-used numbers ---
def number_generator(existing):
    n = 1
    while True:
        if n not in existing:
            yield n
        n += 1

counter = number_generator(existing_numbers)

# --- Report already-coded walls ---
for wall in already_coded:
    p     = wall.LookupParameter("B6_Wall_Unique_Code")
    value = p.AsString()
    output.print_md("— Wall **{}** already has `{}` – left unchanged.".format(
        wall.Id.IntegerValue, value))

# --- Write parameter values inside a single transaction ---
processed        = 0
skipped_missing  = 0
skipped_readonly = 0

with revit.Transaction("Set B6_Wall_Unique_Code"):
    for wall in needs_code:
        element_id = wall.Id.IntegerValue

        target_param = wall.LookupParameter("B6_Wall_Unique_Code")
        if target_param is None:
            output.print_md("⚠ Wall **{}** is missing parameter 'B6_Wall_Unique_Code' – skipped.".format(element_id))
            skipped_missing += 1
            continue

        if target_param.IsReadOnly:
            output.print_md("⚠ Parameter on wall **{}** is read-only – skipped.".format(element_id))
            skipped_readonly += 1
            continue

        # Base Constraint → first 2 characters of level name
        base_param = wall.get_Parameter(DB.BuiltInParameter.WALL_BASE_CONSTRAINT)
        if base_param is None:
            output.print_md("⚠ Wall **{}** has no Base Constraint – skipped.".format(element_id))
            skipped_missing += 1
            continue

        level = doc.GetElement(base_param.AsElementId())
        if level is None or not level.Name:
            output.print_md("⚠ Wall **{}** Base Constraint resolves to no level – skipped.".format(element_id))
            skipped_missing += 1
            continue

        level_prefix = level.Name[:2]
        number       = next(counter)
        code_value   = "MAS-{}-{:03d}".format(level_prefix, number)

        target_param.Set(code_value)
        output.print_md("✔ Wall **{}** → `{}`".format(element_id, code_value))
        processed += 1

# --- Summary ---
output.print_md("---")
output.print_md("**Done.** {} newly coded, {} already set (unchanged), {} missing param, {} read-only.".format(
    processed, len(already_coded), skipped_missing, skipped_readonly))