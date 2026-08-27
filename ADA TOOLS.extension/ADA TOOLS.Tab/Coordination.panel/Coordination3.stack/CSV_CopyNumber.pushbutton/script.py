# -*- coding: utf-8 -*-
__doc__ = """Copy Generic Model elements from a linked model into the active
document, based on OPE_NUMBER values read from a CSV file.

Pick a CSV file (comma-separated, one or more OPE_NUMBER values per
row or column), then select which loaded link to search. The script
scans the link's Generic Model instances for an OPE_NUMBER parameter
value (instance or type) matching one of the CSV values, copies every
match into the active document at its linked position, and prints a
summary table of what was copied."""
__title__ = "CSV Copy\nNumber"
__author__ = "ADA"

import os
import sys
import codecs

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    RevitLinkInstance,
    BuiltInCategory,
    ElementId,
    CopyPasteOptions,
    ElementTransformUtils,
)
from System.Collections.Generic import List

output = script.get_output()
doc    = revit.doc
uidoc  = revit.uidoc


# =============================================================================
# STEP 1 - CSV selection and parsing
# =============================================================================

def pick_csv_file():
    import clr
    clr.AddReference('System.Windows.Forms')
    from System.Windows.Forms import OpenFileDialog, DialogResult
    dlg = OpenFileDialog()
    dlg.Title = 'Select CSV file with OPE_NUMBER values'
    dlg.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    dlg.Multiselect = False
    if dlg.ShowDialog() == DialogResult.OK:
        return dlg.FileName
    return None


def read_ope_numbers(csv_path):
    values = set()
    for enc in ('utf-8-sig', 'utf-8', 'latin-1'):
        try:
            with codecs.open(csv_path, 'r', enc) as fh:
                for line in fh:
                    line = line.rstrip('\r\n')
                    for cell in line.split(','):
                        cell = cell.strip().strip('"').strip()
                        if cell:
                            values.add(cell)
            break
        except (UnicodeDecodeError, LookupError):
            values = set()
            continue
    return values


# =============================================================================
# STEP 2 - Link-model selection
# =============================================================================

def get_link_instances():
    instances = (
        FilteredElementCollector(doc)
        .OfClass(RevitLinkInstance)
        .ToElements()
    )
    result = []
    for inst in instances:
        link_doc = inst.GetLinkDocument()
        if link_doc is None:
            continue
        name = link_doc.Title or inst.Name
        result.append((name, inst))
    return result


def pick_link_instance(link_pairs):
    name_map = {name: inst for name, inst in link_pairs}
    chosen = forms.SelectFromList.show(
        sorted(name_map.keys()),
        title='Select Linked Model',
        button_name='Select',
        multiselect=False,
    )
    if chosen:
        return name_map[chosen]
    return None


# =============================================================================
# STEP 3 - Find matching Generic Models in the link
# =============================================================================

def get_ope_number(element, link_doc):
    param = element.LookupParameter('OPE_NUMBER')
    if param is None:
        type_id = element.GetTypeId()
        if type_id and type_id != ElementId.InvalidElementId:
            elem_type = link_doc.GetElement(type_id)
            if elem_type:
                param = elem_type.LookupParameter('OPE_NUMBER')
    if param and param.HasValue:
        val = param.AsString()
        if val is None:
            val = param.AsValueString()
        if val:
            return val.strip()
    return None


def find_matching_elements(link_doc, ope_numbers):
    collector = (
        FilteredElementCollector(link_doc)
        .OfCategory(BuiltInCategory.OST_GenericModel)
        .WhereElementIsNotElementType()
    )
    matched_ids = []
    for elem in collector:
        val = get_ope_number(elem, link_doc)
        if val and val in ope_numbers:
            matched_ids.append(elem.Id)
    return matched_ids


# =============================================================================
# STEP 4 - Copy elements into the active document
# =============================================================================

def copy_elements_from_link(link_instance, link_doc, element_ids):
    id_list   = List[ElementId](element_ids)
    transform = link_instance.GetTotalTransform()
    copy_opts = CopyPasteOptions()
    with revit.Transaction('Copy Openings from Link'):
        new_ids = ElementTransformUtils.CopyElements(
            link_doc,
            id_list,
            doc,
            transform,
            copy_opts,
        )
    return list(new_ids)


# =============================================================================
# MAIN
# =============================================================================

def main():
    # 1. CSV
    csv_path = pick_csv_file()
    if not csv_path:
        forms.alert('No CSV file selected. Script cancelled.', exitscript=True)

    ope_numbers = read_ope_numbers(csv_path)
    if not ope_numbers:
        forms.alert(
            'No OPE_NUMBER values could be read from the CSV.\n'
            'Check that the file is not empty and uses comma separation.',
            exitscript=True,
        )

    output.print_md(
        '**CSV loaded** - {} unique OPE_NUMBER value(s) found.'.format(len(ope_numbers))
    )

    # 2. Link selection
    link_pairs = get_link_instances()
    if not link_pairs:
        forms.alert(
            'No loaded Revit link instances found in the active document.',
            exitscript=True,
        )

    link_inst = pick_link_instance(link_pairs)
    if not link_inst:
        forms.alert('No linked model selected. Script cancelled.', exitscript=True)

    link_doc = link_inst.GetLinkDocument()
    output.print_md('**Link selected:** {}'.format(link_doc.Title))

    # 3. Find matches
    matched_ids = find_matching_elements(link_doc, ope_numbers)

    if not matched_ids:
        forms.alert(
            'No Generic Model elements with matching OPE_NUMBER values '
            'were found in the selected link.\n\n'
            'Values searched:\n' + '\n'.join(sorted(ope_numbers)),
            exitscript=True,
        )

    output.print_md(
        '**Matched elements:** {} Generic Model instance(s) found.'.format(len(matched_ids))
    )

    # 4. Confirm and copy
    confirm = forms.alert(
        "{} element(s) will be copied from '{}' into the active document.\n\n"
        'Continue?'.format(len(matched_ids), link_doc.Title),
        ok=True,
        cancel=True,
    )
    if not confirm:
        script.exit()

    new_ids = copy_elements_from_link(link_inst, link_doc, matched_ids)

    # 5. Report
    output.print_md(
        '**Done!** {} element(s) successfully copied into *{}*.'.format(
            len(new_ids), doc.Title
        )
    )

    output.print_md('### Matched OPE_NUMBER values')
    rows = []
    for eid in matched_ids:
        elem = link_doc.GetElement(eid)
        val  = get_ope_number(elem, link_doc)
        rows.append([str(eid.IntegerValue), val or '(unknown)'])

    output.print_table(
        rows,
        title='Elements copied from link',
        columns=['Link Element ID', 'OPE_NUMBER'],
    )


main()