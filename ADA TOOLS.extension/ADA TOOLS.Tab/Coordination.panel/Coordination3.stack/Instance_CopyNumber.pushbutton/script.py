# -*- coding: utf-8 -*-
__doc__ = """Copy Generic Model elements from a linked model into the active
document, based on a single OPE_NUMBER value you type in.

Type an OPE_NUMBER, pick which loaded link to search, and every
Generic Model in that link whose OPE_NUMBER matches (instance or
type parameter) is copied into the active document at its linked
position. Same idea as "CSV Copy Number", but for one value at a
time instead of a batch from a file."""
__title__ = "Instance Copy\nNumber"
__author__ = "ADA"

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
# STEP 1 - Ask user to type an OPE_NUMBER
# =============================================================================

def ask_ope_number():
    value = forms.ask_for_string(
        prompt='Enter the OPE_NUMBER to search for:',
        title='OPE_NUMBER Input',
        default='',
    )
    if value is not None:
        return value.strip()
    return None


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


def find_matching_elements(link_doc, ope_number):
    collector = (
        FilteredElementCollector(link_doc)
        .OfCategory(BuiltInCategory.OST_GenericModel)
        .WhereElementIsNotElementType()
    )
    matched_ids = []
    for elem in collector:
        val = get_ope_number(elem, link_doc)
        if val and val == ope_number:
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
    # 1. Ask for OPE_NUMBER
    ope_number = ask_ope_number()
    if not ope_number:
        forms.alert('No OPE_NUMBER entered. Script cancelled.', exitscript=True)

    output.print_md('**OPE_NUMBER entered:** {}'.format(ope_number))

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
    matched_ids = find_matching_elements(link_doc, ope_number)

    if not matched_ids:
        forms.alert(
            'No Generic Model elements with OPE_NUMBER "{}" '
            'were found in the selected link.'.format(ope_number),
            exitscript=True,
        )

    output.print_md(
        '**Matched elements:** {} Generic Model instance(s) found.'.format(len(matched_ids))
    )

    # 4. Confirm and copy
    confirm = forms.alert(
        "{} element(s) with OPE_NUMBER '{}' will be copied "
        "from '{}' into the active document.\n\n"
        'Continue?'.format(len(matched_ids), ope_number, link_doc.Title),
        ok=True,
        cancel=True,
    )
    if not confirm:
        script.exit()

    new_ids = copy_elements_from_link(link_inst, link_doc, matched_ids)

    # 5. Report
    output.print_md(
        '**Done!** {} element(s) with OPE_NUMBER "{}" copied into *{}*.'.format(
            len(new_ids), ope_number, doc.Title
        )
    )

    rows = [[str(eid.IntegerValue), ope_number] for eid in matched_ids]
    output.print_table(
        rows,
        title='Elements copied from link',
        columns=['Link Element ID', 'OPE_NUMBER'],
    )


main()