# -*- coding: utf-8 -*-
"""Helpers shared by the Planning panel tools (Import Planning, Fill from
Planning): phase-name cleaning, Excel file / sheet picking and the
creation of text shared parameters in the Phasing group."""

import os

from Autodesk.Revit.DB import CategoryType, ExternalDefinitionCreationOptions

from pyrevit import forms, script

from GUI.forms import select_from_dict
from Snippets._xlsx import XlsxWorkbook

# characters Revit does not accept in names
FORBIDDEN_CHARS = u'\\:{}[]|;<>?`~'

ADA_DATA_DIR = os.path.join(os.environ.get("APPDATA", ""), "ADA Tools")
SHARED_PARAM_FILE = os.path.join(ADA_DATA_DIR, "ADA_SharedParameters.txt")
SHARED_PARAM_GROUP = "ADA Planning"


def clean_name(text):
    """One-line name Revit accepts: line breaks -> ' - ' (leading list
    dashes dropped), forbidden characters removed, spaces squeezed. Used
    for phase names so that Excel rows and Revit phases match."""
    lines = [line.strip().lstrip(u"-•*").strip()
             for line in (text or u"").splitlines()]
    name = u" - ".join(line for line in lines if line)
    name = u"".join(ch for ch in name if ch not in FORBIDDEN_CHARS)
    return u" ".join(name.split())


def name_key(name):
    """Case / spacing insensitive key to match a planning row to a phase."""
    return u" ".join((name or u"").split()).lower()


def eid_value(eid):
    """ElementId numeric value, 2024+ safe."""
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

def pick_workbook(config, title):
    """File dialog (starts in the folder of the last file used) - returns an
    XlsxWorkbook, exits the script on cancel / unreadable file."""
    path = forms.pick_file(files_filter="Excel workbook (*.xlsx;*.xlsm)|"
                                        "*.xlsx;*.xlsm",
                           init_dir=os.path.dirname(
                               config.get_option("last_file", "")),
                           title="Select the planning (Excel)")
    if not path:
        script.exit()
    try:
        return XlsxWorkbook(path)
    except Exception as ex:
        forms.alert("Could not read the Excel file:\n{}\n\nOnly .xlsx / .xlsm "
                    "files are supported (save .xls files as .xlsx)."
                    .format(ex), title=title, exitscript=True)


def pick_sheet(workbook, title, version):
    names = workbook.sheet_names
    if not names:
        forms.alert("The workbook has no sheets.", title=title, exitscript=True)
    if len(names) == 1:
        return names[0]
    chosen = select_from_dict(
        {name: name for name in names},
        title=title,
        label="Sheet containing the planning:",
        button_name="Next",
        version=version,
        SelectMultiple=False)
    if not chosen:
        script.exit()
    return chosen[0]


# ---------------------------------------------------------------------------
# shared parameters
# ---------------------------------------------------------------------------

def bound_definition(doc, name):
    """Definition of the project / shared parameter bound under `name`, or
    None."""
    iterator = doc.ParameterBindings.ForwardIterator()
    while iterator.MoveNext():
        if iterator.Key.Name == name:
            return iterator.Key
    return None


def _external_definition(app, name, description):
    """Text definition `name` from the ADA shared parameter file (created if
    needed). The user's shared parameter file setting is restored after."""
    if not os.path.isdir(ADA_DATA_DIR):
        os.makedirs(ADA_DATA_DIR)
    if not os.path.isfile(SHARED_PARAM_FILE):
        open(SHARED_PARAM_FILE, "w").close()

    original = app.SharedParametersFilename
    try:
        app.SharedParametersFilename = SHARED_PARAM_FILE
        def_file = app.OpenSharedParameterFile()
        group = (def_file.Groups.get_Item(SHARED_PARAM_GROUP)
                 or def_file.Groups.Create(SHARED_PARAM_GROUP))
        definition = group.Definitions.get_Item(name)
        if definition is None:
            try:
                from Autodesk.Revit.DB import SpecTypeId   # Revit 2022+
                options = ExternalDefinitionCreationOptions(
                    name, SpecTypeId.String.Text)
            except Exception:
                from Autodesk.Revit.DB import ParameterType
                options = ExternalDefinitionCreationOptions(
                    name, ParameterType.Text)
            options.Description = description
            definition = group.Definitions.Create(options)
        return definition
    finally:
        try:
            app.SharedParametersFilename = original or ""
        except Exception:
            pass


def ensure_shared_text_parameter(doc, name, description):
    """Make sure an instance parameter `name` is bound to every model
    category, in the Phasing group - creating a text shared parameter if it
    does not exist yet. Must run inside a transaction. Returns True if it
    was created, False if it already existed."""
    if bound_definition(doc, name) is not None:
        return False

    app = doc.Application
    definition = _external_definition(app, name, description)
    categories = app.Create.NewCategorySet()
    for category in doc.Settings.Categories:
        if (category.CategoryType == CategoryType.Model
                and category.AllowsBoundParameters):
            categories.Insert(category)
    binding = app.Create.NewInstanceBinding(categories)
    try:
        from Autodesk.Revit.DB import GroupTypeId   # Revit 2022+
        inserted = doc.ParameterBindings.Insert(definition, binding,
                                                GroupTypeId.Phasing)
    except Exception:
        from Autodesk.Revit.DB import BuiltInParameterGroup
        inserted = doc.ParameterBindings.Insert(
            definition, binding, BuiltInParameterGroup.PG_PHASING)
    if not inserted:
        raise Exception("Revit refused to add the '{}' parameter".format(name))
    return True
