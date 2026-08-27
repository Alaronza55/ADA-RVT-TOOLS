# -*- coding: utf-8 -*-
__doc__ = """Match Pipe Insulation Thickness.

Pick a source element (pipe, pipe fitting or pipe accessory), read its
insulation thickness, then pick one or several target elements and apply the
same thickness to them.

If a target has no insulation yet, it is created using the source's
insulation type."""

__title__ = "Match\nInsulation"
__author__ = "ADA TOOLS"

from pyrevit import revit, DB, forms, script

from Autodesk.Revit.DB import BuiltInCategory, InsulationLiningBase, Transaction
from Autodesk.Revit.DB.Plumbing import PipeInsulation
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException

doc = revit.doc
uidoc = revit.uidoc
logger = script.get_logger()

FT_TO_MM = 304.8

# Categories that can host PipeInsulation.
# Remove OST_PipeAccessory if you only want pipes + fittings.
ALLOWED_BICS = [
    int(BuiltInCategory.OST_PipeCurves),
    int(BuiltInCategory.OST_FlexPipeCurves),
    int(BuiltInCategory.OST_PipeFitting),
    int(BuiltInCategory.OST_PipeAccessory),
]


class PipeInsulationHostFilter(ISelectionFilter):
    """Allow pipes, flex pipes, fittings and accessories."""

    def AllowElement(self, element):
        try:
            return element.Category \
                and element.Category.Id.IntegerValue in ALLOWED_BICS
        except Exception:
            return False

    def AllowReference(self, reference, position):
        return False


def get_insulation(element):
    """Return the first PipeInsulation element hosted by the element, or None."""
    ins_ids = InsulationLiningBase.GetInsulationIds(doc, element.Id)
    if not ins_ids:
        return None
    for ins_id in ins_ids:
        ins = doc.GetElement(ins_id)
        if isinstance(ins, PipeInsulation):
            return ins
    return None


def describe(element):
    cat = element.Category.Name if element.Category else "?"
    return "{} [{}]".format(cat, element.Id.IntegerValue)


def pick_one(prompt):
    ref = uidoc.Selection.PickObject(
        ObjectType.Element, PipeInsulationHostFilter(), prompt
    )
    return doc.GetElement(ref.ElementId)


def pick_many(prompt):
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element, PipeInsulationHostFilter(), prompt
    )
    return [doc.GetElement(r.ElementId) for r in refs]


# --- Source -----------------------------------------------------------------

try:
    source = pick_one("Select the SOURCE pipe / fitting (insulation to copy)")
except OperationCanceledException:
    script.exit()

source_ins = get_insulation(source)
if source_ins is None:
    forms.alert(
        "{} has no insulation.".format(describe(source)),
        title="Match Insulation",
        exitscript=True,
    )

thickness = source_ins.Thickness # internal units (feet)
ins_type_id = source_ins.GetTypeId()
ins_type_name = DB.Element.Name.GetValue(doc.GetElement(ins_type_id))

forms.alert(
    "Source insulation:\n\n"
    "Type: {}\n"
    "Thickness: {:.1f} mm\n\n"
    "Now select the target element(s), then click Finish.".format(
        ins_type_name, thickness * FT_TO_MM
    ),
    title="Match Insulation",
)

# --- Targets ----------------------------------------------------------------

try:
    targets = pick_many("Select the TARGET pipes / fittings - then click Finish")
except OperationCanceledException:
    script.exit()

if not targets:
    script.exit()

updated = 0
created = 0
failed = []

t = Transaction(doc, "Match pipe insulation thickness")
t.Start()
try:
    for elem in targets:
        if elem.Id == source.Id:
            continue
        try:
            existing = get_insulation(elem)
            if existing:
                existing.Thickness = thickness
                updated += 1
            else:
                PipeInsulation.Create(doc, elem.Id, ins_type_id, thickness)
                created += 1
        except Exception as ex:
            failed.append((elem, str(ex)))
            logger.debug("Failed on %s: %s", elem.Id, ex)
    t.Commit()
except Exception:
    t.RollBack()
    raise

# --- Report -----------------------------------------------------------------

msg = "Thickness applied: {:.1f} mm\n\nUpdated: {}\nCreated: {}".format(
    thickness * FT_TO_MM, updated, created
)
if failed:
    msg += "\nFailed: {}".format(len(failed))
    for elem, err in failed:
        logger.warning("%s -> %s", describe(elem), err)

forms.alert(msg, title="Match Insulation")
