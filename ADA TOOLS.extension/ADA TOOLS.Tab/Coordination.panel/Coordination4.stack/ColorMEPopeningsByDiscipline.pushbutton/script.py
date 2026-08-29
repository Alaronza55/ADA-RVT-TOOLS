# -*- coding: utf-8 -*-
__title__ = "Color Tags\nby Discipline"
__doc__ = "Colors generic model tags in the active view based on the host's OPE_DISCIPLINE parameter."

import System
from pyrevit import revit, DB, script

# Shared ADA-Tools dark/gold themed report (see lib/GUI/ReportTheme.py)
from GUI.ReportTheme import ADAReport

output = script.get_output()

# ── Color map ─────────────────────────────────────────────────────────────────
DISCIPLINE_COLORS = {
    "ELE": DB.Color(0, 29, 222),   # Blue
    "EHS": DB.Color(255, 0, 0),   # Red
    "BLU": DB.Color(255, 0, 255),   # Purple
    "HVC": DB.Color(0, 128, 0),   # Green
    "CLU": DB.Color(255, 128, 0),   # Orange
    "DRS": DB.Color(0, 255, 255),   # Turquoise
    "CRD": DB.Color(0, 255, 255),   # Turquoise
}

# ── Helpers ───────────────────────────────────────────────────────────────────
def get_solid_fill_id(doc):
    collector = DB.FilteredElementCollector(doc)\
                  .OfClass(DB.FillPatternElement)
    for fp in collector:
        pattern = fp.GetFillPattern()
        if pattern.Target == DB.FillPatternTarget.Drafting and pattern.IsSolidFill:
            return fp.Id
    return DB.ElementId.InvalidElementId

def make_override(color, solid_id):
    ogs = DB.OverrideGraphicSettings()
    try:
        if solid_id != DB.ElementId.InvalidElementId:
            ogs.SetProjectionFillPatternId(solid_id)
            ogs.SetProjectionFillColor(color)
    except System.MissingMemberException:
        pass
    except Exception:
        pass
    ogs.SetProjectionLineColor(color)
    return ogs

def get_param_value(element, param_name):
    param = element.LookupParameter(param_name)
    if param and param.HasValue:
        return param.AsString() or param.AsValueString()
    return None

def get_tagged_element(tag, doc):
    try:
        tagged_refs = tag.GetTaggedReferences()
    except Exception:
        tagged_refs = None

    if tagged_refs:
        ref = tagged_refs[0]
        try:
            # Extract raw integers to avoid IronPython 2 overload ambiguity
            link_id_int = ref.LinkedElementId.IntegerValue
            elem_id_int = ref.ElementId.IntegerValue

            if link_id_int != -1:
                # Linked element
                link_inst = doc.GetElement(DB.ElementId(link_id_int))
                if link_inst and hasattr(link_inst, "GetLinkDocument"):
                    link_doc = link_inst.GetLinkDocument()
                    if link_doc:
                        elem = link_doc.GetElement(DB.ElementId(elem_id_int))
                        return elem, link_doc
                return None, None
            else:
                # Local element
                elem = doc.GetElement(DB.ElementId(elem_id_int))
                return elem, doc
        except Exception:
            pass

    return None, None


# ── Main ──────────────────────────────────────────────────────────────────────
doc  = revit.doc
view = doc.ActiveView

solid_id = get_solid_fill_id(doc)

tags = (DB.FilteredElementCollector(doc, view.Id)
          .OfClass(DB.IndependentTag)
          .ToElements())

colored = 0
skipped = 0
skipped_details = []
errors  = []

with revit.Transaction("Color Tags by OPE_DISCIPLINE"):
    for tag in tags:
        try:
            elem, source_doc = get_tagged_element(tag, doc)
            if elem is None:
                skipped += 1
                skipped_details.append(u"Tag {} — could not resolve host element".format(
                    tag.Id.IntegerValue))
                continue

            discipline = get_param_value(elem, "OPE_DISCIPLINE")
            if not discipline:
                skipped += 1
                skipped_details.append(u"Tag {} — host {} has no OPE_DISCIPLINE".format(
                    tag.Id.IntegerValue, elem.Id.IntegerValue))
                continue

            color = DISCIPLINE_COLORS.get(discipline.strip().upper())
            if color is None:
                skipped += 1
                skipped_details.append(u"Tag {} — unknown discipline '{}'".format(
                    tag.Id.IntegerValue, discipline))
                continue

            view.SetElementOverrides(tag.Id, make_override(color, solid_id))
            colored += 1

        except Exception as ex:
            errors.append(u"Tag {}: {}".format(tag.Id.IntegerValue, str(ex)))

# ── Report ────────────────────────────────────────────────────────────────────
report = ADAReport(__title__.replace(chr(10), " "))
report.success(u"Colored: <b>{}</b>".format(colored))
report.line(u"Skipped: <b>{}</b>".format(skipped))
if skipped_details:
    report.subheader("Skip Reasons")
    for d in skipped_details:
        report.warn(d)
if errors:
    report.subheader("Errors")
    for e in errors:
        report.error(e)
report.flush()