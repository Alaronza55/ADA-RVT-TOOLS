# -*- coding: utf-8 -*-
"""
Set OPE_ABSOLUTE LEVEL for nested void generic model families
and BES_Lintel Absolute for nested structural framing families.
Only processes elements in the active view with family name containing "BES_Opening + Lintel".
"""

__title__ = "Nested Void\nAbsolute Level"
__author__ = "BESIX"

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    Transaction,
    BasePoint,
    Options
)
from pyrevit import revit, DB, forms
import sys

doc = revit.doc
uidoc = revit.uidoc


def get_survey_point_z():
    """Get the Survey Point Z offset in internal units (feet)."""
    try:
        collector = FilteredElementCollector(doc).OfClass(BasePoint)
        for bp in collector:
            if bp.IsShared:
                return bp.Position.Z
        return 0.0
    except:
        return 0.0


def has_target_family(elem):
    """Safely check if element belongs to a family containing 'BES_Opening + Lintel'."""
    try:
        symbol = elem.Symbol
        if symbol is None:
            return False
        family = symbol.Family
        if family is None:
            return False
        return "BES_Opening + Lintel" in family.Name
    except:
        return False


def get_nested_void_bottom_z(host_instance, survey_z):
    """Get bottom Z of void geometry from host geometry traversal."""
    try:
        opt = Options()
        opt.ComputeReferences = True
        opt.IncludeNonVisibleObjects = True

        geom_elem = host_instance.get_Geometry(opt)
        if geom_elem is None:
            return None

        host_transform = host_instance.GetTransform()
        min_z = None

        for geom_obj in geom_elem:
            if not hasattr(geom_obj, 'GetInstanceGeometry'):
                continue
            for solid in geom_obj.GetInstanceGeometry():
                try:
                    if not hasattr(solid, 'Faces') or solid.Volume == 0:
                        continue
                    bbox = solid.GetBoundingBox()
                    if bbox is None:
                        continue
                    z_world = host_transform.OfPoint(
                        DB.XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z)
                    ).Z
                    if min_z is None or z_world < min_z:
                        min_z = z_world
                except:
                    continue

        return (min_z - survey_z) if min_z is not None else None
    except Exception as e:
        print("Void geometry error on {}: {}".format(host_instance.Id, str(e)))
        return None


def get_nested_lintel(host_instance):
    """
    Find and return the nested FamilyInstance of 'BES_Linteau' via GetSubComponentIds.
    Returns (sub_elem, bottom_z_in_feet) or (None, None).
    """
    try:
        sub_ids = host_instance.GetSubComponentIds()
        for sub_id in sub_ids:
            sub_elem = doc.GetElement(sub_id)
            if sub_elem is None:
                continue
            try:
                symbol = sub_elem.Symbol
                if symbol is None:
                    continue
                family = symbol.Family
                if family is None:
                    continue
                if "BES_Linteau" not in family.Name:
                    continue
            except:
                continue

            # Found — get world-space bounding box
            bbox = sub_elem.get_BoundingBox(None)
            if bbox is not None:
                return sub_elem, bbox.Min.Z

        return None, None
    except Exception as e:
        print("Lintel search error on {}: {}".format(host_instance.Id, str(e)))
        return None, None


def main():
    active_view = uidoc.ActiveView

    collector = (
        FilteredElementCollector(doc, active_view.Id)
        .OfCategory(BuiltInCategory.OST_GenericModel)
        .WhereElementIsNotElementType()
    )

    elements = [elem for elem in collector if has_target_family(elem)]

    if not elements:
        forms.alert("No 'BES_Opening + Lintel' family instances found in the active view.")
        sys.exit()

    survey_z = get_survey_point_z()

    void_success = 0
    void_error = 0
    void_no_param = 0
    lintel_success = 0
    lintel_error = 0
    lintel_no_param = 0

    t = Transaction(doc, "Set Absolute Levels for Opening + Lintel families")
    t.Start()

    try:
        for elem in elements:

            # --- OPE_ABSOLUTE LEVEL (void) ---
            param_ope = elem.LookupParameter("OPE_ABSOLUTE LEVEL")
            if param_ope is None:
                void_no_param += 1
            elif param_ope.IsReadOnly:
                void_error += 1
            else:
                abs_z = get_nested_void_bottom_z(elem, survey_z)
                if abs_z is None:
                    bbox = elem.get_BoundingBox(None)
                    abs_z = (bbox.Min.Z - survey_z) if bbox else None

                if abs_z is not None:
                    try:
                        param_ope.Set(abs_z)
                        void_success += 1
                    except Exception as e:
                        print("OPE set error on {}: {}".format(elem.Id, str(e)))
                        void_error += 1
                else:
                    void_error += 1

            # --- BES_Lintel Absolute (host + nested sub-element) ---
            sub_elem, lintel_min_z = get_nested_lintel(elem)

            if sub_elem is None:
                lintel_error += 1
                continue

            lintel_z = lintel_min_z - survey_z

            # Set on host element
            param_lintel_host = elem.LookupParameter("BES_Lintel Absolute")
            if param_lintel_host is None:
                lintel_no_param += 1
            elif param_lintel_host.IsReadOnly:
                lintel_error += 1
            else:
                try:
                    param_lintel_host.Set(lintel_z)
                except Exception as e:
                    print("Lintel host set error on {}: {}".format(elem.Id, str(e)))
                    lintel_error += 1

            # Set on nested sub-element
            param_lintel_sub = sub_elem.LookupParameter("BES_Lintel Absolute")
            if param_lintel_sub is None:
                lintel_no_param += 1
            elif param_lintel_sub.IsReadOnly:
                lintel_error += 1
            else:
                try:
                    param_lintel_sub.Set(lintel_z)
                    lintel_success += 1
                except Exception as e:
                    print("Lintel sub set error on {}: {}".format(sub_elem.Id, str(e)))
                    lintel_error += 1

        t.Commit()

    except Exception as e:
        t.RollBack()
        forms.alert("Transaction failed: {}".format(str(e)))
        sys.exit()

    msg = (
        "Done!\n\n"
        "OPE_ABSOLUTE LEVEL:\n"
        "  Updated: {}  |  Errors: {}  |  Missing param: {}\n\n"
        "BES_Lintel Absolute:\n"
        "  Updated: {}  |  Errors: {}  |  Missing param: {}"
    ).format(void_success, void_error, void_no_param,
             lintel_success, lintel_error, lintel_no_param)

    if void_error > 0 or void_no_param > 0 or lintel_error > 0 or lintel_no_param > 0:
        forms.alert(msg, warn_icon=True)
    else:
        forms.alert(msg)


if __name__ == "__main__":
    main()