# -*- coding: utf-8 -*-
"""
Generate the ACO Servokat GD/SS D400 800x1000 gully cover family from
scratch, from geometry hardcoded in this script (extracted from the
manufacturer's DXF - stock no. 58612, rev 1). Creates a new family
document from the Metric Generic Model template, builds the frame,
lid and hinge solids, adds type parameters, and saves the .rfa to a
fixed local path (OUTPUT_PATH, edit the constant to change it). Not
family-type-swap like the other Testing tools - this builds the
family file itself. Can be run from any open document.
"""
__title__ = "ACO Servokat GD\nFamily Generator"
__author__ = "ADA"
#
# Source: ACO Industries k.s. drawing, stock no. 58612, rev 1, created 11.12.2009
#         Title: "ACO Servokat GD 800x1000 standard 1.4301", project D400
#
# ASCII-only source, IronPython 2 compatible (ADA TOOLS convention).
#
# ---------------------------------------------------------------------------
# GEOMETRY EXTRACTED FROM THE DXF (all mm, 1:1, verified against the hatched
# cut regions in sections A-A / B-B and the plan view outline)
#
#   Clear opening (daylight)          800 x 1000
#   Frame outer (body)                944 x 1144      -> 72 mm rim all round
#   Hinge-side flange plate           145 wide x 1144 long x 8 thick
#   Overall footprint                 1089 x 1144
#   Frame depth (below top surface)   130
#   Cover lid plate                   928 x 1128 x 8 thick (8 mm gap in frame)
#   Sheet thickness (typ.)            3
#   Lid open angle                    88.5 deg
#   Cover lid mass                    130.57 kg (frame mass not in parts list)
#   Material                          stainless 1.4301 / gasket EPDM (D400)
#
# Origin convention: X = 800 direction (hinge flange on +X),
#                    Y = 1000 direction, Z = 0 at the finished top surface,
#                    frame runs down to Z = -130.
# ---------------------------------------------------------------------------

import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import (
    Transaction, XYZ, Line, Plane, SketchPlane, CurveArray, CurveArrArray,
    SaveAsOptions, Material, Color, BuiltInParameter, UnitUtils,
    BuiltInCategory, FilteredElementCollector
)

# BuiltInParameterGroup and ParameterType were removed from the API in
# Revit 2024. Probe for them instead of importing at module level, so the
# script loads on both old and new versions.
try:
    from Autodesk.Revit.DB import BuiltInParameterGroup, ParameterType
    HAS_LEGACY_ENUMS = True
except ImportError:
    BuiltInParameterGroup = None
    ParameterType = None
    HAS_LEGACY_ENUMS = False

import os

# --------------------------------------------------------------------------
# USER SETTINGS
# --------------------------------------------------------------------------
OUTPUT_PATH = r"C:\Temp\ACO_Servokat_GD_D400_800x1000.rfa"

# Leave as None to auto-detect the Metric Generic Model template.
TEMPLATE_PATH = None

# Candidate templates, first hit wins (adjust for your Revit version/language)
TEMPLATE_CANDIDATES = [
    r"C:\ProgramData\Autodesk\RVT 2026\Family Templates\English\Metric Generic Model.rft",
    r"C:\ProgramData\Autodesk\RVT 2025\Family Templates\English\Metric Generic Model.rft",
    r"C:\ProgramData\Autodesk\RVT 2024\Family Templates\English\Metric Generic Model.rft",
    r"C:\ProgramData\Autodesk\RVT 2023\Family Templates\English\Metric Generic Model.rft",
    r"C:\ProgramData\Autodesk\RVT 2022\Family Templates\English\Metric Generic Model.rft",
    r"C:\ProgramData\Autodesk\RVT 2021\Family Templates\English\Metric Generic Model.rft",
]

# --------------------------------------------------------------------------
# PRODUCT DIMENSIONS (mm)
# --------------------------------------------------------------------------
CO_W = 800.0        # clear opening, X
CO_L = 1000.0       # clear opening, Y
FR_W = 944.0        # frame outer, X
FR_L = 1144.0       # frame outer, Y
FR_D = 130.0        # frame depth
LID_W = 928.0       # lid plate, X
LID_L = 1128.0      # lid plate, Y
LID_T = 8.0         # lid / rebate thickness
HINGE_W = 145.0     # hinge-side flange plate width
HINGE_T = 8.0
SHEET_T = 3.0
OPEN_ANGLE = 88.5
LID_MASS = 130.57

ARTICLE = "58612"
MODEL_NAME = "Servokat GD 800x1000 D400"
MANUFACTURER = "ACO Industries k.s."
MATERIAL_GRADE = "1.4301 (AISI 304)"
GASKET = "EPDM"
LOAD_CLASS = "D400 (EN 124)"

doc_ui = __revit__.ActiveUIDocument                     # noqa: F821
app = __revit__.Application                             # noqa: F821


# --------------------------------------------------------------------------
# HELPERS
# --------------------------------------------------------------------------
def mm(value):
    """Millimetres -> Revit internal units, version tolerant."""
    try:
        from Autodesk.Revit.DB import UnitTypeId
        return UnitUtils.ConvertToInternalUnits(float(value), UnitTypeId.Millimeters)
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import DisplayUnitType
        return UnitUtils.ConvertToInternalUnits(
            float(value), DisplayUnitType.DUT_MILLIMETERS)
    except Exception:
        return float(value) / 304.8


def find_template():
    if TEMPLATE_PATH and os.path.exists(TEMPLATE_PATH):
        return TEMPLATE_PATH
    for candidate in TEMPLATE_CANDIDATES:
        if os.path.exists(candidate):
            return candidate
    # last resort: walk the Autodesk ProgramData tree
    root = r"C:\ProgramData\Autodesk"
    if os.path.isdir(root):
        for dirpath, dirnames, filenames in os.walk(root):
            for name in filenames:
                if name.lower() == "metric generic model.rft":
                    return os.path.join(dirpath, name)
    return None


def rect_loop(x0, y0, x1, y1, z):
    """Closed rectangular CurveArray on plane z (values in mm)."""
    pts = [
        XYZ(mm(x0), mm(y0), mm(z)),
        XYZ(mm(x1), mm(y0), mm(z)),
        XYZ(mm(x1), mm(y1), mm(z)),
        XYZ(mm(x0), mm(y1), mm(z)),
    ]
    arr = CurveArray()
    for i in range(4):
        arr.Append(Line.CreateBound(pts[i], pts[(i + 1) % 4]))
    return arr


def make_extrusion(fdoc, loops, z_bottom, z_top, name):
    """loops = list of CurveArray. First loop is outer, the rest are holes."""
    profile = CurveArrArray()
    for loop in loops:
        profile.Append(loop)

    plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(0, 0, 0))
    sk = SketchPlane.Create(fdoc, plane)

    height = abs(z_top - z_bottom)
    ext = fdoc.FamilyCreate.NewExtrusion(True, profile, sk, mm(height))
    # reposition: offsets are measured from the sketch plane (Z = 0)
    ext.StartOffset = mm(z_bottom)
    ext.EndOffset = mm(z_top)
    return ext


def add_param(fmgr, name, ptype_name, group_name, is_instance=False):
    """Version tolerant FamilyManager.AddParameter."""
    # Revit 2022+ : ForgeTypeId signature
    try:
        from Autodesk.Revit.DB import SpecTypeId, GroupTypeId
        spec_map = {
            "Length": SpecTypeId.Length,
            "Text": SpecTypeId.String.Text,
            "Number": SpecTypeId.Number,
            "Mass": SpecTypeId.Mass,
            "Angle": SpecTypeId.Angle,
        }
        group_map = {
            "Dimensions": GroupTypeId.Geometry,
            "Identity": GroupTypeId.IdentityData,
        }
        return fmgr.AddParameter(name,
                                 group_map.get(group_name, GroupTypeId.Data),
                                 spec_map.get(ptype_name, SpecTypeId.String.Text),
                                 is_instance)
    except Exception:
        pass
    # Revit 2023 and older
    if not HAS_LEGACY_ENUMS:
        raise Exception("AddParameter failed: no usable signature for '%s'" % name)

    ptype_map = {
        "Length": ParameterType.Length,
        "Text": ParameterType.Text,
        "Number": ParameterType.Number,
        "Mass": ParameterType.Number,
        "Angle": ParameterType.Angle,
    }
    group_map = {
        "Dimensions": BuiltInParameterGroup.PG_GEOMETRY,
        "Identity": BuiltInParameterGroup.PG_IDENTITY_DATA,
    }
    return fmgr.AddParameter(name,
                             group_map.get(group_name, BuiltInParameterGroup.PG_DATA),
                             ptype_map.get(ptype_name, ParameterType.Text),
                             is_instance)


def set_param(fmgr, fparam, value):
    try:
        fmgr.Set(fparam, value)
    except Exception:
        try:
            fmgr.Set(fparam, str(value))
        except Exception:
            pass


def get_or_create_material(fdoc, name, r, g, b):
    for m in FilteredElementCollector(fdoc).OfClass(Material):
        if m.Name == name:
            return m
    mid = Material.Create(fdoc, name)
    mat = fdoc.GetElement(mid)
    mat.Color = Color(r, g, b)
    mat.Shininess = 128
    mat.Smoothness = 80
    return mat


def make_subcategory(fdoc, parent_cat, name):
    try:
        subs = parent_cat.SubCategories
        for sc in subs:
            if sc.Name == name:
                return sc
        return fdoc.Settings.Categories.NewSubcategory(parent_cat, name)
    except Exception:
        return None


# --------------------------------------------------------------------------
# MAIN
# --------------------------------------------------------------------------
def main():
    template = find_template()
    if not template:
        print("ERROR: Metric Generic Model.rft not found. "
              "Set TEMPLATE_PATH manually at the top of this script.")
        return

    fdoc = app.NewFamilyDocument(template)
    if fdoc is None:
        print("ERROR: could not create the family document.")
        return

    t = Transaction(fdoc, "Build ACO Servokat GD 800x1000")
    t.Start()
    try:
        # ---- category ------------------------------------------------------
        cats = fdoc.Settings.Categories
        gm = cats.get_Item(BuiltInCategory.OST_GenericModel)
        fdoc.OwnerFamily.FamilyCategory = gm

        sc_frame = make_subcategory(fdoc, gm, "ACO - Frame")
        sc_cover = make_subcategory(fdoc, gm, "ACO - Cover")
        sc_hinge = make_subcategory(fdoc, gm, "ACO - Hinge flange")

        mat_ss = get_or_create_material(fdoc, "Stainless Steel 1.4301",
                                        180, 182, 185)

        # ---- half dimensions ----------------------------------------------
        fw, fl = FR_W / 2.0, FR_L / 2.0
        cw, cl = CO_W / 2.0, CO_L / 2.0
        lw, ll = LID_W / 2.0, LID_L / 2.0

        # 1. Lower frame ring: outer 944x1144, hole 800x1000, -130 .. -8
        lower = make_extrusion(
            fdoc,
            [rect_loop(-fw, -fl, fw, fl, 0),
             rect_loop(-cw, -cl, cw, cl, 0)],
            -FR_D, -LID_T, "Frame body")

        # 2. Upper frame rim (forms the 8 mm rebate the lid sits in):
        #    outer 944x1144, hole 928x1128, -8 .. 0
        upper = make_extrusion(
            fdoc,
            [rect_loop(-fw, -fl, fw, fl, 0),
             rect_loop(-lw, -ll, lw, ll, 0)],
            -LID_T, 0.0, "Frame rim")

        # 3. Hinge-side flange plate: 145 wide x 1144 long x 8 thick, on +X
        hinge = make_extrusion(
            fdoc,
            [rect_loop(fw, -fl, fw + HINGE_W, fl, 0)],
            -HINGE_T, 0.0, "Hinge flange")

        # 4. Cover lid plate: 928 x 1128 x 8, top flush with frame
        lid = make_extrusion(
            fdoc,
            [rect_loop(-lw, -ll, lw, ll, 0)],
            -LID_T, 0.0, "Cover lid")

        # ---- subcategory + material ---------------------------------------
        pairs = [(lower, sc_frame), (upper, sc_frame),
                 (hinge, sc_hinge), (lid, sc_cover)]
        for elem, sub in pairs:
            if sub is not None:
                elem.Subcategory = sub
            p = elem.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
            if p is not None and not p.IsReadOnly:
                p.Set(mat_ss.Id)

        # ---- parameters ----------------------------------------------------
        fm = fdoc.FamilyManager
        text_params = [
            ("ACO_Article_Number", ARTICLE),
            ("ACO_Model", MODEL_NAME),
            ("ACO_Manufacturer", MANUFACTURER),
            ("ACO_Material_Grade", MATERIAL_GRADE),
            ("ACO_Gasket_Material", GASKET),
            ("ACO_Load_Class", LOAD_CLASS),
            ("ACO_Source_Drawing", "DXF 58612 rev.1, 11.12.2009"),
        ]
        for name, val in text_params:
            p = add_param(fm, name, "Text", "Identity", False)
            set_param(fm, p, val)

        length_params = [
            ("ACO_Clear_Opening_Width", CO_W),
            ("ACO_Clear_Opening_Length", CO_L),
            ("ACO_Frame_Outer_Width", FR_W),
            ("ACO_Frame_Outer_Length", FR_L),
            ("ACO_Overall_Width", FR_W + HINGE_W),
            ("ACO_Frame_Depth", FR_D),
            ("ACO_Cover_Thickness", LID_T),
            ("ACO_Sheet_Thickness", SHEET_T),
            ("ACO_Hinge_Flange_Width", HINGE_W),
        ]
        for name, val in length_params:
            p = add_param(fm, name, "Length", "Dimensions", False)
            set_param(fm, p, mm(val))

        p = add_param(fm, "ACO_Cover_Mass_kg", "Number", "Identity", False)
        set_param(fm, p, LID_MASS)
        p = add_param(fm, "ACO_Open_Angle_deg", "Number", "Identity", False)
        set_param(fm, p, OPEN_ANGLE)

        # built-in identity data
        bi = [(BuiltInParameter.ALL_MODEL_MANUFACTURER, MANUFACTURER),
              (BuiltInParameter.ALL_MODEL_MODEL, MODEL_NAME),
              (BuiltInParameter.ALL_MODEL_DESCRIPTION,
               "Hinged access cover, class D400, clear opening 800x1000, "
               "stainless 1.4301")]
        for bip, val in bi:
            try:
                fp = fm.get_Parameter(bip)
                if fp is not None:
                    fm.Set(fp, val)
            except Exception:
                pass

        t.Commit()
    except Exception as ex:
        t.RollBack()
        print("FAILED: " + str(ex))
        try:
            fdoc.Close(False)
        except Exception:
            pass
        return

    # ---- save --------------------------------------------------------------
    folder = os.path.dirname(OUTPUT_PATH)
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)

    opts = SaveAsOptions()
    opts.OverwriteExistingFile = True
    fdoc.SaveAs(OUTPUT_PATH, opts)
    print("Family written to: " + OUTPUT_PATH)
    print("Open it in Revit to review, then Load into Project.")


main()