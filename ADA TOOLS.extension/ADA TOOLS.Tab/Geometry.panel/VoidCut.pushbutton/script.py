# -*- coding: utf-8 -*-
__doc__ = """Pick one or more elements (in this model or inside a Revit
link), rebuild their solid geometry as VOID geometry, then click the
elements of this model that the void should cut.

Revit only cuts with voids that live inside a family, so the script:
  1. extracts the picked elements' solids (taken through the link's
     placement transform when the source is linked, and boolean-unioned
     into as few lumps as Revit allows),
  2. builds a temporary Generic Model family around them - each solid
     becomes a FreeForm element turned into a void, and the family is
     set to "Cut with Voids When Loaded",
  3. loads that family and places one instance exactly where the
     original elements are,
  4. lets you click the elements to cut directly in the model - the
     cursor only accepts categories Revit allows to be cut with voids
     (walls, floors, roofs, framing, generic models...) - and applies
     the cuts with InstanceVoidCutUtils.

Note: a void cut only removes material where the void and the element
actually overlap - clicking an element the void never touches changes
nothing visible.

The cuts are ordinary void cuts: they update if the cut elements move,
and DELETING THE VOID INSTANCE REMOVES ALL ITS CUTS - the report ends
with a clickable link to select it. Linked elements cannot be cut
(links are read-only), only elements of the current model."""
__title__ = "Void Cut\nfrom Element"
__version__ = "Version 1.0"
__author__ = "ADA"

import os
import clr
import tempfile
from datetime import datetime

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from Autodesk.Revit.DB import (
    BooleanOperationsType, BooleanOperationsUtils, BuiltInParameter,
    Family, FilteredElementCollector, FreeFormElement, GeometryElement,
    GeometryInstance, IFamilyLoadOptions, InstanceVoidCutUtils, Level,
    Options, RevitLinkInstance, SaveAsOptions, Solid, SolidUtils,
    Transaction, Transform, ViewDetailLevel, XYZ
)
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import forms, revit, script

# Custom ADA GUI - small button-choice popup (see
# lib/GUI/SelectFromButtons.py) and the shared dark/gold themed report
# (see lib/GUI/ReportTheme.py)
from GUI.forms import select_from_buttons
from GUI.ReportTheme import ADAReport

doc = revit.doc
uidoc = revit.uidoc
app = doc.Application
logger = script.get_logger()

TITLE = __title__.replace("\n", " ")
MIN_VOLUME = 1e-7  # cubic feet - discard slivers


def eid_value(eid):
    """ElementId numeric value, 2024+ safe."""
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue


class FamilyLoadHandler(IFamilyLoadOptions):
    """Overwrite silently."""

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source,
                            overwriteParameterValues):
        return True


class AllowAll(ISelectionFilter):
    def AllowElement(self, element):
        return True

    def AllowReference(self, reference, point):
        return True


class CuttableFilter(ISelectionFilter):
    """Only lets the user click elements that Revit can cut with a void
    (and never the source elements themselves)."""

    def __init__(self, exclude_ids):
        self.exclude_ids = exclude_ids

    def AllowElement(self, element):
        if eid_value(element.Id) in self.exclude_ids:
            return False
        try:
            return InstanceVoidCutUtils.CanBeCutWithVoid(element)
        except Exception:
            return False

    def AllowReference(self, reference, point):
        return False


# ---------------------------------------------------------------------------
# selection
# ---------------------------------------------------------------------------

def pick_source_elements():
    """Ask where the void-source elements are, then pick one or more.
    Returns [(element, link_transform_or_None), ...]. Elements already
    selected before launching count as the host-model choice."""
    source = select_from_buttons(
        ['Elements in this model', 'Elements inside a Revit link'],
        title=TITLE,
        label='Where are the elements to turn into a void?',
        version=__version__)
    if not source:
        script.exit()

    result = []
    if source.startswith('Elements in this'):
        pre = list(uidoc.Selection.GetElementIds())
        if pre:
            for eid in pre:
                el = doc.GetElement(eid)
                if el is not None:
                    result.append((el, None))
            return result

        refs = uidoc.Selection.PickObjects(
            ObjectType.Element, AllowAll(),
            "Select the elements whose geometry becomes the void, then Finish")
        for r in refs:
            el = doc.GetElement(r.ElementId)
            if el is not None:
                result.append((el, None))
        return result

    refs = uidoc.Selection.PickObjects(
        ObjectType.LinkedElement,
        "Select the linked elements whose geometry becomes the void, "
        "then Finish")
    for r in refs:
        link = doc.GetElement(r.ElementId)
        if not isinstance(link, RevitLinkInstance):
            continue
        ldoc = link.GetLinkDocument()
        if ldoc is None:
            continue
        el = ldoc.GetElement(r.LinkedElementId)
        if el is not None:
            result.append((el, link.GetTotalTransform()))
    return result


def pick_cut_targets(exclude_ids):
    """Let the user click, in the model, the elements the void should
    cut - only cuttable elements are selectable."""
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element, CuttableFilter(exclude_ids),
        "Select the elements to cut with the void, then Finish")
    targets = []
    for r in refs:
        el = doc.GetElement(r.ElementId)
        if el is not None:
            targets.append(el)
    return targets


# ---------------------------------------------------------------------------
# geometry extraction (same technique as Geo to Mass)
# ---------------------------------------------------------------------------

def _walk(geo, solids):
    for g in geo:
        if isinstance(g, Solid):
            if g.Faces.Size > 0 and g.Volume > MIN_VOLUME:
                solids.append(g)
        elif isinstance(g, GeometryInstance):
            # GetInstanceGeometry() is already in model coordinates
            _walk(g.GetInstanceGeometry(), solids)
        elif isinstance(g, GeometryElement):
            _walk(g, solids)


def extract_solids(element, transform=None):
    """The element's solids in THIS project's internal coordinates."""
    opt = Options()
    opt.DetailLevel = ViewDetailLevel.Fine
    opt.ComputeReferences = False
    opt.IncludeNonVisibleObjects = False

    solids = []
    try:
        geo = element.get_Geometry(opt)
    except Exception as ex:
        logger.debug('geometry failed on %s: %s', eid_value(element.Id), ex)
        return solids
    if geo is not None:
        _walk(geo, solids)

    if transform is not None and not transform.IsIdentity:
        solids = [SolidUtils.CreateTransformed(s, transform) for s in solids]
    return solids


def _try_union(a, b):
    """Union two solids, with a micro-nudge retry for coincident-face failures."""
    try:
        return BooleanOperationsUtils.ExecuteBooleanOperation(
            a, b, BooleanOperationsType.Union)
    except Exception:
        pass
    for d in (1e-5, -1e-5, 5e-5):
        try:
            nudged = SolidUtils.CreateTransformed(
                b, Transform.CreateTranslation(XYZ(d, d, d)))
            return BooleanOperationsUtils.ExecuteBooleanOperation(
                a, nudged, BooleanOperationsType.Union)
        except Exception:
            continue
    return None


def union_solids(solids):
    """Multi-pass boolean union. Returns as few solids as Revit will allow."""
    pool = [s for s in solids if s is not None and s.Volume > MIN_VOLUME]
    if len(pool) < 2:
        return pool
    pool.sort(key=lambda s: s.Volume, reverse=True)

    progress = True
    while progress and len(pool) > 1:
        progress = False
        result = [pool[0]]
        for s in pool[1:]:
            merged = None
            for i in range(len(result)):
                merged = _try_union(result[i], s)
                if merged is not None:
                    result[i] = merged
                    progress = True
                    break
            if merged is None:
                result.append(s)
        pool = result
    return pool


def solids_center(solids):
    """Center of the overall bounding box of the solids (host coords)."""
    xs, ys, zs = [], [], []
    for s in solids:
        bb = s.GetBoundingBox()
        for cx in (bb.Min.X, bb.Max.X):
            for cy in (bb.Min.Y, bb.Max.Y):
                for cz in (bb.Min.Z, bb.Max.Z):
                    p = bb.Transform.OfPoint(XYZ(cx, cy, cz))
                    xs.append(p.X)
                    ys.append(p.Y)
                    zs.append(p.Z)
    return XYZ((min(xs) + max(xs)) / 2.0,
               (min(ys) + max(ys)) / 2.0,
               (min(zs) + max(zs)) / 2.0)


# ---------------------------------------------------------------------------
# void family
# ---------------------------------------------------------------------------

def find_generic_model_template():
    """Locate the plain (non-hosted) Generic Model family template,
    handling localized installs (e.g. French 'Modèle générique
    métrique'). Falls back to asking the user for the .rft."""
    base = app.FamilyTemplatePath
    exact = ('generic model.rft', 'metric generic model.rft',
             u'modèle générique.rft', u'modèle générique métrique.rft')
    # variants like "face based"/"line based"/"adaptive" would host or
    # distort the void placement - reject them
    reject = ('face', 'line', 'adaptive', 'ceiling', 'wall', 'floor',
              'roof', 'pattern', 'rpc', 'tag',
              u'adaptatif', u'plafond', u'mur', u'sol', u'toit', u'ligne')

    best = None
    if base and os.path.isdir(base):
        for root, _dirs, files in os.walk(base):
            for f in files:
                low = f.lower()
                if not low.endswith('.rft'):
                    continue
                if low in exact:
                    return os.path.join(root, f)
                if ('generic model' in low or u'générique' in low
                        or 'generique' in low):
                    if any(k in low for k in reject):
                        continue
                    if best is None or len(f) < len(os.path.basename(best)):
                        best = os.path.join(root, f)
    if best:
        return best

    forms.alert("Could not locate the Generic Model family template (.rft) "
                "automatically.\nPlease pick it in the next dialog.")
    picked = forms.pick_file(file_ext='rft')
    if not picked:
        forms.alert("Cancelled.", exitscript=True)
    return picked


def build_void_family(solids_at_origin, family_name):
    """Create, save and load a Generic Model family whose only geometry
    is the given solids as VOIDS, with 'Cut with Voids When Loaded'
    enabled. Returns (FamilySymbol_or_None, voids_created)."""
    template = find_generic_model_template()
    fam_doc = app.NewFamilyDocument(template)

    made = 0
    ft = Transaction(fam_doc, "ADA - Void geometry")
    ft.Start()
    try:
        allow_cut = fam_doc.OwnerFamily.get_Parameter(
            BuiltInParameter.FAMILY_ALLOW_CUT_WITH_VOIDS)
        if allow_cut is not None:
            allow_cut.Set(1)

        for s in solids_at_origin:
            try:
                ff = FreeFormElement.Create(fam_doc, s)
                is_void = ff.get_Parameter(BuiltInParameter.ELEMENT_IS_CUTTING)
                if is_void is None:
                    raise Exception("no Solid/Void parameter on FreeForm")
                is_void.Set(1)  # 1 = Void
                made += 1
            except Exception as ex:
                logger.debug('FreeFormElement failed: %s', ex)
        ft.Commit()
    except Exception:
        ft.RollBack()
        fam_doc.Close(False)
        raise

    if made == 0:
        fam_doc.Close(False)
        return None, 0

    folder = os.path.join(tempfile.gettempdir(), "ADA_VoidCut")
    if not os.path.isdir(folder):
        os.makedirs(folder)
    path = os.path.join(folder, family_name + ".rfa")

    sao = SaveAsOptions()
    sao.OverwriteExistingFile = True
    fam_doc.SaveAs(path, sao)
    fam_doc.Close(False)

    fam_ref = clr.Reference[Family]()
    if not doc.LoadFamily(path, FamilyLoadHandler(), fam_ref):
        for f in FilteredElementCollector(doc).OfClass(Family):
            if f.Name == family_name:
                fam_ref.Value = f
                break

    fam = fam_ref.Value
    if fam is None:
        return None, made
    sym_ids = list(fam.GetFamilySymbolIds())
    if not sym_ids:
        return None, made
    return doc.GetElement(sym_ids[0]), made


def place_void_instance(symbol, location):
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()
    try:
        return doc.Create.NewFamilyInstance(
            location, symbol, StructuralType.NonStructural)
    except Exception:
        lvl = FilteredElementCollector(doc).OfClass(Level).FirstElement()
        return doc.Create.NewFamilyInstance(
            location, symbol, lvl, StructuralType.NonStructural)


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

try:
    try:
        picked = pick_source_elements()
    except OperationCanceledException:
        script.exit()

    if not picked:
        forms.alert("Nothing selected.", exitscript=True)

    # --- harvest the void geometry ------------------------------------
    solids = []
    no_geometry = []
    for el, tf in picked:
        el_solids = extract_solids(el, tf)
        if el_solids:
            solids.extend(el_solids)
        else:
            no_geometry.append(el)

    if not solids:
        forms.alert("The selected element(s) have no usable solid geometry.",
                    exitscript=True)
    solids = union_solids(solids)

    # Family geometry is built around the family origin; the instance is
    # then placed back at the original center, restoring the position
    center = solids_center(solids)
    to_origin = Transform.CreateTranslation(XYZ.Zero.Subtract(center))
    solids_at_origin = [SolidUtils.CreateTransformed(s, to_origin)
                        for s in solids]

    # host-model sources must not be offered as things to cut
    exclude_ids = set(eid_value(el.Id) for el, tf in picked if tf is None)

    # --- pick what to cut, directly in the model ----------------------
    try:
        targets = pick_cut_targets(exclude_ids)
    except OperationCanceledException:
        script.exit()
    if not targets:
        forms.alert("No elements selected to cut - nothing was changed.",
                    exitscript=True)

    report = ADAReport(TITLE)
    report.line("Void source: <b>{}</b> element(s), {} void solid(s) "
                "after union".format(len(picked), len(solids)))
    for el in no_geometry:
        report.warn("Source element <b>{}</b> [{}] has no solid geometry, "
                    "skipped.".format(el.Name, eid_value(el.Id)))

    family_name = "ADA_VoidCut_{}".format(datetime.now().strftime("%H%M%S"))

    cut_ok = []
    cut_failed = []
    instance = None

    t = Transaction(doc, "ADA - Void Cut from Element")
    t.Start()
    try:
        symbol, void_count = build_void_family(solids_at_origin, family_name)
        if symbol is None:
            raise Exception("could not build/load the void family "
                            "({} void solid(s) created)".format(void_count))

        instance = place_void_instance(symbol, center)
        doc.Regenerate()

        for target in targets:
            try:
                InstanceVoidCutUtils.AddInstanceVoidCut(doc, target, instance)
                cut_ok.append(target)
            except Exception as cut_err:
                cut_failed.append((target, str(cut_err)))
        t.Commit()
    except Exception as ex:
        t.RollBack()
        report.error("Failed, nothing was changed: {}".format(ex))
        report.flush()
        script.exit()

    if cut_ok:
        report.subheader("Cut Elements")
        report.table(
            ["Element", "Category", "Select"],
            [[el.Name,
              el.Category.Name if el.Category else "?",
              report.link(el.Id, title=str(eid_value(el.Id)))]
             for el in cut_ok])

    for target, err in cut_failed:
        report.warn("Could not cut <b>{}</b> [{}]: {}".format(
            target.Name, eid_value(target.Id), err))

    report.subheader("Summary")
    report.line("Elements cut: <b>{}</b> of {} chosen".format(
        len(cut_ok), len(targets)))
    report.line("Void family: <b>{}</b>".format(family_name))
    if instance is not None:
        report.line("Void instance: {} &nbsp;<i>(deleting it removes all "
                    "its cuts)</i>".format(
                        report.link(instance.Id, title="Select void instance")))
    if cut_ok:
        report.success("{} element(s) cut by the new void.".format(len(cut_ok)))
    report.flush()

except Exception as e:
    msg = str(e).lower()
    if 'cancel' not in msg and 'abort' not in msg:
        report = report if 'report' in globals() else ADAReport(TITLE)
        report.error("Error: {}".format(e))
        report.flush()
        import traceback
        print(traceback.format_exc())
