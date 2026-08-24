# -*- coding: utf-8 -*-
"""Convert the geometry of selected elements into Mass elements.

ADA-RVT-TOOLS / Coordination
IronPython 2 - Revit 2021+
"""

__title__ = "Geo to\nMass"
__author__ = "ADA"
__doc__ = ("Select elements (in this model or inside a link) and rebuild their\n"
           "geometry as Mass elements.\n\n"
           "Two modes:\n"
           " - DirectShape (Mass category): instant, closest equivalent to an\n"
           "   in-place mass. Also handles mesh-only geometry (imported DWG).\n"
           " - Mass family (.rfa): builds a real conceptual mass family with\n"
           "   FreeForm elements, loads it and places it at the internal origin.\n"
           "   Solids only.")

import os
import clr
import tempfile

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BooleanOperationsType, BooleanOperationsUtils, BuiltInCategory, DirectShape,
    ElementId, Family, FamilySymbol, FilteredElementCollector, FreeFormElement,
    GeometryElement, GeometryInstance, GeometryObject, IFamilyLoadOptions, Level,
    Mesh, Options, RevitLinkInstance, SaveAsOptions, Solid, SolidUtils,
    TessellatedFace, TessellatedShapeBuilder, TessellatedShapeBuilderFallback,
    TessellatedShapeBuilderTarget, Transaction, Transform, ViewDetailLevel, XYZ
)
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException

from pyrevit import forms, revit, script

doc = revit.doc
uidoc = revit.uidoc
app = doc.Application
output = script.get_output()
logger = script.get_logger()

MIN_VOLUME = 1e-7          # cubic feet - discard slivers
MASS_CAT = ElementId(BuiltInCategory.OST_Mass)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def eid_value(eid):
    """ElementId numeric value, 2024+ safe."""
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue


def geom_options():
    opt = Options()
    opt.DetailLevel = ViewDetailLevel.Fine
    opt.ComputeReferences = False
    opt.IncludeNonVisibleObjects = False
    return opt


class AllowAll(ISelectionFilter):
    def AllowElement(self, element):
        return True

    def AllowReference(self, reference, point):
        return True


class FamilyLoadHandler(IFamilyLoadOptions):
    """Overwrite silently."""

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source,
                            overwriteParameterValues):
        return True


# ---------------------------------------------------------------------------
# selection
# ---------------------------------------------------------------------------

def pick_host_elements():
    """Returns [(element, transform_or_None), ...] from the current model."""
    result = []
    pre = list(uidoc.Selection.GetElementIds())
    if pre:
        for eid in pre:
            el = doc.GetElement(eid)
            if el is not None:
                result.append((el, None))
        return result

    refs = uidoc.Selection.PickObjects(
        ObjectType.Element, AllowAll(),
        "Select the elements to convert to mass, then Finish")
    for r in refs:
        el = doc.GetElement(r.ElementId)
        if el is not None:
            result.append((el, None))
    return result


def pick_linked_elements():
    """Returns [(element, link_transform), ...] from a Revit link."""
    result = []
    refs = uidoc.Selection.PickObjects(
        ObjectType.LinkedElement,
        "Select the linked elements to convert to mass, then Finish")
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


# ---------------------------------------------------------------------------
# geometry extraction
# ---------------------------------------------------------------------------

def _walk(geo, solids, meshes):
    for g in geo:
        if isinstance(g, Solid):
            if g.Faces.Size > 0 and g.Volume > MIN_VOLUME:
                solids.append(g)
        elif isinstance(g, Mesh):
            if g.NumTriangles > 0:
                meshes.append(g)
        elif isinstance(g, GeometryInstance):
            # GetInstanceGeometry() is already in model coordinates
            _walk(g.GetInstanceGeometry(), solids, meshes)
        elif isinstance(g, GeometryElement):
            _walk(g, solids, meshes)


def extract_geometry(element, transform=None):
    """Return (solids, meshes) in project internal coordinates."""
    solids, meshes = [], []
    try:
        geo = element.get_Geometry(geom_options())
    except Exception as ex:
        logger.debug('geometry failed on %s: %s', eid_value(element.Id), ex)
        return solids, meshes
    if geo is None:
        return solids, meshes

    _walk(geo, solids, meshes)

    if transform is not None and not transform.IsIdentity:
        solids = [SolidUtils.CreateTransformed(s, transform) for s in solids]
        meshes = [m.get_Transformed(transform) for m in meshes]

    return solids, meshes


def _try_union(a, b):
    """Union two solids, with a micro-nudge retry for coincident-face failures."""
    try:
        return BooleanOperationsUtils.ExecuteBooleanOperation(
            a, b, BooleanOperationsType.Union)
    except Exception:
        pass
    # coincident faces are the usual culprit - shift b by ~0.003 mm and retry
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

    # biggest first: unioning into a large solid is more robust than the reverse
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


def mesh_to_geometry(mesh):
    """Tessellated fallback so imported/mesh-only geometry still becomes a mass."""
    builder = TessellatedShapeBuilder()
    builder.OpenConnectedFaceSet(False)
    for i in range(mesh.NumTriangles):
        tri = mesh.get_Triangle(i)
        pts = List[XYZ]()
        pts.Add(tri.get_Vertex(0))
        pts.Add(tri.get_Vertex(1))
        pts.Add(tri.get_Vertex(2))
        builder.AddFace(TessellatedFace(pts, ElementId.InvalidElementId))
    builder.CloseConnectedFaceSet()

    try:
        builder.Target = TessellatedShapeBuilderTarget.AnyGeometry
        builder.Fallback = TessellatedShapeBuilderFallback.Mesh
        builder.GraphicsStyleId = ElementId.InvalidElementId
        builder.Build()
    except Exception:
        builder.Build(TessellatedShapeBuilderTarget.AnyGeometry,
                      TessellatedShapeBuilderFallback.Mesh,
                      ElementId.InvalidElementId)

    return list(builder.GetBuildResult().GetGeometricalObjects())


# ---------------------------------------------------------------------------
# mode A - DirectShape
# ---------------------------------------------------------------------------

def create_directshape(geom_objects, name):
    geo_list = List[GeometryObject]()
    for g in geom_objects:
        geo_list.Add(g)

    ds = DirectShape.CreateElement(doc, MASS_CAT)
    ds.ApplicationId = "ADA-RVT-TOOLS"
    ds.ApplicationDataId = "GeoToMass"
    ds.SetShape(geo_list)
    try:
        ds.Name = name
    except Exception:
        pass
    return ds


# ---------------------------------------------------------------------------
# mode B - real mass family
# ---------------------------------------------------------------------------

def find_mass_template():
    base = app.FamilyTemplatePath
    if not base or not os.path.isdir(base):
        return None

    exact = ('mass.rft', 'metric mass.rft')
    best = None
    for root, _dirs, files in os.walk(base):
        for f in files:
            low = f.lower()
            if low in exact:
                path = os.path.join(root, f)
                if 'conceptual mass' in root.lower():
                    return path
                if best is None:
                    best = path
    return best


def build_mass_family(solids, family_name):
    """Create, save and load a conceptual mass family. Returns the FamilySymbol."""
    template = find_mass_template()
    if not template:
        forms.alert("Could not locate the conceptual Mass template (.rft).\n"
                    "Check Options > File Locations > Family Template Files.",
                    exitscript=True)

    fam_doc = app.NewFamilyDocument(template)

    ft = Transaction(fam_doc, "ADA - Add geometry")
    ft.Start()
    made = 0
    for s in solids:
        try:
            FreeFormElement.Create(fam_doc, s)
            made += 1
        except Exception as ex:
            logger.debug('FreeFormElement failed: %s', ex)
    ft.Commit()

    if made == 0:
        fam_doc.Close(False)
        return None

    folder = os.path.join(tempfile.gettempdir(), "ADA_GeoToMass")
    if not os.path.isdir(folder):
        os.makedirs(folder)
    path = os.path.join(folder, family_name + ".rfa")

    sao = SaveAsOptions()
    sao.OverwriteExistingFile = True
    fam_doc.SaveAs(path, sao)
    fam_doc.Close(False)

    fam_ref = clr.Reference[Family]()
    if not doc.LoadFamily(path, FamilyLoadHandler(), fam_ref):
        # already loaded with that name - fetch it
        for f in FilteredElementCollector(doc).OfClass(Family):
            if f.Name == family_name:
                fam_ref.Value = f
                break

    fam = fam_ref.Value
    if fam is None:
        return None

    sym_ids = list(fam.GetFamilySymbolIds())
    if not sym_ids:
        return None
    return doc.GetElement(sym_ids[0])


def place_mass(symbol):
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()
    try:
        return doc.Create.NewFamilyInstance(
            XYZ.Zero, symbol, StructuralType.NonStructural)
    except Exception:
        lvl = FilteredElementCollector(doc).OfClass(Level).FirstElement()
        return doc.Create.NewFamilyInstance(
            XYZ.Zero, symbol, lvl, StructuralType.NonStructural)


def unhide_mass_category():
    view = doc.ActiveView
    try:
        if view.CanCategoryBeHidden(MASS_CAT) and view.GetCategoryHidden(MASS_CAT):
            view.SetCategoryHidden(MASS_CAT, False)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

source = forms.CommandSwitchWindow.show(
    ['Elements in this model', 'Elements inside a Revit link'],
    message='Where are the elements?')
if not source:
    script.exit()

try:
    if source.startswith('Elements in this'):
        picked = pick_host_elements()
    else:
        picked = pick_linked_elements()
except OperationCanceledException:
    script.exit()

if not picked:
    forms.alert("Nothing selected.", exitscript=True)

mode = forms.CommandSwitchWindow.show(
    ['DirectShape (Mass category)', 'Mass family (.rfa)'],
    message='How should the mass be created?')
if not mode:
    script.exit()

grouping = forms.CommandSwitchWindow.show(
    ['One mass per element', 'One single merged mass'],
    message='Grouping?')
if not grouping:
    script.exit()

use_directshape = mode.startswith('DirectShape')
merge_all = grouping.startswith('One single')

delete_source = False
if source.startswith('Elements in this'):
    delete_source = forms.alert("Delete the original elements afterwards?",
                                yes=True, no=True)

# --- harvest geometry ------------------------------------------------------
buckets = []          # [(label, [GeometryObject, ...]), ...]
skipped = []
all_solids = []
all_meshes = []

with forms.ProgressBar(title='Reading geometry {value}/{max_value}') as pb:
    total = len(picked)
    for i, (el, tf) in enumerate(picked):
        pb.update_progress(i + 1, total)
        solids, meshes = extract_geometry(el, tf)

        if not solids and not meshes:
            skipped.append(el)
            continue

        if merge_all:
            all_solids.extend(solids)
            all_meshes.extend(meshes)
        else:
            try:
                label = "{0} [{1}]".format(el.Name, eid_value(el.Id))
            except Exception:
                label = "Element {0}".format(eid_value(el.Id))
            buckets.append((label, solids, meshes))

union_report = None
if merge_all:
    if not all_solids and not all_meshes:
        forms.alert("No usable geometry found.", exitscript=True)
    fused = union_solids(all_solids)
    union_report = (len(all_solids), len(fused), len(all_meshes))
    buckets = [("MergedMass", fused, all_meshes)]

if not buckets:
    forms.alert("No usable geometry found.", exitscript=True)

# --- create ----------------------------------------------------------------
created = []
failed = []

t = Transaction(doc, "ADA - Geometry to Mass")
t.Start()
try:
    for label, solids, meshes in buckets:
        safe = "".join(c for c in label if c.isalnum() or c in " _-").strip()
        safe = safe or "Mass"

        if use_directshape:
            geo = list(solids)
            for m in meshes:
                try:
                    geo.extend(mesh_to_geometry(m))
                except Exception as ex:
                    logger.debug('mesh conversion failed: %s', ex)
            if not geo:
                failed.append(label)
                continue
            try:
                created.append(create_directshape(geo, "MASS_" + safe))
            except Exception as ex:
                logger.debug('DirectShape failed on %s: %s', label, ex)
                failed.append(label)
        else:
            if not solids:
                failed.append(label + " (mesh-only, needs DirectShape mode)")
                continue
            sym = build_mass_family(solids, "ADA_MASS_" + safe)
            if sym is None:
                failed.append(label)
                continue
            created.append(place_mass(sym))

    if delete_source and created:
        for el, _tf in picked:
            try:
                doc.Delete(el.Id)
            except Exception:
                pass

    unhide_mass_category()
    t.Commit()
except Exception as ex:
    t.RollBack()
    forms.alert("Failed, nothing was changed.\n\n{0}".format(ex), exitscript=True)

# --- report ----------------------------------------------------------------
output.print_md("## Geometry to Mass")
output.print_md("- Masses created: **{0}**".format(len(created)))

if union_report:
    src, res, msh = union_report
    if res == 1:
        output.print_md("- Union: **{0} solids fused into 1**".format(src))
    else:
        output.print_md("- Union: {0} solids reduced to **{1} lumps** "
                        "(Revit refused the rest - they are still inside the "
                        "same single element)".format(src, res))
    if msh:
        output.print_md("- {0} mesh(es) added un-fused "
                        "(meshes cannot be booleaned)".format(msh))

if created:
    ids = List[ElementId]()
    for e in created:
        ids.Add(e.Id)
    uidoc.Selection.SetElementIds(ids)

if skipped:
    output.print_md("- Skipped (no geometry): **{0}**".format(len(skipped)))
    for el in skipped[:25]:
        print(output.linkify(el.Id))

if failed:
    output.print_md("- Failed: **{0}**".format(len(failed)))
    for f in failed[:25]:
        print(f)

if created and not use_directshape:
    output.print_md("_Turn on Massing & Site > Show Mass if the new masses "
                    "are not visible._")