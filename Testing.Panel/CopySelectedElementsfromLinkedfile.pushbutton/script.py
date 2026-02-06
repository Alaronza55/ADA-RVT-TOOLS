# -*- coding: utf-8 -*-
__title__ = "Copy Pre-Selected\nfrom Link"
__author__ = "Alaronza"
__doc__ = """SELECT ELEMENTS FROM LINK FIRST, THEN RUN THIS SCRIPT.
Copies pre-selected elements from link to host document."""

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import StructuralType
from pyrevit import revit, forms, script
import math

doc = revit.doc
uidoc = revit.uidoc

# Get current selection
selection = [doc.GetElement(id) for id in uidoc.Selection.GetElementIds()]

if not selection:
    forms.alert("Please select elements from a link first, then run the script.", exitscript=True)

forms.alert("Processing {} selected elements...".format(len(selection)))

# Process
created = 0

t = Transaction(doc, "Copy")
t.Start()

levels = sorted(FilteredElementCollector(doc).OfClass(Level), key=lambda x: x.Elevation)

for elem in selection:
    try:
        if not isinstance(elem, FamilyInstance):
            continue
        
        s = elem.Symbol
        fam = s.Family.Name
        typ = s.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
        
        # Find symbol
        hs = None
        for sym in FilteredElementCollector(doc).OfClass(FamilySymbol):
            try:
                if (sym.Family.Name == fam and 
                    sym.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString() == typ):
                    hs = sym
                    break
            except:
                pass
        
        if not hs:
            continue
        
        if not hs.IsActive:
            hs.Activate()
        
        loc = elem.Location
        if not hasattr(loc, 'Point'):
            continue
        
        pt = loc.Point
        
        lvl = levels[0]
        for l in levels:
            if l.Elevation <= pt.Z:
                lvl = l
        
        ne = doc.Create.NewFamilyInstance(
            XYZ(pt.X, pt.Y, lvl.Elevation),
            hs,
            lvl,
            StructuralType.NonStructural
        )
        
        if ne:
            # Copy params
            for p in elem.Parameters:
                try:
                    if not p.HasValue:
                        continue
                    n = p.Definition.Name
                    if n in ['Level', 'Elevation from Level']:
                        continue
                    np = ne.LookupParameter(n)
                    if not np or np.IsReadOnly:
                        continue
                    
                    if p.StorageType == StorageType.Integer:
                        np.Set(p.AsInteger())
                    elif p.StorageType == StorageType.Double:
                        np.Set(p.AsDouble())
                    elif p.StorageType == StorageType.String:
                        v = p.AsString()
                        if v:
                            np.Set(v)
                except:
                    pass
            
            # Move Z
            nl = ne.Location
            if hasattr(nl, 'Point'):
                nl.Move(XYZ(0, 0, pt.Z - nl.Point.Z))
            
            created += 1
            
    except:
        pass

t.Commit()

forms.alert("Created: {}".format(created))