# -*- coding: utf-8 -*-
__doc__ = """Delete every visualization marker left behind by the QTO
measurement tools (Get Surface, Get Volume, Get Length, Get Length
(Curve), Get Surface (Test - Plane)) - the cones, duplicate
face/volume shells, dimension arrows, and 3D digit readouts they draw
in the model to show what was measured.

These markers are real DirectShape elements created in the active
(host) document, tagged with a fixed name per tool. This script
finds every one of them by that name, across all of the QTO tools at
once, shows you how many it found, and deletes them after
confirmation. It does not touch anything else in the model.
"""
__title__ = "Clear QTO\nMarkers"
__author__ = "ADA"

from pyrevit import revit, DB
from pyrevit import forms, script
from System.Collections.Generic import List

doc = revit.doc

# Every marker name used by the QTO measurement tools. Keep this in
# sync if a tool's MARKER_NAME / TEXT_MARKER_NAME constants change.
MARKER_NAMES = {
    "ADA_QTO_FaceMarker",         # Get Surface - cone
    "ADA_QTO_FacePlaneMarker",    # Get Surface (Test - Plane) - duplicate face plane
    "ADA_QTO_FaceAreaText",       # Get Surface (Test - Plane) - 3D area digits
    "ADA_QTO_VolumeFaceMarker",   # Get Volume - green face shell
    "ADA_QTO_VolumeText",         # Get Volume - 3D volume digits
    "ADA_QTO_LengthArrowMarker",  # Get Length - red dimension arrow
    "ADA_QTO_LengthText",         # Get Length - 3D length digits
    "ADA_QTO_CurveLengthArrowMarker",  # Get Length (Curve) - red dimension arrow
    "ADA_QTO_CurveLengthText",         # Get Length (Curve) - 3D length digits
}

try:
    marker_ids = []
    for ds in DB.FilteredElementCollector(doc).OfClass(DB.DirectShape):
        try:
            if ds.Name in MARKER_NAMES:
                marker_ids.append(ds.Id)
        except Exception:
            pass

    if not marker_ids:
        forms.alert("No QTO measurement markers found in this model.",
                    exitscript=True)

    confirm = forms.alert(
        "{} QTO measurement marker(s) found.\n\nDelete them all?".format(
            len(marker_ids)),
        ok=True, cancel=True
    )
    if not confirm:
        script.exit()

    with revit.Transaction("Clear QTO Markers"):
        doc.Delete(List[DB.ElementId](marker_ids))

    forms.alert("Deleted {} marker(s).".format(len(marker_ids)))

except Exception as e:
    print("Error: {}".format(e))
    import traceback
    traceback.print_exc()
