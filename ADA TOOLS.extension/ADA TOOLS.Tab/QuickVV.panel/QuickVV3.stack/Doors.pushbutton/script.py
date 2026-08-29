# -*- coding: utf-8 -*-
__doc__ = """Pick a floor, then pick a door: the floor's thickness (from its
compound structure) is written to the door's "BES_RESA_Under Door"
parameter (positive value) and to its "Sill Height" parameter
(negative value, i.e. the door sits recessed by that thickness)."""
__title__ = "Floor Thickness\nto Door"
__author__ = "ADA"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import revit, forms

# Shared ADA-Tools dark/gold themed report (see lib/GUI/ReportTheme.py)
from GUI.ReportTheme import ADAReport

doc = revit.doc
uidoc = revit.uidoc


class FloorSelectionFilter(ISelectionFilter):
    """Filter to allow only Floor elements"""
    def AllowElement(self, element):
        return isinstance(element, Floor)
    
    def AllowReference(self, reference, point):
        return False


class DoorSelectionFilter(ISelectionFilter):
    """Filter to allow only Door family instances"""
    def AllowElement(self, element):
        if isinstance(element, FamilyInstance):
            return element.Category.Id.IntegerValue == int(BuiltInCategory.OST_Doors)
        return False
    
    def AllowReference(self, reference, point):
        return False


def get_floor_thickness(floor):
    """Get the thickness of a floor element"""
    try:
        floor_type = doc.GetElement(floor.GetTypeId())
        compound_structure = floor_type.GetCompoundStructure()
        
        if compound_structure:
            # Get thickness in internal units (feet)
            thickness = compound_structure.GetWidth()
            return thickness
        else:
            forms.alert("Selected floor does not have a compound structure.", exitscript=True)
    except Exception as e:
        forms.alert("Error getting floor thickness: {}".format(str(e)), exitscript=True)


def main():
    report = ADAReport(__title__.replace(chr(10), " "))
    try:
        # Step 1: Ask user to select a floor
        try:
            floor_ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                FloorSelectionFilter(),
                "Select a Floor"
            )
            floor = doc.GetElement(floor_ref.ElementId)
        except:
            # User cancelled selection
            forms.alert("Floor selection cancelled.", exitscript=True)

        # Step 2: Get floor thickness
        floor_thickness = get_floor_thickness(floor)

        # Convert to millimeters for display
        floor_thickness_mm = UnitUtils.ConvertFromInternalUnits(
            floor_thickness,
            UnitTypeId.Millimeters
        )

        report.line("Floor thickness: <b>{:.2f} mm</b>".format(floor_thickness_mm))

        # Step 3: Ask user to select a door
        try:
            door_ref = uidoc.Selection.PickObject(
                ObjectType.Element,
                DoorSelectionFilter(),
                "Select a Door"
            )
            door = doc.GetElement(door_ref.ElementId)
        except:
            # User cancelled selection
            forms.alert("Door selection cancelled.", exitscript=True)

        # Step 4: Update door parameters
        t = Transaction(doc, "Update Door Parameters from Floor Thickness")
        t.Start()

        try:
            # Set BES_RESA_Under Door to positive floor thickness
            param_under_door = door.LookupParameter("BES_RESA_Under Door")
            if param_under_door and not param_under_door.IsReadOnly:
                param_under_door.Set(floor_thickness)
                report.line("Set 'BES_RESA_Under Door' to <b>{:.2f} mm</b>".format(floor_thickness_mm))
            else:
                report.warn("'BES_RESA_Under Door' parameter not found or is read-only")

            # Set Sill Height to negative floor thickness
            param_sill_height = door.LookupParameter("Sill Height")
            if param_sill_height and not param_sill_height.IsReadOnly:
                param_sill_height.Set(-floor_thickness)
                report.line("Set 'Sill Height' to <b>{:.2f} mm</b>".format(-floor_thickness_mm))
            else:
                report.warn("'Sill Height' parameter not found or is read-only")

            t.Commit()
            report.success("Door parameters updated successfully!")
            report.flush()

        except Exception as e:
            t.RollBack()
            forms.alert("Error updating door parameters: {}".format(str(e)), exitscript=True)

    except Exception as e:
        forms.alert("Script error: {}".format(str(e)), exitscript=True)


if __name__ == "__main__":
    main()