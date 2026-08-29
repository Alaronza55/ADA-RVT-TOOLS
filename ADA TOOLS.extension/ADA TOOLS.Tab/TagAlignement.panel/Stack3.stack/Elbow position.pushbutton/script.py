__doc__ = """Tag Leader Alignment
Aligns tag leaders to a vertical detail line with user-defined offset"""
__title__ = "Align Tag\nLeaders to Line"
__author__ = "Alaronza"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from pyrevit import revit, DB, UI, forms
import math
import clr

# Shared ADA-Tools dark/gold themed report (see lib/GUI/ReportTheme.py)
from GUI.ReportTheme import ADAReport

doc = revit.doc
uidoc = revit.uidoc


class DetailLineSelectionFilter(ISelectionFilter):
    """Filter to select only detail lines"""
    def AllowElement(self, element):
        # Detail lines are CurveElement with specific category
        if isinstance(element, CurveElement):
            if element.Category and element.Category.Id.IntegerValue == int(BuiltInCategory.OST_Lines):
                curve = element.GeometryCurve
                if curve and isinstance(curve, Line):
                    return True
        return False
    
    def AllowReference(self, reference, position):
        return False


class TagSelectionFilter(ISelectionFilter):
    """Filter to select only tags"""
    def AllowElement(self, element):
        return isinstance(element, IndependentTag)
    
    def AllowReference(self, reference, position):
        return False


def is_line_vertical(line, tolerance=0.01):
    """Check if a line is vertical within tolerance (in 2D view, check X direction)"""
    direction = line.Direction
    # In a 2D view, vertical means minimal X direction change
    return abs(direction.X) < tolerance


def get_line_x_coordinate(line):
    """Get the X coordinate of a vertical line"""
    return line.GetEndPoint(0).X


def mm_to_feet(mm):
    """Convert millimeters to feet (Revit internal units)"""
    return mm / 304.8


def main():
    try:
        # Check if we're in a view that supports detail lines
        active_view = doc.ActiveView
        if not isinstance(active_view, ViewPlan) and not isinstance(active_view, ViewSection):
            forms.alert("Please run this script in a Plan or Section view.", exitscript=True)
        
        # Step 1: Select vertical detail line
        line_filter = DetailLineSelectionFilter()
        line_ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            line_filter,
            "Select a vertical detail line"
        )
        line_element = doc.GetElement(line_ref.ElementId)
        line = line_element.GeometryCurve
        
        if not is_line_vertical(line):
            forms.alert("The selected line is not vertical. Please select a vertical line.", exitscript=True)
        
        line_x = get_line_x_coordinate(line)
        
        # Step 2: Select multiple tags
        tag_filter = TagSelectionFilter()
        tag_refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            tag_filter,
            "Select annotation tags (click Finish when done)"
        )
        
        if not tag_refs:
            forms.alert("No tags selected.", exitscript=True)
        
        tags = [doc.GetElement(ref.ElementId) for ref in tag_refs]
        
        # Step 3: Select a point to determine direction
        try:
            point = uidoc.Selection.PickPoint("Click a point to the LEFT or RIGHT of the line to set direction")
        except:
            forms.alert("Point selection cancelled.", exitscript=True)
        
        # Determine if point is to the right or left of the line
        is_right = point.X > line_x
        direction_text = "RIGHT" if is_right else "LEFT"
        
        # Step 4: Get offset from user in millimeters
        offset_input = forms.ask_for_string(
            default="600",
            prompt="Enter offset distance from line (in millimeters):",
            title="Leader Offset"
        )
        
        if not offset_input:
            forms.alert("No offset provided.", exitscript=True)
        
        try:
            offset_mm = float(offset_input)
            offset_feet = mm_to_feet(offset_mm)
        except:
            forms.alert("Invalid offset value. Please enter a number.", exitscript=True)
        
        # Calculate leader elbow X position
        if is_right:
            leader_elbow_x = line_x + offset_feet
        else:
            leader_elbow_x = line_x - offset_feet
        
        # Step 5: Process tags
        with revit.Transaction("Align Tag Leaders"):
            aligned_count = 0
            skipped_count = 0
            
            for tag in tags:
                if not tag.HasLeader:
                    skipped_count += 1
                    continue
                
                try:
                    # Get the tagged references
                    tagged_refs = tag.GetTaggedReferences()
                    
                    # Convert to list to handle both IList and List types
                    ref_list = list(tagged_refs) if tagged_refs else []
                    
                    if not ref_list:
                        print("Tag {} has no tagged references".format(tag.Id.IntegerValue))
                        skipped_count += 1
                        continue
                    
                    # Get the first reference
                    ref = ref_list[0]
                    
                    print("\nTag {}: Reference type = {}".format(tag.Id.IntegerValue, type(ref)))
                    
                    # Get tag head position (where the tag label is)
                    tag_head = tag.TagHeadPosition
                    
                    # Get current elbow for Z coordinate
                    current_elbow = tag.GetLeaderElbow(ref)
                    
                    # Set elbow position:
                    # X: offset from line in the selected direction
                    # Y: same as tag head (makes leader horizontal from tag to elbow)
                    # Z: same as current elbow (maintains view plane)
                    new_elbow = XYZ(leader_elbow_x, tag_head.Y, current_elbow.Z)
                    
                    print("  Tag head position: X={:.3f}, Y={:.3f}, Z={:.3f}".format(
                        tag_head.X, tag_head.Y, tag_head.Z))
                    print("  Current elbow: X={:.3f}, Y={:.3f}, Z={:.3f}".format(
                        current_elbow.X, current_elbow.Y, current_elbow.Z))
                    print("  New elbow: X={:.3f}, Y={:.3f}, Z={:.3f}".format(
                        new_elbow.X, new_elbow.Y, new_elbow.Z))
                    
                    tag.SetLeaderElbow(ref, new_elbow)
                    
                    # Verify the change
                    updated_elbow = tag.GetLeaderElbow(ref)
                    print("  Updated elbow: X={:.3f}, Y={:.3f}, Z={:.3f}".format(
                        updated_elbow.X, updated_elbow.Y, updated_elbow.Z))
                    
                    aligned_count += 1
                    
                except Exception as e:
                    import traceback
                    print("\nCould not set leader for tag {}: {}".format(
                        tag.Id.IntegerValue, str(e)
                    ))
                    print(traceback.format_exc())
                    skipped_count += 1
        
        # Report results
        report = ADAReport(__title__.replace(chr(10), " "))
        report.line("Direction: <b>{}</b>".format(direction_text))
        report.line("Offset: <b>{} mm</b>".format(offset_mm))
        report.line("Line X position: <b>{:.3f} ft</b>".format(line_x))
        report.line("Leader elbow X: <b>{:.3f} ft</b>".format(leader_elbow_x))
        report.subheader("Summary")
        report.success("Tags aligned: <b>{}</b>".format(aligned_count))
        if skipped_count > 0:
            report.warn("Tags skipped (no leader or no references): <b>{}</b>".format(skipped_count))
        report.flush()

        message = "Leader alignment complete!\n\n"
        message += "Direction: {}\n".format(direction_text)
        message += "Offset: {} mm\n".format(offset_mm)
        message += "Line X position: {:.3f} ft\n".format(line_x)
        message += "Leader elbow X: {:.3f} ft\n\n".format(leader_elbow_x)
        message += "Tags aligned: {}\n".format(aligned_count)
        if skipped_count > 0:
            message += "Tags skipped (no leader or no references): {}".format(skipped_count)

        forms.alert(message, title="Success")

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        report = ADAReport(__title__.replace(chr(10), " "))
        report.error("Error: {}".format(e))
        report.flush()
        forms.alert("Error: {}".format(str(e)), title="Error")


if __name__ == "__main__":
    main()