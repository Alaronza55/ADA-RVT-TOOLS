# -*- coding: utf-8 -*-
__title__ = "Distribute Tags\nVertically"
__doc__ = """Distributes annotation tags vertically along a selected line.
- Select a vertical reference line
- Select multiple annotation tags
- Click a point to determine left/right placement
- Tags will be distributed evenly with aligned leaders"""

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import revit, forms, script

# Shared ADA-Tools dark/gold themed report (see lib/GUI/ReportTheme.py)
from GUI.ReportTheme import ADAReport

doc = revit.doc
uidoc = revit.uidoc

# Selection filter for model lines
class LineSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        if isinstance(element, CurveElement):
            curve = element.GeometryCurve
            if curve and isinstance(curve, Line):
                return True
        return False
    
    def AllowReference(self, reference, position):
        return False

# Selection filter for annotation tags
class TagSelectionFilter(ISelectionFilter):
    def AllowElement(self, element):
        return isinstance(element, IndependentTag)
    
    def AllowReference(self, reference, position):
        return False

def get_vertical_line():
    """Prompt user to select a vertical line"""
    try:
        line_filter = LineSelectionFilter()
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            line_filter,
            "Select a vertical reference line"
        )
        
        element = doc.GetElement(ref.ElementId)
        curve = element.GeometryCurve
        
        if not isinstance(curve, Line):
            forms.alert("Selected element is not a line.", exitscript=True)
        
        # Check if line is vertical (or close to vertical)
        direction = curve.Direction
        # For 2D views, check Y direction; for 3D views, check Z direction
        # Allow small tolerance for "vertical" (more than 80 degrees from horizontal)
        if abs(direction.Z) > 0.1:  # 3D line
            if abs(direction.Z) < 0.985:  # cos(10°) ≈ 0.985
                forms.alert("Selected line is not vertical enough. Please select a vertical line.", exitscript=True)
        else:  # 2D line (detail line)
            if abs(direction.Y) < 0.985:  # cos(10°) ≈ 0.985
                forms.alert("Selected line is not vertical enough. Please select a vertical line.", exitscript=True)
        
        return curve
    
    except OperationCanceledException:
        script.exit()

def get_tags():
    """Prompt user to select multiple annotation tags"""
    try:
        tag_filter = TagSelectionFilter()
        refs = uidoc.Selection.PickObjects(
            ObjectType.Element,
            tag_filter,
            "Select annotation tags to distribute"
        )
        
        if not refs:
            forms.alert("No tags selected.", exitscript=True)
        
        tags = [doc.GetElement(ref.ElementId) for ref in refs]
        return tags
    
    except OperationCanceledException:
        script.exit()

def get_reference_point():
    """Prompt user to select a reference point"""
    try:
        point = uidoc.Selection.PickPoint("Click a point to determine tag placement (left or right of line)")
        return point
    
    except OperationCanceledException:
        script.exit()

def is_point_right_of_line(point, line):
    """
    Determine if a point is to the right of a vertical line.
    For a vertical line, "right" means greater X coordinate.
    """
    # Get line's X coordinate (should be consistent for vertical line)
    line_x = line.GetEndPoint(0).X
    return point.X > line_x

def distribute_tags_vertically(tags, line, place_right):
    """
    Distribute tags evenly along the vertical line.
    
    Args:
        tags: List of IndependentTag elements
        line: Line object representing the reference line
        place_right: Boolean indicating if tags should be placed to the right
    """
    if len(tags) == 0:
        return
    
    # Get line endpoints
    start_point = line.GetEndPoint(0)
    end_point = line.GetEndPoint(1)
    
    # Determine if this is a 2D or 3D line based on Z values
    is_2d = abs(start_point.Z - end_point.Z) < 0.001
    
    if is_2d:
        # 2D line - use Y coordinate for vertical distribution
        # Ensure start is lower than end for proper distribution
        if start_point.Y > end_point.Y:
            start_point, end_point = end_point, start_point
        
        # Sort tags by leader end Y coordinate (larger to smaller - top to bottom)
        tags_with_pos = []
        for tag in tags:
            if tag.HasLeader:
                try:
                    tagged_refs_collection = tag.GetTaggedReferences()
                    for ref in tagged_refs_collection:
                        leader_end = tag.GetLeaderEnd(ref)
                        if leader_end:
                            tags_with_pos.append((tag, leader_end.Y))
                        break
                except:
                    # If we can't get leader end, use tag head position
                    tags_with_pos.append((tag, tag.TagHeadPosition.Y))
            else:
                tags_with_pos.append((tag, tag.TagHeadPosition.Y))
    else:
        # 3D line - use Z coordinate for vertical distribution
        # Ensure start is lower than end for proper distribution
        if start_point.Z > end_point.Z:
            start_point, end_point = end_point, start_point
        
        # Sort tags by leader end Z coordinate (larger to smaller - top to bottom)
        tags_with_pos = []
        for tag in tags:
            if tag.HasLeader:
                try:
                    tagged_refs_collection = tag.GetTaggedReferences()
                    for ref in tagged_refs_collection:
                        leader_end = tag.GetLeaderEnd(ref)
                        if leader_end:
                            tags_with_pos.append((tag, leader_end.Z))
                        break
                except:
                    # If we can't get leader end, use tag head position
                    tags_with_pos.append((tag, tag.TagHeadPosition.Z))
            else:
                tags_with_pos.append((tag, tag.TagHeadPosition.Z))
    
    tags_with_pos.sort(key=lambda x: x[1], reverse=False)  # Reverse=False for smaller to larger (bottom to top)
    sorted_tags = [tag for tag, pos in tags_with_pos]
    
    # Calculate line properties
    line_length = line.Length
    line_x = start_point.X
    
    # Calculate spacing
    count = len(tags)
    spacing = line_length / count
    
    # Apply 50cm offset in the direction indicated by place_right
    offset_distance = 500.0 / 304.8  # Convert 50cm (500mm) to feet
    if place_right:
        offset_x = line_x + offset_distance
    else:
        offset_x = line_x - offset_distance
    
    # Start transaction
    t = Transaction(doc, "Distribute Tags Vertically")
    t.Start()

    report = ADAReport(__title__.replace(chr(10), " "))

    try:
        for i, tag in enumerate(sorted_tags):
            if is_2d:
                # Calculate new Y position (center of each segment)
                new_y = start_point.Y + (i + 0.5) * spacing
                
                # Set tag head position with offset from the line's X coordinate
                new_tag_position = XYZ(offset_x, new_y, start_point.Z)
            else:
                # Calculate new Z position (center of each segment)
                new_z = start_point.Z + (i + 0.5) * spacing
                
                # Set tag head position with offset from the line's X coordinate
                new_tag_position = XYZ(offset_x, start_point.Y, new_z)
            
            tag.TagHeadPosition = new_tag_position
            
            # Check if tag has a leader
            if tag.HasLeader:
                try:
                    # Set leader end condition to Free End directly
                    tag.LeaderEndCondition = LeaderEndCondition.Free
                    
                    # Get the tagged references
                    tagged_refs_collection = tag.GetTaggedReferences()
                    
                    # Iterate through the references and set leader elbow
                    for ref in tagged_refs_collection:
                        # Get the leader end point for this reference
                        leader_end = tag.GetLeaderEnd(ref)
                        
                        if leader_end:
                            # Set the leader elbow to the same position as the leader end
                            tag.SetLeaderElbow(ref, leader_end)
                        break  # Only process first reference
                        
                except Exception as e:
                    # Some tags might not support all leader operations
                    report.warn("Could not set leader for tag <b>{}</b>: {}".format(
                        tag.Id.IntegerValue, str(e)
                    ))

        t.Commit()

        report.success("Successfully distributed <b>{}</b> tags vertically.".format(len(sorted_tags)))
        report.flush()

        forms.alert(
            "Successfully distributed {} tags vertically.".format(len(sorted_tags)),
            title="Success"
        )

    except Exception as e:
        t.RollBack()
        report.error("Error distributing tags: {}".format(str(e)))
        report.flush()
        forms.alert(
            "Error distributing tags: {}".format(str(e)),
            title="Error"
        )

# Main execution
def main():
    # Step 1: Select vertical line
    line = get_vertical_line()
    
    # Step 2: Select tags
    tags = get_tags()
    
    # Step 3: Select reference point
    ref_point = get_reference_point()
    
    # Step 4: Determine placement direction
    place_right = is_point_right_of_line(ref_point, line)
    
    # Step 5: Distribute tags
    distribute_tags_vertically(tags, line, place_right)

if __name__ == "__main__":
    main()