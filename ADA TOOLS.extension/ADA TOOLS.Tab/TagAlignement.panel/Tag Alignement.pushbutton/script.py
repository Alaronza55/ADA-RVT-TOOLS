# -*- coding: utf-8 -*-
__title__ = "Tag Alignement Tool"
__doc__ = """This tool will batch align tags to a selected line."""

from pyrevit import revit, DB, UI
from pyrevit import forms

# Get the active document and view
doc = revit.doc
uidoc = revit.uidoc
active_view = doc.ActiveView

try:
    # Prompt user to select a detail line
    forms.alert("Please select a Detail Line in the active view.", 
                title="Select Detail Line")
    
    # Create a selection filter for Detail Lines
    class DetailLineSelectionFilter(UI.Selection.ISelectionFilter):
        def AllowElement(self, element):
            # Check if element is a Detail Line
            if isinstance(element, DB.CurveElement):
                curve_element = element
                # Detail Lines have the category "Lines"
                if curve_element.Category and curve_element.Category.Id.IntegerValue == int(DB.BuiltInCategory.OST_Lines):
                    return True
            return False
        
        def AllowReference(self, reference, position):
            return False
    
    # Get user selection with filter
    selection_filter = DetailLineSelectionFilter()
    selected_ref = uidoc.Selection.PickObject(
        UI.Selection.ObjectType.Element,
        selection_filter,
        "Select a Detail Line"
    )
    
    # Get the selected element
    selected_element = doc.GetElement(selected_ref.ElementId)
    
    # Display information about the selected detail line
    element_id = selected_element.Id.IntegerValue
    forms.alert("Successfully selected Detail Line!\n\n"
                "Element ID: {}\n"
                "Category: {}\n\n"
                "Now select annotation tags.".format(element_id, selected_element.Category.Name),
                title="Detail Line Selected")
    
    # Create a selection filter for Annotation Tags
    class AnnotationTagSelectionFilter(UI.Selection.ISelectionFilter):
        def AllowElement(self, element):
            # Check if element is an annotation tag
            if element.Category:
                category_name = element.Category.Name
                # Common annotation tag categories
                tag_categories = ["Tags", "Room Tags", "Door Tags", "Window Tags", 
                                "Generic Annotations", "Text Notes"]
                if any(tag_cat in category_name for tag_cat in tag_categories):
                    return True
                # Also check for IndependentTag type
                if isinstance(element, DB.IndependentTag):
                    return True
            return False
        
        def AllowReference(self, reference, position):
            return False
    
    # Get user selection of multiple tags
    tag_filter = AnnotationTagSelectionFilter()
    selected_tag_refs = uidoc.Selection.PickObjects(
        UI.Selection.ObjectType.Element,
        tag_filter,
        "Select annotation tags (click Finish when done)"
    )
    
    # Get the selected tag elements
    selected_tags = [doc.GetElement(ref.ElementId) for ref in selected_tag_refs]
    
    # Display information about selected tags
    tag_info = "\n".join(["Tag {}: ID {} - {}".format(i+1, tag.Id.IntegerValue, tag.Category.Name) 
                          for i, tag in enumerate(selected_tags)])
    
    forms.alert("Successfully selected {} annotation tag(s)!\n\n{}\n\n"
                "Now click a point in the view.".format(len(selected_tags), tag_info),
                title="Tags Selected")
    
    # Ask user to pick a point
    picked_point = uidoc.Selection.PickPoint("Click a point in the view")
    
    # Get the curve from the detail line
    curve = selected_element.GeometryCurve
    
    # Get the start and end points of the line
    start_point = curve.GetEndPoint(0)
    end_point = curve.GetEndPoint(1)
    
    # Calculate direction vector of the line (from start to end)
    line_direction = (end_point - start_point).Normalize()
    
    # Check if line is horizontal in 2D view (Y component is negligible)
    tolerance = 0.001
    is_horizontal = abs(line_direction.Y) < tolerance
    
    # Calculate vector from start point to picked point
    to_picked_point = picked_point - start_point
    
    # Calculate cross product (in 2D, we only care about Z component)
    cross_product_z = (line_direction.X * to_picked_point.Y) - (line_direction.Y * to_picked_point.X)
    
    if is_horizontal:
        # For horizontal lines, check if point is above or below
        if cross_product_z < 0:
            side = "DOWN (BELOW)"
        elif cross_product_z > 0:
            side = "UP (ABOVE)"
        else:
            side = "ON THE LINE"
    else:
        # For non-horizontal lines, use cross product for left/right
        if cross_product_z < 0:
            side = "LEFT"
        elif cross_product_z > 0:
            side = "RIGHT"
        else:
            side = "ON THE LINE"
    
    # Display result
    line_type = "HORIZONTAL" if is_horizontal else "NON-HORIZONTAL"
    forms.alert("The selected point is {} the detail line!\n\n"
                "Line type: {}\n"
                "Point coordinates: ({:.2f}, {:.2f})\n"
                "Cross product Z: {:.4f}\n\n"
                "Now the tags will be aligned to the line on the {} side.".format(
                side,
                line_type,
                picked_point.X, 
                picked_point.Y,
                cross_product_z,
                side),
                title="Result")
    
    # Start a transaction to modify tag positions
    from Autodesk.Revit.DB import Transaction
    
    t = Transaction(doc, "Align Tags to Line")
    t.Start()
    
    try:
        # Check if line is vertical (X component is negligible)
        is_vertical = abs(line_direction.X) < tolerance
        
        # FIRST: Store all original tag head positions and leader info before moving anything
        tags_with_leaders = []
        
        for tag in selected_tags:
            original_tag_head = None
            tag_ref = None
            tagged_element_point = None
            
            if hasattr(tag, 'TagHeadPosition'):
                original_tag_head = tag.TagHeadPosition
            elif hasattr(tag, 'Location') and hasattr(tag.Location, 'Point'):
                original_tag_head = tag.Location.Point
            
            if hasattr(tag, 'HasLeader') and tag.HasLeader:
                try:
                    # Get the tagged references
                    tagged_refs = tag.GetTaggedReferences()
                    print("Tag {}: Number of tagged refs: {}".format(tag.Id.IntegerValue, len(tagged_refs)))
                    
                    if tagged_refs and len(tagged_refs) > 0:
                        # Use the first reference
                        tag_ref = tagged_refs[0]
                        
                        print("  Reference type: {}".format(type(tag_ref)))
                        print("  ElementId: {}".format(tag_ref.ElementId))
                        print("  LinkedElementId: {}".format(tag_ref.LinkedElementId))
                        
                        # For linked elements, we need to use the reference to get the point
                        try:
                            # Try to get the global point from the reference
                            if hasattr(tag_ref, 'GlobalPoint'):
                                tagged_element_point = tag_ref.GlobalPoint
                                print("  Got GlobalPoint: {}".format(tagged_element_point))
                        except Exception as e:
                            print("  Error getting GlobalPoint: {}".format(e))
                        
                        if not tagged_element_point:
                            # Check if it's a linked element
                            if tag_ref.LinkedElementId != DB.ElementId.InvalidElementId:
                                print("  This is a linked element")
                                # It's a linked element
                                link_instance = doc.GetElement(tag_ref.ElementId)
                                print("  Link instance: {}".format(link_instance))
                                if link_instance:
                                    link_doc = link_instance.GetLinkDocument()
                                    print("  Link document: {}".format(link_doc))
                                    if link_doc:
                                        linked_elem = link_doc.GetElement(tag_ref.LinkedElementId)
                                        print("  Linked element: {}".format(linked_elem))
                                        if linked_elem:
                                            if hasattr(linked_elem, 'Location') and linked_elem.Location:
                                                print("  Element has Location")
                                                if hasattr(linked_elem.Location, 'Point'):
                                                    print("  Location has Point")
                                                    local_point = linked_elem.Location.Point
                                                    transform = link_instance.GetTransform()
                                                    tagged_element_point = transform.OfPoint(local_point)
                                                    print("  Transformed point: {}".format(tagged_element_point))
                                                elif hasattr(linked_elem.Location, 'Curve'):
                                                    print("  Location has Curve")
                                                    curve = linked_elem.Location.Curve
                                                    local_point = curve.Evaluate(0.5, True)
                                                    transform = link_instance.GetTransform()
                                                    tagged_element_point = transform.OfPoint(local_point)
                                                    print("  Transformed midpoint: {}".format(tagged_element_point))
                                            else:
                                                print("  Element has no Location or Location is None")
                        
                        print("Tag ID {}: TagHead={}, TaggedElementPoint={}".format(
                            tag.Id.IntegerValue, original_tag_head, tagged_element_point))
                except Exception as e:
                    print("Error getting leader info for tag {}: {}".format(tag.Id.IntegerValue, e))
            
            tags_with_leaders.append((tag, original_tag_head, tag_ref, tagged_element_point))
        
        if is_vertical:
            # For vertical lines, align tags vertically with X offset
            offset_distance = 600 / 304.8  # Convert 600mm to feet
            spacing = 200 / 304.8  # Convert 200mm to feet
            
            # Determine the X offset direction based on which side was picked
            if cross_product_z > 0:  # Point is on the RIGHT
                x_offset = offset_distance
            else:  # Point is on the LEFT
                x_offset = -offset_distance
            
            # Get the X coordinate of the line (vertical lines have constant X)
            line_x = start_point.X
            
            # Calculate new X position for all tags
            new_x = line_x + x_offset
            
            # Determine the direction of the line (up or down)
            if end_point.Y > start_point.Y:
                spacing_direction = 1  # Line goes up
            else:
                spacing_direction = -1  # Line goes down
            
            # Sort tags by their current Y position to maintain order
            tags_sorted = []
            sort_direction = None
            
            # Determine if point is left or right of line
            point_is_right = picked_point.X > line_x
            
            for tag, orig_head, ref, elem_point in tags_with_leaders:
                if hasattr(tag, 'TagHeadPosition'):
                    y_pos = tag.TagHeadPosition.Y
                else:
                    y_pos = tag.Location.Point.Y
                
                # Check if we need to sort based on start point Y vs leader end Y
                if elem_point:
                    if point_is_right:
                        # Point is to the RIGHT
                        if start_point.Y > elem_point.Y:
                            # Start point Y is greater - sort ascending (smaller to bigger)
                            sort_direction = 'ascending'
                        elif start_point.Y < elem_point.Y:
                            # Start point Y is smaller - sort descending (bigger to smaller)
                            sort_direction = 'descending'
                    else:
                        # Point is to the LEFT
                        if start_point.Y > elem_point.Y:
                            # Start point Y is greater - sort descending (bigger to smaller)
                            sort_direction = 'descending'
                        elif start_point.Y < elem_point.Y:
                            # Start point Y is smaller - sort ascending (smaller to bigger)
                            sort_direction = 'ascending'
                
                tags_sorted.append((tag, orig_head, ref, elem_point, y_pos))
            
            # Sort by leader end X position based on direction
            if sort_direction == 'ascending':
                tags_sorted.sort(key=lambda x: x[3].X if x[3] else float('inf'))
            elif sort_direction == 'descending':
                tags_sorted.sort(key=lambda x: x[3].X if x[3] else float('-inf'), reverse=True)
            else:
                # No sorting needed, just sort by original Y position
                tags_sorted.sort(key=lambda x: x[4])
            
            # Use the Y position from the start point of the line
            start_y = start_point.Y
            
            for i, (tag, orig_head, tag_ref, tagged_elem_point, original_y) in enumerate(tags_sorted):
                new_y = start_y + (i * spacing * spacing_direction)
                
                # Get current Z coordinate
                if hasattr(tag, 'TagHeadPosition'):
                    current_z = tag.TagHeadPosition.Z
                    
                    # Move tag head
                    new_position = DB.XYZ(new_x, new_y, current_z)
                    tag.TagHeadPosition = new_position
                    
                    # Set elbow position based on where the tagged element is
                    if tag_ref and tagged_elem_point:
                        try:
                            # Elbow should be at: X from tagged element, Y from new tag position
                            elbow_position = DB.XYZ(tagged_elem_point.X, new_y, current_z)
                            tag.SetLeaderElbow(tag_ref, elbow_position)
                            print("Set elbow for tag {} to: {}".format(tag.Id.IntegerValue, elbow_position))
                        except Exception as e:
                            print("Error setting leader for tag {}: {}".format(tag.Id.IntegerValue, e))
                
                elif hasattr(tag.Location, 'Point'):
                    current_z = tag.Location.Point.Z
                    new_position = DB.XYZ(new_x, new_y, current_z)
                    tag.Location.Point = new_position
        
        elif is_horizontal:
            # For horizontal lines, align tags vertically with Y offset
            offset_distance = 200 / 304.8  # Convert 200mm to feet
            spacing = 200 / 304.8  # Convert 200mm to feet
            
            # Get the Y coordinate of the line
            line_y = start_point.Y
            
            # Check if picked point is above or below the line
            if picked_point.Y > line_y:
                # Point is ABOVE the line - stack upwards
                y_offset = offset_distance
                spacing_direction = 1
            else:
                # Point is BELOW the line - stack downwards
                y_offset = -offset_distance
                spacing_direction = -1
            
            # Use the endpoint's Y as the starting position
            start_y = end_point.Y + y_offset
            
            # Use the X position from the end point of the line
            align_x = end_point.X
            
            # Check if we need to sort tags based on leader end positions
            tags_sorted = []
            sort_direction = None
            
            for tag, orig_head, ref, elem_point in tags_with_leaders:
                if elem_point:
                    if start_point.X < elem_point.X:
                        # Start point is to the left - sort ascending (smallest X first)
                        sort_direction = 'ascending'
                    elif start_point.X > elem_point.X:
                        # Start point is to the right - sort descending (biggest X first)
                        sort_direction = 'descending'
                tags_sorted.append((tag, orig_head, ref, elem_point))
            
            # Sort by leader end X position based on direction
            if sort_direction == 'ascending':
                tags_sorted.sort(key=lambda x: x[3].X if x[3] else float('inf'))
            elif sort_direction == 'descending':
                tags_sorted.sort(key=lambda x: x[3].X if x[3] else float('-inf'), reverse=True)
            
            # Position tags vertically with spacing
            for i, (tag, orig_head, tag_ref, tagged_elem_point) in enumerate(tags_sorted):
                new_y = start_y + (i * spacing * spacing_direction)
                
                # Get current Z coordinate
                if hasattr(tag, 'TagHeadPosition'):
                    current_z = tag.TagHeadPosition.Z
                    
                    # Move tag head
                    new_position = DB.XYZ(align_x, new_y, current_z)
                    tag.TagHeadPosition = new_position
                    
                    # Set elbow position based on where the tagged element is
                    if tag_ref and tagged_elem_point:
                        try:
                            # Elbow should be at: X from tagged element, Y from new tag position
                            elbow_position = DB.XYZ(tagged_elem_point.X, new_y, current_z)
                            tag.SetLeaderElbow(tag_ref, elbow_position)
                            print("Set elbow for tag {} to: {}".format(tag.Id.IntegerValue, elbow_position))
                        except Exception as e:
                            print("Error setting leader for tag {}: {}".format(tag.Id.IntegerValue, e))
                
                elif hasattr(tag.Location, 'Point'):
                    current_z = tag.Location.Point.Z
                    new_position = DB.XYZ(align_x, new_y, current_z)
                    tag.Location.Point = new_position
        
        else:
            # For non-vertical/non-horizontal lines, keep original behavior
            curve = selected_element.GeometryCurve
            
            for tag in selected_tags:
                # Get tag head position
                if hasattr(tag, 'TagHeadPosition'):
                    tag_head_pos = tag.TagHeadPosition
                elif hasattr(tag, 'Location') and hasattr(tag.Location, 'Point'):
                    tag_head_pos = tag.Location.Point
                else:
                    continue
                
                # Project the tag head position onto the line
                result = curve.Project(tag_head_pos)
                
                if result:
                    closest_point_on_line = result.XYZPoint
                    vector_to_tag = tag_head_pos - closest_point_on_line
                    
                    if vector_to_tag.GetLength() > 0:
                        perpendicular_direction = vector_to_tag.Normalize()
                        
                        # Calculate which side the tag is currently on
                        to_tag = tag_head_pos - start_point
                        tag_cross_product = (line_direction.X * to_tag.Y) - (line_direction.Y * to_tag.X)
                        
                        # Check if tag is on the same side as the picked point
                        same_side = (tag_cross_product * cross_product_z) > 0
                        
                        if not same_side:
                            perpendicular_direction = perpendicular_direction.Negate()
                        
                        distance_from_line = vector_to_tag.GetLength()
                        new_position = closest_point_on_line + (perpendicular_direction * distance_from_line)
                        
                        # Move the tag
                        if hasattr(tag, 'TagHeadPosition'):
                            tag.TagHeadPosition = new_position
                        elif hasattr(tag.Location, 'Point'):
                            tag.Location.Point = new_position
        
        t.Commit()
        
        forms.alert("Successfully aligned {} tag(s) to the {} side of the line!".format(
                    len(selected_tags), side),
                    title="Tags Aligned")
    
    except Exception as e:
        t.RollBack()
        forms.alert("Error aligning tags: {}".format(str(e)), title="Error")

except Exception as e:
    # Handle cancellation or errors
    if "cancelled" in str(e).lower():
        forms.alert("Selection cancelled.", title="Cancelled")
    else:
        forms.alert("Error: {}".format(str(e)), title="Error"