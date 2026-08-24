"""
Select tags with a leader and square off each leader's elbow: the
elbow point is moved to (X of the leader end, Y of the leader start)
- turning a diagonal leader into a clean 90-degree bend (horizontal
from the tag, then vertical to the element).
"""

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import revit, forms

__title__ = 'Set Leader\nElbow'
__author__ = 'Alaronza'

# Get current document
doc = revit.doc
uidoc = revit.uidoc

def get_tag_leader_info(tag):
    """Get leader end and start positions from a tag"""
    if not tag.HasLeader:
        return None, None, None
    
    # Get the tagged reference
    tagged_refs = tag.GetTaggedReferences()
    if not tagged_refs or len(tagged_refs) == 0:
        return None, None, None
    
    tagged_ref = tagged_refs[0]
    
    # Check if the leader has a free end
    try:
        leader_end = tag.GetLeaderEnd(tagged_ref)
        if leader_end is None:
            return None, None, None
    except:
        # If GetLeaderEnd fails, the tag doesn't have a free end leader
        return None, None, None
    
    # Leader start is the tag head position
    leader_start = tag.TagHeadPosition
    
    return leader_end, leader_start, tagged_ref

try:
    # Prompt user to select multiple tags
    selections = uidoc.Selection.PickObjects(
        ObjectType.Element, 
        "Select tags to set their leader elbow positions (must have free end leaders)"
    )
    
    if not selections:
        forms.alert('No tags selected.', exitscript=True)
    
    # Start transaction
    t = Transaction(doc, 'Set Tag Leader Elbows')
    t.Start()
    
    success_count = 0
    skipped_count = 0
    error_count = 0
    
    try:
        for selection in selections:
            # Get the selected element
            tag = doc.GetElement(selection.ElementId)
            
            # Check if it's a tag with a leader
            if not hasattr(tag, 'HasLeader') or not tag.HasLeader:
                print('Skipped: Element ID {} is not a tag with a leader.'.format(
                    selection.ElementId))
                skipped_count += 1
                continue
            
            # Get leader positions
            leader_end, leader_start, tagged_ref = get_tag_leader_info(tag)
            
            if leader_end is None or leader_start is None or tagged_ref is None:
                print('Skipped: Tag ID {} does not have a free end leader or leader is not visible.'.format(
                    selection.ElementId))
                skipped_count += 1
                continue
            
            # Create new elbow position: X from leader end, Y from leader start
            new_elbow = XYZ(leader_end.X, leader_start.Y, leader_end.Z)
            
            try:
                # Set the new leader elbow position using SetLeaderElbow method
                tag.SetLeaderElbow(tagged_ref, new_elbow)
                success_count += 1
                
            except Exception as e:
                print('Error on tag ID {}: {}'.format(selection.ElementId, str(e)))
                error_count += 1
        
        t.Commit()
        
        # Print summary
        print('\n' + '='*50)
        print('SUMMARY')
        print('='*50)
        print('Successfully processed: {} tags'.format(success_count))
        if skipped_count > 0:
            print('Skipped: {} tags (no free end leader or leader not visible)'.format(skipped_count))
        if error_count > 0:
            print('Errors: {} tags'.format(error_count))
        print('='*50)
        
        if success_count == 0 and skipped_count > 0:
            forms.alert('All selected tags were skipped.\n\n'
                       'Make sure the tags have:\n'
                       '- A visible leader\n'
                       '- A free end (not attached to element)', 
                       title='No Tags Processed')
        
    except Exception as e:
        t.RollBack()
        forms.alert('Transaction error: {}'.format(str(e)), exitscript=True)

except Exception as e:
    if 'cancelled' not in str(e).lower():
        forms.alert('Error: {}'.format(str(e)), exitscript=True)