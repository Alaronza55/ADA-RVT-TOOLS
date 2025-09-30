# -*- coding: utf-8 -*-
"""List All Model Categories Elements"""

__title__ = "List Model\nCategory Elements"
__author__ = "Your Name"

from pyrevit import revit, DB, script
from datetime import datetime

doc = revit.doc
output = script.get_output()

def get_element_info(element):
    """Extract basic information from an element"""
    try:
        element_id = element.Id.IntegerValue
        element_name = element.Name if hasattr(element, 'Name') and element.Name else "Unnamed"
        
        # Try to get type name
        type_name = "N/A"
        if hasattr(element, 'GetTypeId'):
            type_id = element.GetTypeId()
            if type_id and type_id != DB.ElementId.InvalidElementId:
                element_type = doc.GetElement(type_id)
                if element_type:
                    type_name = element_type.Name
        
        # Try to get level
        level_name = "N/A"
        if hasattr(element, 'LevelId') and element.LevelId:
            level = doc.GetElement(element.LevelId)
            if level:
                level_name = level.Name
        elif hasattr(element, 'Level') and element.Level:
            level_name = element.Level.Name
        
        return {
            'id': element_id,
            'name': element_name,
            'type': type_name,
            'level': level_name
        }
    except:
        return {
            'id': element.Id.IntegerValue if element.Id else 0,
            'name': "Unknown",
            'type': "N/A",
            'level': "N/A"
        }

def format_number(num):
    """Format number with commas - IronPython safe"""
    return "{:,}".format(int(num))

def format_percentage(num):
    """Format percentage - IronPython safe"""
    return "{:.1f}%".format(float(num))

def main():
    """Main function to list all model category elements"""
    
    project_name = doc.Title if doc.Title else "Current Project"
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    output.print_md("# Model Category Elements Analysis")
    output.print_md("**Project:** {}".format(project_name))
    output.print_md("**Analysis Date:** {}".format(current_time))
    output.print_md("---")
    
    # Get all model categories
    categories = doc.Settings.Categories
    model_categories = []
    
    for category in categories:
        try:
            if category.CategoryType == DB.CategoryType.Model:
                model_categories.append(category)
        except:
            continue
    
    output.print_md("**Found {} model categories**".format(len(model_categories)))
    output.print_md("*Analyzing elements...*")
    output.print_md("---")
    
    # Dictionary to store elements by category
    category_elements = {}
    total_elements = 0
    
    # Get all elements (excluding element types)
    collector = DB.FilteredElementCollector(doc).WhereElementIsNotElementType()
    
    for element in collector:
        try:
            if element.Category and element.Category.CategoryType == DB.CategoryType.Model:
                category_name = element.Category.Name
                
                if category_name not in category_elements:
                    category_elements[category_name] = []
                
                element_info = get_element_info(element)
                category_elements[category_name].append(element_info)
                total_elements += 1
        except:
            continue
    
    # Sort categories by element count (descending)
    sorted_categories = sorted(category_elements.items(), key=lambda x: len(x[1]), reverse=True)
    
    # Print summary
    output.print_md("## Summary")
    output.print_md("- **Total Model Elements**: {}".format(format_number(total_elements)))
    output.print_md("- **Categories with Elements**: {}".format(len(category_elements)))
    output.print_md("")
    
    # Create summary table
    summary_table = [["Rank", "Category", "Element Count", "Percentage"]]
    
    for i, (cat_name, elements) in enumerate(sorted_categories, 1):
        count = len(elements)
        percentage = (count * 100.0 / total_elements) if total_elements > 0 else 0.0
        summary_table.append([
            str(i),
            cat_name,
            format_number(count),
            format_percentage(percentage)
        ])
    
    output.print_table(
        table_data=summary_table,
        title="Model Categories Summary"
    )
    
    output.print_md("---")
    
    # Detailed listing for each category
    output.print_md("## Detailed Element Listing")
    
    for cat_name, elements in sorted_categories:
        output.print_md("### {} ({} elements)".format(cat_name, format_number(len(elements))))
        
        # Sort elements by ID for consistent ordering
        elements.sort(key=lambda x: x['id'])
        
        # Create table for this category
        if len(elements) <= 100:  # Show all if <= 100 elements
            element_table = [["Element ID", "Name", "Type", "Level"]]
            
            for element_info in elements:
                element_table.append([
                    str(element_info['id']),
                    element_info['name'][:50] + "..." if len(element_info['name']) > 50 else element_info['name'],
                    element_info['type'][:30] + "..." if len(element_info['type']) > 30 else element_info['type'],
                    element_info['level']
                ])
            
            output.print_table(element_table)
        
        else:  # Show first 50 and last 10 if more than 100 elements
            output.print_md("*Showing first 50 and last 10 elements (total: {})*".format(format_number(len(elements))))
            
            element_table = [["Element ID", "Name", "Type", "Level"]]
            
            # First 50
            for element_info in elements[:50]:
                element_table.append([
                    str(element_info['id']),
                    element_info['name'][:50] + "..." if len(element_info['name']) > 50 else element_info['name'],
                    element_info['type'][:30] + "..." if len(element_info['type']) > 30 else element_info['type'],
                    element_info['level']
                ])
            
            # Separator
            element_table.append(["...", "...", "...", "..."])
            
            # Last 10
            for element_info in elements[-10:]:
                element_table.append([
                    str(element_info['id']),
                    element_info['name'][:50] + "..." if len(element_info['name']) > 50 else element_info['name'],
                    element_info['type'][:30] + "..." if len(element_info['type']) > 30 else element_info['type'],
                    element_info['level']
                ])
            
            output.print_table(element_table)
        
        output.print_md("")  # Add spacing between categories
    
    # Final statistics
    output.print_md("---")
    output.print_md("## Statistics")
    
    if category_elements:
        element_counts = [len(elements) for elements in category_elements.values()]
        avg_elements = sum(element_counts) / float(len(element_counts))
        max_elements = max(element_counts)
        min_elements = min(element_counts)
        
        max_category = max(category_elements.items(), key=lambda x: len(x[1]))[0]
        min_category = min(category_elements.items(), key=lambda x: len(x[1]))[0]
        
        output.print_md("- **Average elements per category**: {:.1f}".format(avg_elements))
        output.print_md("- **Category with most elements**: {} ({} elements)".format(
            max_category, format_number(max_elements)))
        output.print_md("- **Category with fewest elements**: {} ({} elements)".format(
            min_category, format_number(min_elements)))
    
    print("Analysis complete! Found {} model elements across {} categories.".format(
        format_number(total_elements), len(category_elements)))

# Execute the script
if __name__ == '__main__':
    main()
