# -*- coding: utf-8 -*-
__title__ = "Workset Visibility"
__doc__ = "Show or hide worksets in the active view"

from pyrevit import forms, revit, DB, script

# Shared ADA-Tools dark/gold themed report (see lib/GUI/ReportTheme.py) and
# the my_WPF base class (see lib/GUI/WPF_Base.py) that every other ADA-Tools
# popup window uses for its dark/gold look - inheriting from it and calling
# add_wpf_resource() pulls in lib/GUI/Resources/WPF_styles.xaml, which styles
# Button/CheckBox/TextBox/ListBox/ScrollBar implicitly (no XAML needed here).
from GUI.ReportTheme import ADAReport, THEME_BG, THEME_ROW_BG, THEME_GOLD, THEME_GOLD_DARK, THEME_TEXT
from GUI.forms import my_WPF
from System.Collections.Generic import List
from System.Collections.ObjectModel import ObservableCollection
from System.Windows import Window, MessageBox
from System.Windows.Controls import (
    TextBox, ListBox, Button, StackPanel, DockPanel,
    SelectionMode, TextBlock, Grid, RowDefinition, ColumnDefinition,
    CheckBox, ScrollBarVisibility
)
from System.Windows.Data import Binding
from System.Windows.Input import KeyEventArgs, Key
from System.Windows.Media import SolidColorBrush, ColorConverter, LinearGradientBrush, GradientStop
from System import Windows

doc = revit.doc
uidoc = revit.uidoc


def theme_brush(hex_color):
    """Build a SolidColorBrush from a '#RRGGBB' string (same palette as
    lib/GUI/ReportTheme.py / WPF_styles.xaml)."""
    return SolidColorBrush(ColorConverter.ConvertFromString(hex_color))

class WorksetItemUI:
    """UI wrapper for workset with checkbox"""
    def __init__(self, workset, visibility_status):
        self.workset = workset
        self.name = workset.Name
        self.id = workset.Id
        self.visibility_status = visibility_status
        self.is_visible_in_view = visibility_status in ["VISIBLE", "BY_TEMPLATE (Visible)"]
        self.is_checked = False
        self.checkbox = None
        
    def __str__(self):
        return "{} [{}]".format(self.name.replace("_", " ", 1), self.visibility_status)

class WorksetVisibilityWindow(my_WPF):
    """Main window for workset visibility management - themed to match
    the rest of ADA-Tools (dark/gold, borderless with a custom header),
    via the shared my_WPF base and WPF_styles.xaml resource dictionary."""

    def __init__(self, workset_items, view_name):
        self.all_worksets = workset_items
        self.filtered_worksets = list(workset_items)
        self.action = None
        self.selected_worksets = []
        self.view_name = view_name

        # Window setup
        self.Title = "Workset Visibility Manager - {}".format(view_name)
        self.Width = 520
        self.Height = 700
        self.MinHeight = 450
        self.MinWidth = 450
        self.WindowStartupLocation = Windows.WindowStartupLocation.CenterScreen
        self.SizeToContent = Windows.SizeToContent.Manual

        # Same theme as every other ADA-Tools popup (SelectFromDict,
        # SelectFromButtons...): borderless window, own dark background,
        # WPF_styles.xaml pulled in for implicit Button/CheckBox/TextBox/
        # ListBox/ScrollBar styling.
        # "None" is a Python keyword, so the WindowStyle.None enum member
        # has to be reached via getattr rather than dotted attribute access.
        self.WindowStyle = getattr(Windows.WindowStyle, "None")
        self.AllowsTransparency = True
        self.add_wpf_resource()

        gradient = LinearGradientBrush()
        gradient.StartPoint = Windows.Point(0, 1)
        gradient.EndPoint = Windows.Point(1, 0)
        gradient.GradientStops.Add(GradientStop(ColorConverter.ConvertFromString(THEME_BG), 0))
        gradient.GradientStops.Add(GradientStop(ColorConverter.ConvertFromString(THEME_ROW_BG), 1))
        self.Background = gradient

        # Create UI
        self._create_ui()

    def _create_ui(self):
        """Create the user interface"""
        # Main grid with rows
        main_grid = Grid()
        main_grid.Margin = Windows.Thickness(0)

        # Define rows - header/footer added on top of the original 4 rows
        header_row = RowDefinition()
        header_row.Height = Windows.GridLength(25)  # Header - fixed

        row0 = RowDefinition()
        row0.Height = Windows.GridLength(50)  # Search box - fixed

        row1 = RowDefinition()
        row1.Height = Windows.GridLength(40)  # Toggle buttons - fixed

        row2 = RowDefinition()
        row2.Height = Windows.GridLength(1, Windows.GridUnitType.Star)  # List - flexible

        row3 = RowDefinition()
        row3.Height = Windows.GridLength(60)  # Action buttons - fixed

        footer_row = RowDefinition()
        footer_row.Height = Windows.GridLength(25)  # Footer - fixed

        main_grid.RowDefinitions.Add(header_row)
        main_grid.RowDefinitions.Add(row0)
        main_grid.RowDefinitions.Add(row1)
        main_grid.RowDefinitions.Add(row2)
        main_grid.RowDefinitions.Add(row3)
        main_grid.RowDefinitions.Add(footer_row)

        # --- Header: ADA-Tools logo + title + Close button (same layout
        # as SelectFromButtons/SelectFromDict) --------------------------
        header_grid = Grid()
        header_grid.Background = theme_brush(THEME_BG)
        header_grid.MouseDown += self.header_drag

        col_logo = ColumnDefinition()
        col_logo.Width = Windows.GridLength(75)
        col_title = ColumnDefinition()
        col_close = ColumnDefinition()
        col_close.Width = Windows.GridLength(60)
        header_grid.ColumnDefinitions.Add(col_logo)
        header_grid.ColumnDefinitions.Add(col_title)
        header_grid.ColumnDefinitions.Add(col_close)

        logo_text = TextBlock()
        logo_text.Text = "ADA-Tools"
        logo_text.FontWeight = Windows.FontWeights.Heavy
        logo_text.FontSize = 14
        logo_text.Margin = Windows.Thickness(5, 0, 0, 0)
        logo_text.VerticalAlignment = Windows.VerticalAlignment.Center
        logo_text.Foreground = theme_brush(THEME_TEXT)
        Grid.SetColumn(logo_text, 0)
        header_grid.Children.Add(logo_text)

        title_text = TextBlock()
        title_text.Text = "Workset Visibility Manager"
        title_text.VerticalAlignment = Windows.VerticalAlignment.Center
        title_text.HorizontalAlignment = Windows.HorizontalAlignment.Center
        title_text.Foreground = theme_brush(THEME_TEXT)
        Grid.SetColumn(title_text, 1)
        header_grid.Children.Add(title_text)

        close_button = Button()
        close_button.Content = "Close"
        close_button.Width = 60
        close_button.Height = 20
        close_button.FontSize = 10
        close_button.VerticalAlignment = Windows.VerticalAlignment.Center
        close_button.HorizontalAlignment = Windows.HorizontalAlignment.Right
        close_button.Click += self._on_cancel_click
        Grid.SetColumn(close_button, 2)
        header_grid.Children.Add(close_button)

        Grid.SetRow(header_grid, 0)
        main_grid.Children.Add(header_grid)

        # Small subtitle with the active view name, since the header
        # itself has no room for a long dynamic title
        view_label = TextBlock()
        view_label.Text = "View: {}".format(self.view_name)
        view_label.Foreground = theme_brush(THEME_GOLD)
        view_label.FontSize = 11
        view_label.Margin = Windows.Thickness(10, 2, 10, 0)
        view_label.VerticalAlignment = Windows.VerticalAlignment.Top
        Grid.SetRow(view_label, 1)
        main_grid.Children.Add(view_label)

        # Search box panel
        search_panel = DockPanel()
        search_panel.Margin = Windows.Thickness(10, 18, 10, 5)
        search_panel.VerticalAlignment = Windows.VerticalAlignment.Bottom

        search_label = TextBlock()
        search_label.Text = "Search:"
        search_label.Foreground = theme_brush(THEME_GOLD)
        search_label.VerticalAlignment = Windows.VerticalAlignment.Center
        search_label.Margin = Windows.Thickness(0, 0, 10, 0)
        DockPanel.SetDock(search_label, Windows.Controls.Dock.Left)
        search_panel.Children.Add(search_label)

        self.search_box = TextBox()
        self.search_box.VerticalAlignment = Windows.VerticalAlignment.Center
        self.search_box.Height = 25
        self.search_box.TextChanged += self._on_search_changed
        self.search_box.KeyDown += self._on_key_down
        search_panel.Children.Add(self.search_box)

        Grid.SetRow(search_panel, 1)
        main_grid.Children.Add(search_panel)

        # Toggle buttons panel
        toggle_panel = StackPanel()
        toggle_panel.Orientation = Windows.Controls.Orientation.Horizontal
        toggle_panel.HorizontalAlignment = Windows.HorizontalAlignment.Left
        toggle_panel.Margin = Windows.Thickness(10, 0, 10, 5)

        toggle_all_button = Button()
        toggle_all_button.Content = "Toggle All"
        toggle_all_button.Width = 100
        toggle_all_button.Height = 25
        toggle_all_button.Margin = Windows.Thickness(0, 0, 5, 0)
        toggle_all_button.Click += self._on_toggle_all
        toggle_panel.Children.Add(toggle_all_button)

        disable_all_button = Button()
        disable_all_button.Content = "Disable All"
        disable_all_button.Width = 100
        disable_all_button.Height = 25
        disable_all_button.Margin = Windows.Thickness(5, 0, 0, 0)
        disable_all_button.Click += self._on_disable_all
        toggle_panel.Children.Add(disable_all_button)

        Grid.SetRow(toggle_panel, 2)
        main_grid.Children.Add(toggle_panel)

        # Workset list with checkboxes
        self.workset_listbox = ListBox()
        self.workset_listbox.SelectionMode = SelectionMode.Single
        self.workset_listbox.Margin = Windows.Thickness(10, 5, 10, 10)
        self.workset_listbox.VerticalAlignment = Windows.VerticalAlignment.Stretch
        self.workset_listbox.HorizontalAlignment = Windows.HorizontalAlignment.Stretch
        self.workset_listbox.Background = theme_brush(THEME_BG)
        self.workset_listbox.BorderBrush = theme_brush(THEME_GOLD_DARK)
        self._update_listbox()

        Grid.SetRow(self.workset_listbox, 3)
        main_grid.Children.Add(self.workset_listbox)

        # Action button panel
        button_panel = StackPanel()
        button_panel.Orientation = Windows.Controls.Orientation.Horizontal
        button_panel.HorizontalAlignment = Windows.HorizontalAlignment.Center
        button_panel.VerticalAlignment = Windows.VerticalAlignment.Center
        button_panel.Margin = Windows.Thickness(10, 5, 10, 10)

        self.show_button = Button()
        self.show_button.Content = "Show Checked"
        self.show_button.Width = 120
        self.show_button.Height = 30
        self.show_button.Margin = Windows.Thickness(5, 0, 5, 0)
        self.show_button.Click += self._on_show_click
        button_panel.Children.Add(self.show_button)

        self.hide_button = Button()
        self.hide_button.Content = "Hide Checked"
        self.hide_button.Width = 120
        self.hide_button.Height = 30
        self.hide_button.Margin = Windows.Thickness(5, 0, 5, 0)
        self.hide_button.Click += self._on_hide_click
        button_panel.Children.Add(self.hide_button)

        cancel_button = Button()
        cancel_button.Content = "Cancel"
        cancel_button.Width = 80
        cancel_button.Height = 30
        cancel_button.Margin = Windows.Thickness(5, 0, 5, 0)
        cancel_button.Click += self._on_cancel_click
        button_panel.Children.Add(cancel_button)

        Grid.SetRow(button_panel, 4)
        main_grid.Children.Add(button_panel)

        # --- Footer: version, matching the other ADA-Tools windows -----
        footer_grid = Grid()
        footer_grid.Background = theme_brush(THEME_BG)

        footer_text = TextBlock()
        footer_text.Text = "ADA-Tools - Workset Visibility"
        footer_text.Foreground = theme_brush(THEME_GOLD_DARK)
        footer_text.FontSize = 10
        footer_text.HorizontalAlignment = Windows.HorizontalAlignment.Center
        footer_text.VerticalAlignment = Windows.VerticalAlignment.Center
        footer_grid.Children.Add(footer_text)

        Grid.SetRow(footer_grid, 5)
        main_grid.Children.Add(footer_grid)

        # Set content
        self.Content = main_grid

        # Focus search box
        self.Loaded += self._on_loaded
    
    def _on_loaded(self, sender, args):
        """Handle window loaded event"""
        self.search_box.Focus()
    
    def _update_listbox(self):
        """Update listbox with filtered worksets"""
        self.workset_listbox.Items.Clear()
        
        for item in self.filtered_worksets:
            # Create a checkbox for each workset
            checkbox = CheckBox()
            checkbox.IsChecked = item.is_checked
            checkbox.Content = str(item).replace("_", "__")
            checkbox.Tag = item
            checkbox.Checked += self._on_checkbox_changed
            checkbox.Unchecked += self._on_checkbox_changed
            
            # Store reference to checkbox in item
            item.checkbox = checkbox
            
            self.workset_listbox.Items.Add(checkbox)
    
    def _on_checkbox_changed(self, sender, args):
        """Handle checkbox state change"""
        checkbox = sender
        item = checkbox.Tag
        item.is_checked = checkbox.IsChecked
    
    def _on_toggle_all(self, sender, args):
        """Toggle all checkboxes"""
        for item in self.workset_listbox.Items:
            if isinstance(item, CheckBox):
                item.IsChecked = not item.IsChecked
    
    def _on_disable_all(self, sender, args):
        """Disable all checkboxes"""
        for item in self.workset_listbox.Items:
            if isinstance(item, CheckBox):
                item.IsChecked = False
    
    def _on_search_changed(self, sender, args):
        """Handle search text changes"""
        search_text = self.search_box.Text.lower()
        
        if not search_text:
            self.filtered_worksets = list(self.all_worksets)
        else:
            self.filtered_worksets = [
                item for item in self.all_worksets
                if search_text in item.name.replace("_", " ", 1).lower()
            ]
        
        self._update_listbox()
    
    def _on_key_down(self, sender, args):
        """Handle keyboard shortcuts"""
        if args.Key == Key.Escape:
            self.DialogResult = False
            self.Close()
    
    def _get_checked_worksets(self):
        """Get all checked worksets"""
        checked = []
        for workset_item in self.all_worksets:
            if workset_item.is_checked:
                checked.append(workset_item)
        return checked
    
    def _on_show_click(self, sender, args):
        """Handle Show button click"""
        selected = self._get_checked_worksets()
        if not selected:
            return
        
        self.action = "show"
        self.selected_worksets = selected
        self.DialogResult = True
        self.Close()
    
    def _on_hide_click(self, sender, args):
        """Handle Hide button click"""
        selected = self._get_checked_worksets()
        if not selected:
            return
        
        self.action = "hide"
        self.selected_worksets = selected
        self.DialogResult = True
        self.Close()
    
    def _on_cancel_click(self, sender, args):
        """Handle Cancel button click"""
        self.DialogResult = False
        self.Close()

def get_workset_visibility_status(view, workset_id):
    """Get the actual visibility status of a workset in a view"""
    try:
        visibility = view.GetWorksetVisibility(workset_id)
        
        if visibility == DB.WorksetVisibility.Visible:
            return "VISIBLE"
        elif visibility == DB.WorksetVisibility.Hidden:
            return "HIDDEN"
        else:
            # Check if using view template
            view_template_id = view.ViewTemplateId
            if view_template_id != DB.ElementId.InvalidElementId:
                # View is using a template, check template visibility
                view_template = doc.GetElement(view_template_id)
                if view_template:
                    template_visibility = view_template.GetWorksetVisibility(workset_id)
                    if template_visibility == DB.WorksetVisibility.Visible:
                        return "BY_TEMPLATE (Visible)"
                    elif template_visibility == DB.WorksetVisibility.Hidden:
                        return "BY_TEMPLATE (Hidden)"
            return "BY_TEMPLATE"
    except Exception as e:
        print("Error getting visibility for workset: {}".format(str(e)))
        return "UNKNOWN"

def main():
    """Main execution"""
    # Get active view
    active_view = doc.ActiveView
    
    # Check if document is workshared
    if not doc.IsWorkshared:
        forms.alert("This document is not workshared.", exitscript=True)
    
    # Check if view type supports workset visibility
    view_type = active_view.ViewType
    if view_type in [DB.ViewType.Schedule, DB.ViewType.DrawingSheet, 
                     DB.ViewType.Report, DB.ViewType.ProjectBrowser,
                     DB.ViewType.SystemBrowser]:
        forms.alert("Workset visibility cannot be controlled in this view type: {}".format(view_type),
                   exitscript=True)
    
    # Get all worksets
    worksets_collector = DB.FilteredWorksetCollector(doc).OfKind(DB.WorksetKind.UserWorkset)
    
    if not worksets_collector:
        forms.alert("No user worksets found in this document.", exitscript=True)
    
    # Create workset items with visibility info
    workset_items = []
    for workset in worksets_collector:
        visibility_status = get_workset_visibility_status(active_view, workset.Id)
        workset_items.append(WorksetItemUI(workset, visibility_status))
    
    if not workset_items:
        forms.alert("Could not retrieve workset visibility information.", exitscript=True)
    
    # Sort by name
    workset_items.sort(key=lambda x: x.name)
    
    # Show window
    window = WorksetVisibilityWindow(workset_items, active_view.Name)
    result = window.ShowDialog()
    
    # Process after window is closed
    if result and window.action:
        action = window.action
        selected_worksets = window.selected_worksets
        
        # Apply visibility changes
        with revit.Transaction("Change Workset Visibility"):
            for workset_item in selected_worksets:
                workset_id = workset_item.id

                if action == "show":
                    active_view.SetWorksetVisibility(workset_id, DB.WorksetVisibility.Visible)

                elif action == "hide":
                    active_view.SetWorksetVisibility(workset_id, DB.WorksetVisibility.Hidden)

        report = ADAReport(__title__)
        report.success("{} workset(s) set to <b>{}</b> in view <b>{}</b>.".format(
            len(selected_worksets), action, active_view.Name))
        report.table(
            ["Workset"],
            [[w.name] for w in selected_worksets]
        )
        report.flush()

if __name__ == "__main__":
    main()