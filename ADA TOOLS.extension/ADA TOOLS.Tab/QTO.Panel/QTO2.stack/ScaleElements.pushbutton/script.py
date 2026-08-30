# -*- coding: utf-8 -*-
__doc__ = """Resize the raw geometry of selected DirectShape / Generic
Model elements - most usefully the markers the other QTO tools (Get
Surface, Get Volume, Get Length, Get Length (Curve)) draw in the active
view, since those are always created at a fixed size and don't scale
with the model.

Select one or more DirectShape elements first (or run the tool with
nothing selected and pick them when prompted), then enter a scale
percentage in the themed dialog - 100% keeps the current size, 200%
doubles it, 50% halves it, and so on. Each element is scaled about its
own bounding-box center, in place. The percentage is relative to each
element's CURRENT size, not its original size, so running the tool
twice at 150% compounds to 225% - this only applies to DirectShape
elements (the category Revit calls "Generic Model" for these markers);
anything else selected is skipped and reported, since ordinary Revit
elements (walls, families, etc.) don't expose their geometry for direct
replacement the way DirectShape does.
"""
__title__ = "Scale Selected\nElements"
__version__ = "Version 1.0"
__author__ = "ADA"

from pyrevit import revit, DB, UI, forms, script
from System.Collections.Generic import List

# Shared ADA-Tools dark/gold themed report (see lib/GUI/ReportTheme.py) and
# the my_WPF base class (see lib/GUI/WPF_Base.py) that every other ADA-Tools
# popup window uses for its dark/gold look - inheriting from it and calling
# add_wpf_resource() pulls in lib/GUI/Resources/WPF_styles.xaml, which styles
# Button/CheckBox/TextBox/ListBox/ScrollBar implicitly (no XAML needed here).
from GUI.ReportTheme import ADAReport, THEME_BG, THEME_GOLD, THEME_GOLD_DARK, THEME_TEXT
from GUI.forms import my_WPF
from System.Windows import MessageBox
from System.Windows.Controls import Button, StackPanel, TextBlock, Grid, RowDefinition, ColumnDefinition, TextBox, DockPanel, Image
from System.Windows.Media import SolidColorBrush, ColorConverter, Stretch
from System import Windows

doc = revit.doc
uidoc = revit.uidoc

PRESET_PERCENTS = ["50", "100", "150", "200"]


def theme_brush(hex_color):
    return SolidColorBrush(ColorConverter.ConvertFromString(hex_color))


class ScaleDialog(my_WPF):
    """Small themed popup - one text box for the scale percentage, quick
    preset buttons, and Apply/Cancel - matching the rest of ADA-Tools."""

    def __init__(self, element_count):
        self.element_count = element_count
        self.scale_factor = None  # set on Apply; stays None on Cancel

        self.Title = "Scale Elements"
        self.Width = 340
        self.Height = 300
        self.WindowStartupLocation = Windows.WindowStartupLocation.CenterScreen
        self.ResizeMode = Windows.ResizeMode.NoResize
        self.WindowStyle = getattr(Windows.WindowStyle, "None")
        self.AllowsTransparency = True
        self.add_wpf_resource()
        self.Background = theme_brush(THEME_BG)

        self._create_ui()

    def _create_ui(self):
        main_grid = Grid()
        main_grid.Margin = Windows.Thickness(0)

        header_row = RowDefinition()
        header_row.Height = Windows.GridLength(25)
        body_row = RowDefinition()
        body_row.Height = Windows.GridLength(1, Windows.GridUnitType.Star)
        footer_row = RowDefinition()
        footer_row.Height = Windows.GridLength(25)
        main_grid.RowDefinitions.Add(header_row)
        main_grid.RowDefinitions.Add(body_row)
        main_grid.RowDefinitions.Add(footer_row)

        # --- Header ---------------------------------------------------
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

        logo_panel = DockPanel()
        logo_panel.VerticalAlignment = Windows.VerticalAlignment.Center
        logo_panel.HorizontalAlignment = Windows.HorizontalAlignment.Left

        logo_icon = Image()
        logo_icon.Source = self.load_logo_icon()
        logo_icon.Width = 18
        logo_icon.Height = 18
        logo_icon.Margin = Windows.Thickness(6, 0, 4, 0)
        logo_icon.Stretch = Stretch.Uniform
        logo_panel.Children.Add(logo_icon)

        logo_text = TextBlock()
        logo_text.Text = "ADA-Tools"
        logo_text.FontWeight = Windows.FontWeights.Heavy
        logo_text.FontSize = 12
        logo_text.VerticalAlignment = Windows.VerticalAlignment.Center
        logo_text.Foreground = theme_brush(THEME_TEXT)
        logo_panel.Children.Add(logo_text)

        Grid.SetColumn(logo_panel, 0)
        Grid.SetColumnSpan(logo_panel, 2)
        header_grid.Children.Add(logo_panel)

        title_text = TextBlock()
        title_text.Text = "Scale Elements"
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
        close_button.Click += self._on_cancel
        Grid.SetColumn(close_button, 2)
        header_grid.Children.Add(close_button)

        Grid.SetRow(header_grid, 0)
        main_grid.Children.Add(header_grid)

        # --- Body -------------------------------------------------------
        body_panel = StackPanel()
        body_panel.Margin = Windows.Thickness(20, 15, 20, 10)

        count_label = TextBlock()
        count_label.Text = "{} element(s) selected".format(self.element_count)
        count_label.Foreground = theme_brush(THEME_GOLD)
        count_label.FontSize = 12
        count_label.Margin = Windows.Thickness(0, 0, 0, 15)
        count_label.HorizontalAlignment = Windows.HorizontalAlignment.Center
        body_panel.Children.Add(count_label)

        factor_label = TextBlock()
        factor_label.Text = "Scale factor (% of current size):"
        factor_label.Foreground = theme_brush(THEME_TEXT)
        factor_label.FontSize = 12
        factor_label.Margin = Windows.Thickness(0, 0, 0, 5)
        body_panel.Children.Add(factor_label)

        self.value_box = TextBox()
        self.value_box.Text = "100"
        self.value_box.Height = 30
        self.value_box.FontSize = 16
        self.value_box.TextAlignment = Windows.TextAlignment.Center
        self.value_box.Margin = Windows.Thickness(0, 0, 0, 12)
        body_panel.Children.Add(self.value_box)

        preset_panel = StackPanel()
        preset_panel.Orientation = Windows.Controls.Orientation.Horizontal
        preset_panel.HorizontalAlignment = Windows.HorizontalAlignment.Center
        for pct in PRESET_PERCENTS:
            preset_button = Button()
            preset_button.Content = pct + "%"
            preset_button.Width = 60
            preset_button.Height = 26
            preset_button.FontSize = 11
            preset_button.Margin = Windows.Thickness(3, 0, 3, 0)
            preset_button.Click += self._make_preset_handler(pct)
            preset_panel.Children.Add(preset_button)
        body_panel.Children.Add(preset_panel)

        action_panel = StackPanel()
        action_panel.Orientation = Windows.Controls.Orientation.Horizontal
        action_panel.HorizontalAlignment = Windows.HorizontalAlignment.Center
        action_panel.Margin = Windows.Thickness(0, 20, 0, 0)

        apply_button = Button()
        apply_button.Content = "Apply"
        apply_button.Width = 100
        apply_button.Height = 30
        apply_button.Margin = Windows.Thickness(5, 0, 5, 0)
        apply_button.Click += self._on_apply
        action_panel.Children.Add(apply_button)

        cancel_button = Button()
        cancel_button.Content = "Cancel"
        cancel_button.Width = 100
        cancel_button.Height = 30
        cancel_button.Margin = Windows.Thickness(5, 0, 5, 0)
        cancel_button.Click += self._on_cancel
        action_panel.Children.Add(cancel_button)

        body_panel.Children.Add(action_panel)

        Grid.SetRow(body_panel, 1)
        main_grid.Children.Add(body_panel)

        # --- Footer -----------------------------------------------------
        footer_grid = Grid()
        footer_grid.Background = theme_brush(THEME_BG)

        footer_text = TextBlock()
        footer_text.Text = "ADA-Tools - Scale Elements"
        footer_text.Foreground = theme_brush(THEME_GOLD_DARK)
        footer_text.FontSize = 10
        footer_text.HorizontalAlignment = Windows.HorizontalAlignment.Center
        footer_text.VerticalAlignment = Windows.VerticalAlignment.Center
        footer_grid.Children.Add(footer_text)

        Grid.SetRow(footer_grid, 2)
        main_grid.Children.Add(footer_grid)

        self.Content = main_grid

    def _make_preset_handler(self, pct):
        def handler(sender, args):
            self.value_box.Text = pct
        return handler

    def _on_apply(self, sender, args):
        text = self.value_box.Text.strip().replace("%", "")
        value = None
        try:
            value = float(text)
        except Exception:
            value = None
        if value is None or value <= 0:
            MessageBox.Show("Enter a positive number, e.g. 150 for 150%.", "Scale Elements")
            return
        self.scale_factor = value / 100.0
        self.DialogResult = True
        self.Close()

    def _on_cancel(self, sender, args):
        self.DialogResult = False
        self.Close()


def bbox_center(bbox):
    return DB.XYZ(
        (bbox.Min.X + bbox.Max.X) / 2.0,
        (bbox.Min.Y + bbox.Max.Y) / 2.0,
        (bbox.Min.Z + bbox.Max.Z) / 2.0,
    )


def scale_transform_about(pivot, factor):
    """Uniform-scale transform that keeps `pivot` fixed:
    new_point = pivot + factor * (point - pivot)."""
    to_origin = DB.Transform.CreateTranslation(pivot.Negate())
    scaled = to_origin.ScaleBasisAndOrigin(factor)
    return DB.Transform.CreateTranslation(pivot).Multiply(scaled)


try:
    selected_ids = list(uidoc.Selection.GetElementIds())
    if not selected_ids:
        refs = uidoc.Selection.PickObjects(
            UI.Selection.ObjectType.Element, "Select elements to scale")
        selected_ids = [r.ElementId for r in refs]

    if not selected_ids:
        forms.alert("No elements selected.", exitscript=True)

    elements = [doc.GetElement(eid) for eid in selected_ids]

    dialog = ScaleDialog(len(elements))
    dialog.ShowDialog()
    if not dialog.scale_factor:
        script.exit()

    factor = dialog.scale_factor
    options = DB.Options()
    options.DetailLevel = DB.ViewDetailLevel.Fine

    report = ADAReport(__title__.replace(chr(10), " "))
    rows = []
    scaled_count = 0
    skipped_count = 0

    with revit.Transaction("Scale Elements"):
        for el in elements:
            id_link = report.link(el.Id, title=str(el.Id.IntegerValue))
            category_name = el.Category.Name if el.Category else "?"

            if not isinstance(el, DB.DirectShape):
                rows.append([id_link, category_name, "Skipped - not a DirectShape / Generic Model element"])
                skipped_count += 1
                continue

            try:
                bbox = el.get_BoundingBox(None)
                if bbox is None:
                    raise Exception("no bounding box")
                pivot = bbox_center(bbox)
                transform = scale_transform_about(pivot, factor)

                geom = el.get_Geometry(options)
                if geom is None:
                    raise Exception("no geometry")
                transformed = geom.GetTransformed(transform)
                new_objects = List[DB.GeometryObject](list(transformed))
                el.SetShape(new_objects)

                scaled_count += 1
                rows.append([id_link, category_name, "Scaled to {:.0f}%".format(factor * 100)])
            except Exception as ex:
                skipped_count += 1
                rows.append([id_link, category_name, "Failed - {}".format(ex)])

    uidoc.RefreshActiveView()

    report.table(["Element ID", "Category", "Result"], rows)
    if scaled_count:
        report.success("Scaled {} element(s) to {:.0f}% of their current size.".format(
            scaled_count, factor * 100))
    if skipped_count:
        report.warn("{} element(s) skipped (not scalable).".format(skipped_count))
    report.flush()

except Exception as e:
    report = ADAReport(__title__.replace(chr(10), " "))
    report.error("Error: {}".format(e))
    report.flush()
    import traceback
    print(traceback.format_exc())
