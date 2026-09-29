# -*- coding: utf-8 -*-

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
#====================================================================================================
import os

#>>>>>>>>>> pyRevit
from pyrevit import forms  # Needed for wpf import to work.

# Custom Imports
from GUI.forms import my_WPF

#>>>>>>>>>> .NET IMPORTS
import clr
clr.AddReference("System")
from System.Collections.Generic import List
import wpf

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#====================================================================================================
PATH_SCRIPT = os.path.dirname(__file__)


class OrderItem:
    """Helper Class for displaying items in the order ListBox."""
    def __init__(self, Name='Unnamed', element=None):
        self.Name    = Name
        self.element = element

# ╔═╗╦  ╔═╗╔═╗╔═╗╔═╗╔═╗
# ║  ║  ╠═╣╚═╗╚═╗║╣ ╚═╗
# ╚═╝╩═╝╩ ╩╚═╝╚═╝╚═╝╚═╝ CLASSES
#====================================================================================================
class OrderList(my_WPF):
    """ADA-Tools styled window to put a list of items in order - select an
    item and move it with Up / Down / To Top / To Bottom."""

    def __init__(self, items,
                 title       = '__title__',
                 label       = "Set the order:",
                 button_name = 'OK',
                 top_text    = 'TOP',
                 bottom_text = 'BOTTOM',
                 version     = 'Version: 1.0'):
        self.items   = [OrderItem(name, element) for name, element in items]
        self.ordered = None

        #>>>>>>>>>> SET RESOURCES FOR WPF
        self.add_wpf_resource()
        path_xaml_file = os.path.join(PATH_SCRIPT, 'OrderList.xaml')
        wpf.LoadComponent(self, path_xaml_file)

        # UPDATE GUI ELEMENTS
        self.logo_icon.Source     = self.load_logo_icon()
        self.main_title.Text      = title
        self.text_label.Content   = label
        self.button_main.Content  = button_name
        self.text_top.Text        = top_text
        self.text_bottom.Text     = bottom_text
        self.footer_version.Text  = version

        self.refresh(0)
        self.ShowDialog()

    #>>>>>>>>>> INHERIT WPF RESOURCES
    def add_wpf_resource(self):
        """Function to get resources from super()"""
        super(OrderList, self).add_wpf_resource()

    def refresh(self, selected_index):
        """Push self.items to the ListBox and keep the moved item selected."""
        list_of_items = List[type(OrderItem())]()
        for item in self.items:
            list_of_items.Add(item)
        self.main_ListBox.ItemsSource = list_of_items
        if self.items:
            self.main_ListBox.SelectedIndex = selected_index
            self.main_ListBox.ScrollIntoView(self.main_ListBox.SelectedItem)

    def move_selected(self, new_index_func):
        index = self.main_ListBox.SelectedIndex
        if index < 0:
            return
        new_index = max(0, min(len(self.items) - 1, new_index_func(index)))
        if new_index == index:
            return
        item = self.items.pop(index)
        self.items.insert(new_index, item)
        self.refresh(new_index)

    # ╔╗ ╦ ╦╔╦╗╔╦╗╔═╗╔╗╔╔═╗
    # ╠╩╗║ ║ ║  ║ ║ ║║║║╚═╗
    # ╚═╝╚═╝ ╩  ╩ ╚═╝╝╚╝╚═╝ BUTTONS
    #==================================================
    def button_up(self, sender, e):
        self.move_selected(lambda i: i - 1)

    def button_down(self, sender, e):
        self.move_selected(lambda i: i + 1)

    def button_to_top(self, sender, e):
        self.move_selected(lambda i: 0)

    def button_to_bottom(self, sender, e):
        self.move_selected(lambda i: len(self.items) - 1)

    def button_ok(self, sender, e):
        """Confirm - keep the order as shown (first = top of the list)."""
        self.ordered = [item.element for item in self.items]
        self.Close()


# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝MAIN
#====================================================================================================
def order_list(items,
               title       = '__title__',
               label       = "Set the order:",
               button_name = 'OK',
               top_text    = 'TOP',
               bottom_text = 'BOTTOM',
               version     = 'Version: 1.0'):
    #type:(list, str, str, str, str, str, str) -> list
    """Function to let the user put items in order in a themed window.
    :param items:       list of (display_name, value) tuples, in the
                        initial order (first = top of the list).
    :param title:       Title of the window.
    :param label:       Label displayed above the list.
    :param button_name: Text on the confirm button.
    :param top_text:    Caption above the list (what the first item means).
    :param bottom_text: Caption below the list (what the last item means).
    :param version:     Version of the script for footer.
    :return:            List of the values in the chosen order (first =
                        top of the list), or None if the window was closed
                        without confirming."""

    GUI_order = OrderList(items,
                          title       = title,
                          label       = label,
                          button_name = button_name,
                          top_text    = top_text,
                          bottom_text = bottom_text,
                          version     = version)
    return GUI_order.ordered
