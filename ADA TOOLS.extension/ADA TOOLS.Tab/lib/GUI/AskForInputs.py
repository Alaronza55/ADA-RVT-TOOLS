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
clr.AddReference("System.Windows.Forms")
clr.AddReference("System")
from System.Windows import Visibility
import wpf

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
#====================================================================================================
PATH_SCRIPT = os.path.dirname(__file__)

MAX_FIELDS = 4

# ╔═╗╦  ╔═╗╔═╗╔═╗╔═╗╔═╗
# ║  ║  ╠═╣╚═╗╚═╗║╣ ╚═╗
# ╚═╝╩═╝╩ ╩╚═╝╚═╝╚═╝╚═╝ CLASSES
#====================================================================================================
class AskForInputs(my_WPF):
    """Small ADA-Tools styled popup with one labelled TextBox per field -
    the themed replacement for pyRevit's plain forms.ask_for_string, and
    it can ask several values in a single window (up to MAX_FIELDS)."""

    def __init__(self, fields,
                 title       = '__title__',
                 label       = "Enter values:",
                 button_name = 'OK',
                 version     = 'Version: 1.0'):
        if len(fields) > MAX_FIELDS:
            raise ValueError(
                "AskForInputs supports at most {} fields, got {}.".format(
                    MAX_FIELDS, len(fields)))

        self.fields = list(fields)
        self.values = None

        #>>>>>>>>>> SET RESOURCES FOR WPF
        self.add_wpf_resource()
        path_xaml_file = os.path.join(PATH_SCRIPT, 'AskForInputs.xaml')
        wpf.LoadComponent(self, path_xaml_file)

        # UPDATE GUI ELEMENTS
        self.logo_icon.Source     = self.load_logo_icon()
        self.main_title.Text      = title
        self.text_label.Content   = label
        self.button_main.Content  = button_name
        self.footer_version.Text  = version

        self.input_boxes = [self.field_input_0, self.field_input_1,
                            self.field_input_2, self.field_input_3]
        rows       = [self.row_0, self.row_1, self.row_2, self.row_3]
        row_labels = [self.field_label_0, self.field_label_1,
                      self.field_label_2, self.field_label_3]

        for i, row in enumerate(rows):
            if i < len(self.fields):
                field_label, field_default    = self.fields[i]
                row_labels[i].Text            = field_label
                self.input_boxes[i].Text      = field_default or ''
            else:
                row.Visibility = Visibility.Collapsed

        self.ShowDialog()

    #>>>>>>>>>> INHERIT WPF RESOURCES
    def add_wpf_resource(self):
        """Function to get resources from super()"""
        super(AskForInputs, self).add_wpf_resource()

    # ╔═╗╦ ╦╦  ╔═╗╦  ╦╔═╗╔╗╔╔╦╗╔═╗
    # ║ ╦║ ║║  ║╣ ╚╗╔╝║╣ ║║║ ║ ╚═╗
    # ╚═╝╚═╝╩  ╚═╝ ╚╝ ╚═╝╝╚╝ ╩ ╚═╝ GUI EVENTS
    #==================================================
    def window_loaded(self, sender, e):
        """Put the caret in the first field so the user can type right away."""
        if self.fields:
            self.field_input_0.Focus()
            self.field_input_0.SelectAll()

    def button_ok_click(self, sender, e):
        """OK (or Enter) - collect the field values and close."""
        self.values = [self.input_boxes[i].Text
                       for i in range(len(self.fields))]
        self.Close()


# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝MAIN
#====================================================================================================
def ask_for_inputs(fields,
                   title       = '__title__',
                   label       = "Enter values:",
                   button_name = 'OK',
                   version     = 'Version: 1.0'):
    #type:(list, str, str, str, str) -> list
    """Function to present a small themed text-input popup to the user -
    one labelled TextBox per field, all asked in a single window (up to 4).
    :param fields:      list of (label, default_value) tuples, one per field.
    :param title:       Title of the window.
    :param label:       Label displayed above the fields.
    :param button_name: Text on the confirm button.
    :param version:     Version of the script for footer.
    :return:            List of the typed strings (same order as fields),
                        or None if the window was closed without confirming."""

    GUI_inputs = AskForInputs(fields,
                              title       = title,
                              label       = label,
                              button_name = button_name,
                              version     = version)
    return GUI_inputs.values
