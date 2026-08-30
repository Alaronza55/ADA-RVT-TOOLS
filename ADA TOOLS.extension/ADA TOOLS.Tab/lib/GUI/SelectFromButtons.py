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

MAX_BUTTONS = 4

# ╔═╗╦  ╔═╗╔═╗╔═╗╔═╗╔═╗
# ║  ║  ╠═╣╚═╗╚═╗║╣ ╚═╗
# ╚═╝╩═╝╩ ╩╚═╝╚═╝╚═╝╚═╝ CLASSES
#====================================================================================================
class SelectFromButtons(my_WPF):
    """Small ADA-Tools styled popup with one button per option - click a
    button to pick it, instead of checking an item in a list and pressing
    a separate Select button. Supports up to MAX_BUTTONS options."""

    def __init__(self, options,
                 title   = '__title__',
                 label   = "Choose an option:",
                 version = 'Version: 1.0'):
        if len(options) > MAX_BUTTONS:
            raise ValueError(
                "SelectFromButtons supports at most {} options, got {}.".format(
                    MAX_BUTTONS, len(options)))

        self.options          = list(options)
        self.selected_option  = None

        #>>>>>>>>>> SET RESOURCES FOR WPF
        self.add_wpf_resource()
        path_xaml_file = os.path.join(PATH_SCRIPT, 'SelectFromButtons.xaml')
        wpf.LoadComponent(self, path_xaml_file)

        # UPDATE GUI ELEMENTS
        self.logo_icon.Source    = self.load_logo_icon()
        self.main_title.Text     = title
        self.text_label.Content  = label
        self.footer_version.Text = version

        option_buttons = [self.btn_option_0, self.btn_option_1,
                           self.btn_option_2, self.btn_option_3]
        for i, btn in enumerate(option_buttons):
            if i < len(self.options):
                btn.Content    = self.options[i]
                btn.Visibility = Visibility.Visible
            else:
                btn.Visibility = Visibility.Collapsed

        self.ShowDialog()

    #>>>>>>>>>> INHERIT WPF RESOURCES
    def add_wpf_resource(self):
        """Function to get resources from super()"""
        super(SelectFromButtons, self).add_wpf_resource()

    # ╔═╗╦ ╦╦  ╔═╗╦  ╦╔═╗╔╗╔╔╦╗╔═╗
    # ║ ╦║ ║║  ║╣ ╚╗╔╝║╣ ║║║ ║ ╚═╗
    # ╚═╝╚═╝╩  ╚═╝ ╚╝ ╚═╝╝╚╝ ╩ ╚═╝ GUI EVENTS
    #==================================================
    def button_option_click(self, sender, e):
        """Any option button was clicked - remember its label and close."""
        self.selected_option = sender.Content
        self.Close()


# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝MAIN
#====================================================================================================
def select_from_buttons(options,
                         title   = '__title__',
                         label   = "Choose an option:",
                         version = 'Version: 1.0'):
    #type:(any, str, str, str) -> any
    """Function to present a small button-choice popup to the user - one
    button per option, click to pick it (up to 4 options).
    :param options: dict {label: value} OR a list of labels (used as their
                    own value).
    :param title:   Title of the window.
    :param label:   Label displayed above the buttons.
    :param version: Version of the script for footer.
    :return:        The selected value, or None if the window was closed
                    without picking an option."""

    # CONVERT LIST TO DICT
    if isinstance(options, dict):
        options_dict = options
    else:
        options_dict = {o: o for o in options}

    GUI_select = SelectFromButtons(list(options_dict.keys()),
                                    title   = title,
                                    label   = label,
                                    version = version)

    if GUI_select.selected_option is None:
        return None
    return options_dict[GUI_select.selected_option]
