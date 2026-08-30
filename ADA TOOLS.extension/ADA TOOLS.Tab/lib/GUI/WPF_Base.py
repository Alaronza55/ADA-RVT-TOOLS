# -*- coding: utf-8 -*-
__title__ = "Template for my custom WPF windows."
__author__ = "ADA"

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
# ==================================================
#>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>> .NET IMPORTS
import os, sys
from pyrevit import revit, forms

import os, clr
clr.AddReference("System")

from System.Diagnostics.Process import Start
from System.Windows.Window import DragMove
from System.Windows.Input import MouseButtonState
import wpf
from System.Windows import Application, Window, ResourceDictionary
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
from System import Uri

# ╦ ╦╔═╗╔═╗  ╔╦╗╔═╗╔╦╗╔═╗╦  ╔═╗╔╦╗╔═╗
# ║║║╠═╝╠╣    ║ ║╣ ║║║╠═╝║  ╠═╣ ║ ║╣
# ╚╩╝╩  ╚     ╩ ╚═╝╩ ╩╩  ╩═╝╩ ╩ ╩ ╚═╝ WPF TEMPLATE
# ==================================================
class my_WPF(Window):

    # ╔╦╗╔═╗╔╦╗╦ ╦╔═╗╔╦╗╔═╗
    # ║║║║╣  ║ ╠═╣║ ║ ║║╚═╗
    # ╩ ╩╚═╝ ╩ ╩ ╩╚═╝═╩╝╚═╝ METHODS
    #==================================================
    def add_wpf_resource(self):
        """Function to add WPF resources."""
        dir_path        = os.path.dirname(__file__)
        path_styles     = os.path.join(dir_path, 'Resources/WPF_styles.xaml')
        r               = ResourceDictionary()
        r.Source        = Uri(path_styles)
        self.Resources  = r

    def load_logo_icon(self):
        """BitmapImage for the ADA.Tools icon glyph (lib/GUI/Resources/
        ADA_Tools_Icon.png) - for the small Image in every window's
        header, next to the "ADA-Tools" link."""
        dir_path  = os.path.dirname(__file__)
        icon_path = os.path.join(dir_path, 'Resources', 'ADA_Tools_Icon.png')
        bmp = BitmapImage()
        bmp.BeginInit()
        bmp.UriSource   = Uri(icon_path)
        bmp.CacheOption = BitmapCacheOption.OnLoad
        bmp.EndInit()
        return bmp

    # ╔═╗╦ ╦╦  ╔═╗╦  ╦╔═╗╔╗╔╔╦╗╔═╗
    # ║ ╦║ ║║  ║╣ ╚╗╔╝║╣ ║║║ ║ ╚═╗
    # ╚═╝╚═╝╩  ╚═╝ ╚╝ ╚═╝╝╚╝ ╩ ╚═╝ GUI EVENTS
    #==================================================
    def button_close(self, sender, e):
        """Stop application by clicking on a <Close> button in the top right corner."""
        self.Close()

    def header_drag(self, sender, e):
        """Drag window by holding LeftButton on the header."""
        if e.LeftButton == MouseButtonState.Pressed:
            DragMove(self)

    def Hyperlink_RequestNavigate(self, sender, e):
        """Forwarding for a Hyperlinks."""
        Start(e.Uri.AbsoluteUri)


