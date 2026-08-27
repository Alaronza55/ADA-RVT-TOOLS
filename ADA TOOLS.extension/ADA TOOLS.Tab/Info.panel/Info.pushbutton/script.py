__doc__ = """Hello, this extension aims to help BIM COORDINATORS in their job. You will
find here some precious and timesaving functions to optimize your work.
Hope you find them useful. Don't hesitate to report any bugs or ask for
some new features here : dalmog.dav@gmail.com"""
__title__ = "Info"
__author__ = "ADA"

from pyrevit import forms

forms.alert(
    "Hello, this extension aims to help BIM COORDINATORS in their job.\n\n"
    "You will find here some precious and timesaving functions to optimize "
    "your work. Hope you find them usefull.\n\n"
    "Don't hesitate to report any bugs or ask for some new features here:\n"
    "dalmog.dav@gmail.com",
    title="ADA Tools"
)
