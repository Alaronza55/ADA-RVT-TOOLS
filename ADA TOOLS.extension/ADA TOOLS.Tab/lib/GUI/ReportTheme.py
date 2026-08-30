# -*- coding: utf-8 -*-
"""Shared ADA-Tools dark/gold themed report for pyRevit's output console -
same palette as lib/GUI/Resources/WPF_styles.xaml (dark background, gold
text/accents), with the ADA.Tools logo banner printed automatically above
the rest of the report when it is flushed.

Usage:
    from GUI.ReportTheme import ADAReport

    report = ADAReport(__title__)
    report.line("Some info")
    report.table(["Col 1", "Col 2"], [["a", "1"], ["b", "2"]])
    report.warn("Something to double-check")
    report.success("Done.")
    report.flush()

Every call returns self, so calls can be chained:
    report.line("...").warn("...").flush()

The console page itself is recolored dark on flush() (pyRevit's output
window is light by default), so build the whole report through one
ADAReport instance and flush() once at the end rather than printing
directly in between - intermediate plain print()/forms.alert() calls in
between are fine, they just won't be part of the themed panel.
"""
import os
import base64

from pyrevit import script

output = script.get_output()

# --- Palette - keep in sync with lib/GUI/Resources/WPF_styles.xaml ---------
THEME_BG = "#1E1E1E"           # header_background
THEME_ROW_BG = "#171512"       # alternating table row shade
THEME_GOLD = "#f6c955"         # text_magenta / text_gray / border_magenta
THEME_GOLD_LIGHT = "#fbe8a0"   # light-yellow chip background (e.g. linkify links)
THEME_GOLD_DARK = "#c5a144"    # button_bg_normal
THEME_GOLD_HOVER = "#ae8e3c"   # button_bg_hover
THEME_TEXT = "#E5E4E2"         # text_white
THEME_FONT = "'Arial Narrow', Arial, sans-serif"

_LOGO_B64 = None  # cached across calls within one script run


def _load_logo_base64():
    global _LOGO_B64
    if _LOGO_B64 is not None:
        return _LOGO_B64
    try:
        import GUI
        logo_path = os.path.join(os.path.dirname(GUI.__file__), "Resources", "ADA_Tools_Logo.png")
        with open(logo_path, "rb") as f:
            _LOGO_B64 = base64.b64encode(f.read()).decode("ascii")
    except Exception:
        _LOGO_B64 = ""
    return _LOGO_B64


def html_header(text):
    return (
        '<div style="color:{gold}; font-family:{font}; font-size:16px; font-weight:bold;">'
        '{text}</div>'
        '<hr style="border:none; border-top:2px solid {gold}; margin:6px 0 12px 0;">'
    ).format(gold=THEME_GOLD, font=THEME_FONT, text=text)


def html_subheader(text):
    return (
        '<div style="color:{gold}; font-family:{font}; font-size:13px; '
        'font-weight:bold; margin:10px 0 4px 0;">{text}</div>'
    ).format(gold=THEME_GOLD, font=THEME_FONT, text=text)


def html_line(text):
    return '<div style="color:{gold}; font-family:{font};">{text}</div>'.format(
        gold=THEME_GOLD, font=THEME_FONT, text=text)


def html_warn(text):
    return (
        '<div style="color:{gold_dark}; font-family:{font};">&#9888; {text}</div>'
    ).format(gold_dark=THEME_GOLD_DARK, font=THEME_FONT, text=text)


def html_error(text):
    return (
        '<div style="color:#c0392b; font-family:{font}; font-weight:bold;">'
        '&#10060; {text}</div>'
    ).format(font=THEME_FONT, text=text)


def html_success(text):
    return (
        '<div style="color:{gold}; font-family:{font}; font-weight:bold;">'
        '&#10004; {text}</div>'
    ).format(gold=THEME_GOLD, font=THEME_FONT, text=text)


def html_table(columns, rows):
    header_cells = ''.join(
        '<th style="background-color:{bg} !important; color:{gold} !important; '
        'text-align:left; padding:4px 10px; border:1px solid {gold_dark};">{c}</th>'.format(
            bg=THEME_BG, gold=THEME_GOLD, gold_dark=THEME_GOLD_DARK, c=c)
        for c in columns)

    body_rows = []
    for i, row in enumerate(rows):
        row_bg = THEME_BG if i % 2 == 0 else THEME_ROW_BG
        cells = ''.join(
            '<td style="background-color:{bg} !important; color:{text} !important; '
            'padding:4px 10px; border:1px solid {gold_dark};">{v}</td>'.format(
                bg=row_bg, text=THEME_TEXT, gold_dark=THEME_GOLD_DARK, v=v)
            for v in row)
        body_rows.append('<tr>{cells}</tr>'.format(cells=cells))

    return (
        '<table style="border-collapse:collapse; font-family:{font}; font-size:12px; '
        'margin:6px 0 12px 0;"><tr>{header}</tr>{body}</table>'
    ).format(font=THEME_FONT, header=header_cells, body=''.join(body_rows))


def html_logo_header():
    """ADA.Tools logo banner + divider, printed once at the top of the
    report, above everything else."""
    b64 = _load_logo_base64()
    if not b64:
        return ""
    return (
        '<div style="margin-bottom:14px;">'
        '<img src="data:image/png;base64,{b64}" alt="ADA.Tools - BIM Coordination &amp; '
        'Auditing Tools" style="max-width:420px; width:100%; height:auto; display:block;">'
        '<hr style="border:none; border-top:2px solid {gold_dark}; margin:10px 0 0 0;">'
        '</div>'
    ).format(gold_dark=THEME_GOLD_DARK, b64=b64)


class ADAReport(object):
    """Accumulates themed HTML fragments and flushes them as one dark/gold
    panel - ADA.Tools logo banner on top, then the report content below -
    in pyRevit's output console."""

    def __init__(self, title=None):
        self.fragments = []
        if title:
            self.header(title)

    def header(self, text):
        self.fragments.append(html_header(text))
        return self

    def subheader(self, text):
        self.fragments.append(html_subheader(text))
        return self

    def line(self, text):
        self.fragments.append(html_line(text))
        return self

    def warn(self, text):
        self.fragments.append(html_warn(text))
        return self

    def error(self, text):
        self.fragments.append(html_error(text))
        return self

    def success(self, text):
        self.fragments.append(html_success(text))
        return self

    def table(self, columns, rows):
        self.fragments.append(html_table(columns, rows))
        return self

    def raw(self, html_fragment):
        """Escape hatch - append a hand-built HTML fragment as-is."""
        self.fragments.append(html_fragment)
        return self

    def link(self, element_id, title=None):
        """pyRevit's own clickable "select in view" link for element_id,
        as an html fragment (use inside a table cell or a line)."""
        return output.linkify(element_id, title=title if title is not None else str(element_id))

    def flush(self):
        """Print everything accumulated so far as one themed panel - the
        ADA.Tools logo banner and a divider first, then the report content
        below it - then clear the buffer (safe to call flush() more than
        once per script run)."""
        if not self.fragments:
            return
        content = ''.join(self.fragments)
        output.print_html(
            '<style>'
            'html, body {{ background-color:{bg} !important; font-family:{font}; }}'
            'a, a:visited {{ background-color:{gold_light} !important; color:{bg} !important; '
            'text-decoration:none !important; }}'
            'a:hover, a:visited:hover {{ background-color:{gold_dark} !important; '
            'color:#ffffff !important; }}'
            '</style>'
            '<div style="padding:10px 4px;">{logo}{content}</div>'.format(
                bg=THEME_BG, font=THEME_FONT, gold_light=THEME_GOLD_LIGHT, gold_dark=THEME_GOLD_DARK,
                logo=html_logo_header(), content=content))
        self.fragments = []
