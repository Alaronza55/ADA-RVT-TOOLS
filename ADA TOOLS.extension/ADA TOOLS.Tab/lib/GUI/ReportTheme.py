# -*- coding: utf-8 -*-
"""Shared ADA-Tools dark/gold themed report for pyRevit's output console -
same palette as lib/GUI/Resources/WPF_styles.xaml (dark background, gold
text/accents), with the ADA.Tools signature appended automatically when
the report is flushed.

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
THEME_BG = "#0b0b09"           # header_background
THEME_ROW_BG = "#171512"       # window background (used for alternating rows)
THEME_GOLD = "#f6c955"         # text_magenta / text_gray / border_magenta
THEME_GOLD_LIGHT = "#fbe8a0"   # light-yellow chip background (e.g. linkify links)
THEME_GOLD_DARK = "#c5a144"    # button_bg_normal
THEME_GOLD_HOVER = "#ae8e3c"   # button_bg_hover
THEME_TEXT = "#E5E4E2"         # text_white
THEME_FONT = "Segoe UI, Arial, sans-serif"

_SIGNATURE_B64 = None  # cached across calls within one script run


def _load_signature_base64():
    global _SIGNATURE_B64
    if _SIGNATURE_B64 is not None:
        return _SIGNATURE_B64
    try:
        import GUI
        sig_path = os.path.join(os.path.dirname(GUI.__file__), "Resources", "ADA_Tools_Signature.png")
        with open(sig_path, "rb") as f:
            _SIGNATURE_B64 = base64.b64encode(f.read()).decode("ascii")
    except Exception:
        _SIGNATURE_B64 = ""
    return _SIGNATURE_B64


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


def html_signature():
    b64 = _load_signature_base64()
    if not b64:
        return ""
    return (
        '<div style="margin-top:20px; padding-top:12px; border-top:1px solid {gold_dark};">'
        '<img src="data:image/png;base64,{b64}" alt="ADA.Tools - BIM Coordination &amp; '
        'Auditing Tools" style="max-width:280px; height:auto; display:block;"></div>'
    ).format(gold_dark=THEME_GOLD_DARK, b64=b64)


class ADAReport(object):
    """Accumulates themed HTML fragments and flushes them as one dark/gold
    panel (plus the ADA.Tools signature) in pyRevit's output console."""

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
        """Print everything accumulated so far as one themed panel with
        the ADA.Tools signature appended, then clear the buffer (safe to
        call flush() more than once per script run)."""
        if not self.fragments:
            return
        self.fragments.append(html_signature())
        output.print_html(
            '<style>'
            'html, body {{ background-color:{bg} !important; }}'
            'a, a:visited {{ background-color:{gold_light} !important; color:{bg} !important; '
            'text-decoration:none !important; }}'
            'a:hover, a:visited:hover {{ background-color:{gold_dark} !important; '
            'color:#ffffff !important; }}'
            '</style>'
            '<div style="padding:10px 4px;">{content}</div>'.format(
                bg=THEME_BG, gold_light=THEME_GOLD_LIGHT, gold_dark=THEME_GOLD_DARK,
                content=''.join(self.fragments)))
        self.fragments = []
