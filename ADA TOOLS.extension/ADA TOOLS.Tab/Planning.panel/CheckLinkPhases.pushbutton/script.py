# -*- coding: utf-8 -*-
__doc__ = """Compare the phases of this project with the phases of a Revit link,
character by character - to find out why Revit does not map them.

Revit maps link phases BY NAME only when the link is first loaded, then
keeps that mapping (a reload does not redo it; phases added later are
mapped to the link's last phase). To redo the automatic mapping: Manage
Links > Remove the link, then link it again.

If the names differ - even by an invisible character - Revit cannot map
them. The report shows, for every phase of this project, the phase with
the same name in the link (or the closest one) and points out the exact
difference: extra / double spaces, non-breaking spaces, different
letters, accents... Nothing in the model is changed."""
__title__ = "Check Link\nPhases"
__version__ = "Version 1.0"
__author__ = "ADA"

import clr

clr.AddReference('RevitAPI')

from Autodesk.Revit.DB import FilteredElementCollector, RevitLinkInstance

from pyrevit import forms, revit, script

# Custom ADA GUI - themed list (lib/GUI/SelectFromDict.py) and the shared
# dark/gold themed report (lib/GUI/ReportTheme.py)
from GUI.forms import select_from_dict
from GUI.ReportTheme import ADAReport
from Snippets._planning import name_key

doc = revit.doc
TITLE = __title__.replace("\n", " ")


def eid_value(eid):
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue


def visible(text):
    """Show what is normally invisible: spaces as '·', other unusual
    characters as <U+XXXX>, leading / trailing spaces kept visible."""
    out = []
    for ch in text:
        code = ord(ch)
        if ch == u" ":
            out.append(u"·")
        elif code < 32 or code in (0xA0, 0x2007, 0x202F, 0x200B, 0xFEFF) or \
                0x2000 <= code <= 0x200F:
            out.append(u"<U+{:04X}>".format(code))
        else:
            out.append(ch)
    return u"".join(out)


def first_difference(a, b):
    """Human description of the first difference between two strings."""
    for i, (ca, cb) in enumerate(zip(a, b)):
        if ca != cb:
            return u"char {}: '{}' (U+{:04X}) vs '{}' (U+{:04X})".format(
                i + 1, visible(ca), ord(ca), visible(cb), ord(cb))
    if len(a) != len(b):
        longer = a if len(a) > len(b) else b
        return u"length {} vs {} - extra '{}' at the end".format(
            len(a), len(b), visible(longer[min(len(a), len(b)):]))
    return u""


def main():
    links = [l for l in FilteredElementCollector(doc).OfClass(RevitLinkInstance)
             if l.GetLinkDocument() is not None]
    if not links:
        forms.alert("There is no loaded Revit link.", title=TITLE, exitscript=True)
    chosen = select_from_dict({u"{}  [{}]".format(l.Name, eid_value(l.Id)): l
                               for l in links},
                              title=TITLE, label="Link to compare:",
                              button_name="Compare", version=__version__,
                              SelectMultiple=False)
    if not chosen:
        script.exit()
    link_doc = chosen[0].GetLinkDocument()

    host = [p.Name for p in doc.Phases]
    linked = [p.Name for p in link_doc.Phases]
    linked_exact = set(linked)
    linked_by_key = {}
    for n in linked:
        linked_by_key.setdefault(name_key(n), n)

    rows = []
    exact = differs = absent = 0
    for i, name in enumerate(host):
        same_pos = linked[i] if i < len(linked) else u"-"
        if name in linked_exact:
            exact += 1
            rows.append([str(i), name, u"identical", u""])
            continue
        near = linked_by_key.get(name_key(name))
        if near is None and same_pos != u"-":
            near = same_pos
        if near is not None and name_key(near) == name_key(name):
            differs += 1
            rows.append([str(i), visible(name), u"DIFFERENT: " + visible(near),
                         first_difference(name, near)])
        else:
            absent += 1
            rows.append([str(i), name, u"NOT IN LINK", u"same position in link: " +
                         (same_pos if same_pos != u"-" else u"none")])

    only_link = [n for n in linked if n not in set(host)]

    report = ADAReport(TITLE)
    report.line(u"Project: <b>{}</b> phase(s) &nbsp;|&nbsp; link "
                u"<b>{}</b>: <b>{}</b> phase(s)".format(
                    len(host), link_doc.Title, len(linked)))
    report.table(["#", "Phase in this project", "In the link", "Difference"], rows)
    if only_link:
        report.subheader("Phases only in the link")
        report.table(["Phase in the link"], [[visible(n)] for n in only_link])

    if differs or absent:
        report.warn(u"{} phase(s) with a different name, {} missing in the link "
                    u"- Revit cannot map those automatically. '·' marks a space."
                    .format(differs, absent))
    else:
        report.success(u"All {} phase names are IDENTICAL. Revit will map them "
                       u"automatically when the link is loaded the first time: "
                       u"Manage Links > Remove the link, then link it again "
                       u"(a reload keeps the old mapping).".format(exact))
    report.flush()


if __name__ == "__main__":
    main()
