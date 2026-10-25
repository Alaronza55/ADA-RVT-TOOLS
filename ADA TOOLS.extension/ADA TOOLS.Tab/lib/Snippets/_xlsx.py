# -*- coding: utf-8 -*-
"""Minimal .xlsx / .xlsm reader for pyRevit (IronPython) - no Excel and no
third-party package needed. An .xlsx file is a zip of XML parts; this
reads the sheet names, shared strings and cell values with .NET only.

    wb = XlsxWorkbook(path)
    wb.sheet_names                      -> ['Planning', 'Sheet2']
    wb.read_range('Planning', 'B2:B20') -> ['Phase 1', '', 'Phase 2', ...]

Values are returned as text. Formulas give their last calculated value
(as saved by Excel). Dates come back as Excel serial numbers. The file
can stay open in Excel while it is read."""

import re

import clr
clr.AddReference('System.IO.Compression')
clr.AddReference('System.Xml')
try:
    clr.AddReference('System.Private.Xml')   # Revit 2025+ (.NET 8): real home of XmlDocument
except Exception:
    pass

from System.IO import FileAccess, FileMode, FileShare, FileStream, StreamReader
from System.IO.Compression import ZipArchive, ZipArchiveMode
from System.Xml import XmlDocument, XmlNamespaceManager

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


def column_index(letters):
    """'A' -> 1, 'Z' -> 26, 'AA' -> 27."""
    index = 0
    for ch in letters.upper():
        index = index * 26 + (ord(ch) - ord('A') + 1)
    return index


def parse_cell(ref):
    """'B12' -> (12, 2). Raises ValueError if not a cell reference."""
    match = _CELL_RE.match(ref.strip())
    if not match:
        raise ValueError("'{}' is not a cell reference".format(ref))
    return int(match.group(2)), column_index(match.group(1))


def parse_range(ref):
    """'B2:C10' -> ((2, 2), (10, 3)), corners sorted. A single cell is
    accepted as a one-cell range. A 'Sheet!' prefix is ignored."""
    ref = ref.strip()
    if "!" in ref:
        ref = ref.split("!")[-1]
    parts = ref.split(":")
    if len(parts) == 1:
        parts = [parts[0], parts[0]]
    if len(parts) != 2:
        raise ValueError("'{}' is not a range (e.g. A2:A20)".format(ref))
    (r1, c1), (r2, c2) = parse_cell(parts[0]), parse_cell(parts[1])
    return (min(r1, r2), min(c1, c2)), (max(r1, r2), max(c1, c2))


def _number_text(text):
    """'3.0' -> '3', other numbers unchanged."""
    try:
        value = float(text)
    except ValueError:
        return text
    if value == int(value) and abs(value) < 1e15:
        return str(int(value))
    return text


class XlsxWorkbook(object):

    def __init__(self, path):
        self.path = path
        self._parts = self._read_parts(path)
        self._shared = self._read_shared_strings()
        self._sheet_parts = self._read_sheet_parts()   # [(name, part)]
        self._cache = {}

    # ------------------------------------------------------------------
    @staticmethod
    def _read_parts(path):
        parts = {}
        stream = FileStream(path, FileMode.Open, FileAccess.Read,
                            FileShare.ReadWrite)
        try:
            archive = ZipArchive(stream, ZipArchiveMode.Read)
            for entry in archive.Entries:
                name = entry.FullName.replace("\\", "/").lstrip("/")
                if name.endswith(".xml") or name.endswith(".rels"):
                    reader = StreamReader(entry.Open())
                    parts[name] = reader.ReadToEnd()
                    reader.Close()
            archive.Dispose()
        finally:
            stream.Close()
        if "xl/workbook.xml" not in parts:
            raise ValueError("Not an .xlsx/.xlsm workbook: {}".format(path))
        return parts

    def _xml(self, part):
        xml = XmlDocument()
        xml.LoadXml(self._parts[part])
        ns = XmlNamespaceManager(xml.NameTable)
        ns.AddNamespace("m", NS_MAIN)
        ns.AddNamespace("r", NS_REL)
        ns.AddNamespace("p", NS_PKG)
        return xml, ns

    def _read_shared_strings(self):
        if "xl/sharedStrings.xml" not in self._parts:
            return []
        xml, ns = self._xml("xl/sharedStrings.xml")
        strings = []
        for si in xml.SelectNodes("//m:sst/m:si", ns):
            # plain text or rich-text runs (ignore phonetic runs)
            strings.append("".join(t.InnerText for t in
                                   si.SelectNodes("./m:t | ./m:r/m:t", ns)))
        return strings

    def _read_sheet_parts(self):
        rels_xml, rels_ns = self._xml("xl/_rels/workbook.xml.rels")
        targets = {}
        for rel in rels_xml.SelectNodes("//p:Relationship", rels_ns):
            target = rel.GetAttribute("Target")
            if target.startswith("/"):
                target = target.lstrip("/")
            else:
                target = "xl/" + target
            targets[rel.GetAttribute("Id")] = target

        xml, ns = self._xml("xl/workbook.xml")
        sheets = []
        for sheet in xml.SelectNodes("//m:sheets/m:sheet", ns):
            part = targets.get(sheet.GetAttribute("id", NS_REL))
            if part in self._parts:
                sheets.append((sheet.GetAttribute("name"), part))
        return sheets

    # ------------------------------------------------------------------
    @property
    def sheet_names(self):
        return [name for name, _ in self._sheet_parts]

    def _cells(self, sheet_name):
        """{(row, col): text} of a sheet (non-empty cells only)."""
        if sheet_name in self._cache:
            return self._cache[sheet_name]
        part = dict(self._sheet_parts).get(sheet_name)
        if part is None:
            raise ValueError("No sheet named '{}'".format(sheet_name))

        xml, ns = self._xml(part)
        cells = {}
        for row_index, row in enumerate(
                xml.SelectNodes("//m:sheetData/m:row", ns), 1):
            row_number = int(row.GetAttribute("r") or row_index)
            col = 0
            for c in row.SelectNodes("m:c", ns):
                ref = c.GetAttribute("r")
                if ref:
                    row_number, col = parse_cell(ref)
                else:
                    col += 1
                cell_type = c.GetAttribute("t")
                if cell_type == "inlineStr":
                    text = "".join(t.InnerText for t in
                                   c.SelectNodes("m:is//m:t", ns))
                else:
                    v = c.SelectSingleNode("m:v", ns)
                    if v is None:
                        continue
                    text = v.InnerText
                    if cell_type == "s":
                        text = self._shared[int(text)]
                    elif cell_type == "b":
                        text = "TRUE" if text == "1" else "FALSE"
                    elif cell_type in ("", "n"):
                        text = _number_text(text)
                if text != "":
                    cells[(row_number, col)] = text
        self._cache[sheet_name] = cells
        return cells

    def read_range(self, sheet_name, ref):
        """Values of a range, row by row, '' for empty cells."""
        (r1, c1), (r2, c2) = parse_range(ref)
        cells = self._cells(sheet_name)
        return [cells.get((r, c), "").strip()
                for r in range(r1, r2 + 1)
                for c in range(c1, c2 + 1)]
