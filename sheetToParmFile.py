#!/usr/bin/env python3
# Convert the first sheet of an .ods or .xlsx spreadsheet to the white space
# separated parameter file format read by parmScanner, written to stdout.
#
# Uses only the Python standard library: both formats are zip archives of XML.
#
# Number cells, including formula results, are written as their stored value in
# its shortest exact form (1e-05, 0.1, 20), not as the cell displays it, since the
# display may be rounded. Text cells are written exactly as they are.
#
# Copyright (C) 2025-2026 Johan Roenby, STROMNING
#
# This program is free software: you can redistribute it and/or modify it under
# the terms of the GNU General Public License as published by the Free Software
# Foundation, either version 3 of the License, or (at your option) any later
# version. See the file LICENSE, or <https://www.gnu.org/licenses/>.

import re
import sys
import zipfile
import xml.etree.ElementTree as ET

# parmScanner reads lines 1-4 by word splitting and the case lines after them
# with xargs, so the two parts must be written differently.
nHeaderLines = 4


def fail(message):
    sys.exit("sheetToParmFile: " + message)


def number(value):
    x = float(value)
    if x.is_integer() and abs(x) < 1e15:
        return str(int(x))
    return repr(x)


# ------------------------------------------------------------------------------
# OpenDocument (.ods)

ODS_TABLE = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
ODS_OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
ODS_TEXT = "{urn:oasis:names:tc:opendocument:xmlns:text:1.0}"


def odsText(element):
    # Runs of spaces are stored as <text:s text:c="n"/> and tabs as <text:tab/>
    parts = [element.text or ""]
    for child in element:
        if child.tag == ODS_TEXT + "s":
            parts.append(" " * int(child.get(ODS_TEXT + "c", "1")))
        elif child.tag == ODS_TEXT + "tab":
            parts.append("\t")
        elif child.tag == ODS_TEXT + "line-break":
            parts.append("\n")
        else:
            parts.append(odsText(child))
        parts.append(child.tail or "")
    return "".join(parts)


def odsCell(cell):
    valueType = cell.get(ODS_OFFICE + "value-type")
    if valueType is None:
        return ""
    if valueType in ("float", "percentage", "currency"):
        return number(cell.get(ODS_OFFICE + "value"))
    if valueType == "boolean":
        return cell.get(ODS_OFFICE + "boolean-value")
    return "\n".join(odsText(p) for p in cell.iter(ODS_TEXT + "p"))


def readOds(archive):
    root = ET.fromstring(archive.read("content.xml"))
    table = next(root.iter(ODS_TABLE + "table"), None)
    if table is None:
        fail("no sheet found")
    rows = []
    for row in table.iter(ODS_TABLE + "table-row"):
        cells = []
        for cell in row:
            if cell.tag not in (ODS_TABLE + "table-cell",
                                ODS_TABLE + "covered-table-cell"):
                continue
            value = odsCell(cell)
            repeat = int(cell.get(ODS_TABLE + "number-columns-repeated", "1"))
            # A sheet ends in one empty cell repeated to the last column, and an
            # empty cell anywhere else is an error, so one copy is enough.
            cells.extend([value] * (repeat if value else 1))
        repeat = int(row.get(ODS_TABLE + "number-rows-repeated", "1"))
        # Likewise the empty rows repeated to the last row of the sheet
        rows.extend([cells] * (repeat if any(cells) else min(repeat, 1000)))
    return rows


# ------------------------------------------------------------------------------
# Office Open XML (.xlsx)

XLSX = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
XLSX_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


def xlsxText(element):
    # Rich text is split into runs, and <rPh> holds phonetic hints, not content
    return "".join(t.text or "" for child in element
                   for t in ([child] if child.tag == XLSX + "t" else
                             child.iter(XLSX + "t") if child.tag == XLSX + "r" else []))


def columnIndex(reference):
    index = 0
    for letter in re.match(r"[A-Z]+", reference).group():
        index = 26 * index + ord(letter) - ord("A") + 1
    return index - 1


def readXlsx(archive):
    names = archive.namelist()
    shared = []
    if "xl/sharedStrings.xml" in names:
        shared = [xlsxText(si) for si in
                  ET.fromstring(archive.read("xl/sharedStrings.xml")).iter(XLSX + "si")]

    # The first sheet is the first one listed in the workbook, not necessarily sheet1.xml
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    sheet = next(workbook.iter(XLSX + "sheet"), None)
    if sheet is None:
        fail("no sheet found")
    relations = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    target = next(r.get("Target") for r in relations.iter(PKG_REL + "Relationship")
                  if r.get("Id") == sheet.get(XLSX_REL + "id"))
    target = target[1:] if target.startswith("/") else "xl/" + target

    rows = {}
    rowIndex = -1
    for row in ET.fromstring(archive.read(target)).iter(XLSX + "row"):
        # Row and cell references are optional; without them the position is the next one
        rowIndex = int(row.get("r")) - 1 if row.get("r") else rowIndex + 1
        cells = {}
        cellIndex = -1
        for cell in row.iter(XLSX + "c"):
            cellIndex = columnIndex(cell.get("r")) if cell.get("r") else cellIndex + 1
            kind = cell.get("t", "n")
            v = cell.find(XLSX + "v")
            v = None if v is None else v.text
            if kind == "inlineStr":
                value = xlsxText(cell.find(XLSX + "is"))
            elif v is None:
                value = ""
            elif kind == "s":
                value = shared[int(v)]
            elif kind == "n":
                value = number(v)
            elif kind == "b":
                value = "true" if v == "1" else "false"
            elif kind == "e":
                fail("cell %s holds the error %s" % (cell.get("r", "?"), v))
            else:  # "str" (text result of a formula) and "d" (ISO date)
                value = v
            cells[cellIndex] = value
        rows[rowIndex] = [cells.get(i, "") for i in range(max(cells, default=-1) + 1)]
    return [rows.get(i, []) for i in range(max(rows, default=-1) + 1)]


# ------------------------------------------------------------------------------

def cellName(row, column):
    letters = ""
    column += 1
    while column:
        column, remainder = divmod(column - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return "%s%d" % (letters, row + 1)


def main():
    if len(sys.argv) != 2:
        sys.exit("Usage: sheetToParmFile.py <file.ods|file.xlsx>")
    fileName = sys.argv[1]
    try:
        archive = zipfile.ZipFile(fileName)
    except (OSError, zipfile.BadZipFile) as error:
        fail("cannot read %s as a spreadsheet: %s" % (fileName, error))
    with archive:
        if "content.xml" in archive.namelist():
            rows = readOds(archive)
        elif "xl/workbook.xml" in archive.namelist():
            rows = readXlsx(archive)
        else:
            fail("%s is neither an .ods nor an .xlsx file" % fileName)

    lines = []
    for i, cells in enumerate(rows):
        while cells and cells[-1] == "":
            cells = cells[:-1]
        fields = []
        for j, value in enumerate(cells):
            where = "cell %s" % cellName(i, j)
            if value == "":
                fail("%s is empty, but cells to the right of it are not" % where)
            if "\n" in value:
                fail("%s contains a line break" % where)
            if i < nHeaderLines:
                if re.search(r"\s", value):
                    fail("%s contains white space, which is not allowed in lines 1-4" % where)
            elif '"' in value:
                fail('%s contains a double quote, which parmScanner cannot pass on' % where)
            elif re.search(r"[\s'\\]", value):
                value = '"%s"' % value
            fields.append(value)
        lines.append("  ".join(fields))
    while lines and lines[-1] == "":
        lines.pop()
    print("\n".join(lines))


if __name__ == "__main__":
    main()
