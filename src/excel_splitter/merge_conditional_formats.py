"""Conservatively restore a shared whole-column rule after Excel's block pastes."""

from __future__ import annotations

import logging
import os
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path
from xml.dom import minidom
from xml.parsers.expat import ExpatError
from zipfile import BadZipFile, ZipFile


_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_LOG = logging.getLogger(__name__)
_LAYOUT_NAMES = ("print_area", "print_titles", "_filterdatabase")


@dataclass
class _Package:
    sheet_part: str
    sheet: minidom.Document
    styles: minidom.Document
    context: tuple


def _elements(node, name):
    return list(node.getElementsByTagNameNS(_NS, name))


def _standard(node):
    # Import only self-contained SpreadsheetML; extension prefixes can depend
    # on declarations on the source document's root.
    return all(
        element.namespaceURI == _NS and element.prefix is None
        and all(attribute.namespaceURI is None for attribute in element.attributes.values())
        for element in [node, *node.getElementsByTagName("*")]
    )


def _read(path):
    with ZipFile(path) as archive:
        members = archive.namelist()
        sheets = [name for name in members if re.fullmatch(r"xl/worksheets/[^/]+\.xml", name)]
        tables = [name for name in members if re.fullmatch(r"xl/tables/[^/]+\.xml", name)]
        if len(sheets) != 1 or len(tables) != 1 or any(name.startswith("xl/externalLinks/") for name in members):
            raise ValueError("unsupported package layout or external links")
        workbook = minidom.parseString(archive.read("xl/workbook.xml"))
        if any(node.getAttribute("name").casefold() not in {"_xlnm." + name for name in _LAYOUT_NAMES}
               for node in _elements(workbook, "definedName")):
            raise ValueError("defined-name context")
        table = minidom.parseString(archive.read(tables[0])).documentElement
        bounds = re.fullmatch(r"([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)", table.getAttribute("ref"))
        if not bounds:
            raise ValueError("unsupported table bounds")
        styles = minidom.parseString(archive.read("xl/styles.xml"))
        colors = tuple(node.toxml() for node in _elements(styles, "colors"))
        themes = tuple((name, archive.read(name)) for name in sorted(members) if name.startswith("xl/theme/"))
        context = (bounds[1], bounds[2], bounds[3], themes, colors)
        return _Package(sheets[0], minidom.parseString(archive.read(sheets[0])), styles, context)


def _formats(package):
    nodes = list(package.sheet.getElementsByTagNameNS("*", "conditionalFormatting"))
    if any(node.namespaceURI != _NS or node.parentNode != package.sheet.documentElement or not _standard(node) for node in nodes):
        raise ValueError("extended conditional formatting")
    return nodes


def _dxf(package, rule):
    if not rule.hasAttribute("dxfId"):
        return None
    index = int(rule.getAttribute("dxfId"))
    dxfs = _elements(package.styles, "dxf")
    if index < 0 or index >= len(dxfs) or not _standard(dxfs[index]):
        raise ValueError("unsupported conditional style")
    return dxfs[index].toxml()


def _column_index(name):
    result = 0
    for letter in name:
        result = result * 26 + ord(letter) - ord("A") + 1
    return result


def _original(package):
    formats = _formats(package)
    if len(formats) != 1 or len(_elements(formats[0], "cfRule")) != 1:
        raise ValueError("not exactly one conditional-format rule")
    node = formats[0]
    scope = re.fullmatch(r"(?P<column>[A-Z]{1,3})1:(?P=column)1048576", node.getAttribute("sqref"))
    if not scope:
        raise ValueError("not one full-column range")
    if not _column_index(package.context[0]) <= _column_index(scope[1]) <= _column_index(package.context[2]):
        raise ValueError("formatted column lies outside the merged table")
    rule = _elements(node, "cfRule")[0]
    formulas = ["".join(child.data for child in formula.childNodes if child.nodeType == child.TEXT_NODE)
                for formula in _elements(rule, "formula")]
    formulas.extend(node.getAttribute("val") for node in _elements(rule, "cfvo") if node.getAttribute("type") == "formula")
    for text in formulas:
        if "!" in text or "[" in text or any(name in text.casefold() for name in _LAYOUT_NAMES):
            raise ValueError("sheet, external, or structured formula reference")
    return node, (node.toxml(), _dxf(package, rule), package.context)


def restore_full_column_rule(sources: tuple[Path, ...], target: Path) -> bool:
    """Restore one identical global rule, or leave the Excel-produced file intact.

    This restores original global semantics: aggregate/absolute references are
    evaluated against the combined data, not separately against each input.
    """
    # ponytail: one standard rule only; broaden when a real mixed-rule case
    # has a proven priority/style/context mapping and a native Excel fixture.
    try:
        first = _read(sources[0])
        original, signature = _original(first)
        for source in sources[1:]:
            if _original(_read(source))[1] != signature:
                raise ValueError("source rules, styles, or layout differ")
        output = _read(target)
        formats = _formats(output)
        rule = _elements(original, "cfRule")[0]
        if not formats or output.context != first.context or _dxf(output, rule) != signature[1]:
            raise ValueError("saved output changed style identifiers or context")
    except (OSError, BadZipFile, KeyError, ValueError, ExpatError) as exc:
        _LOG.info("Full-column conditional-format restoration skipped: %s", exc)
        return False

    root = output.sheet.documentElement
    root.replaceChild(output.sheet.importNode(original, True), formats[0])
    for node in formats[1:]:
        root.removeChild(node)
    replacement = output.sheet.toxml(encoding="utf-8")
    descriptor, filename = tempfile.mkstemp(prefix=".cf-", suffix=".xlsx", dir=target.parent)
    os.close(descriptor)
    temporary = Path(filename)
    try:
        with ZipFile(target) as source, ZipFile(temporary, "w") as destination:
            destination.comment = source.comment
            for member in source.infolist():
                destination.writestr(member, replacement if member.filename == output.sheet_part else source.read(member))
        os.replace(temporary, target)
    finally:
        temporary.unlink(missing_ok=True)
    return True
