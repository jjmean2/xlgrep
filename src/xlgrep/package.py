"""The xlsx package structure: workbook part, relationships, sheets.

Elements are matched by local name so prefixed namespaces and Strict OOXML work.
"""

from __future__ import annotations

import posixpath
import zipfile
from collections.abc import Iterator
from dataclasses import dataclass
from xml.etree import ElementTree as ET

REL_SHEET = "/worksheet"
REL_COMMENTS = "/comments"
REL_OFFICE_DOC = "/officeDocument"
REL_SHARED_STRINGS = "/sharedStrings"
REL_STYLES = "/styles"
REL_TABLE = "/table"
REL_PIVOT_TABLE = "/pivotTable"
REL_DRAWING = "/drawing"
REL_CHART = "/chart"
REL_CONNECTIONS = "/connections"


def local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def children(elem: ET.Element, name: str) -> Iterator[ET.Element]:
    return (c for c in elem if local(c.tag) == name)


def rels(zf: zipfile.ZipFile, part: str, external: bool = False) -> dict[str, tuple[str, str]]:
    """Relationship id -> (type, resolved zip path) for ``part``.

    With ``external``, only targets outside the package, as written (file paths, URLs).
    """
    base, name = posixpath.split(part)
    rels_path = posixpath.join(base, "_rels", name + ".rels")
    if rels_path not in zf.namelist():
        return {}
    out = {}
    for rel in ET.fromstring(zf.read(rels_path)):
        target = rel.get("Target", "")
        if (rel.get("TargetMode") == "External") != external:
            continue
        if external:
            out[rel.get("Id", "")] = (rel.get("Type", ""), target)
            continue
        resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base, target))
        out[rel.get("Id", "")] = (rel.get("Type", ""), resolved)
    return out


@dataclass
class SheetInfo:
    name: str
    hidden: bool
    part: str | None  # zip path of the worksheet XML; None for chartsheets etc.


@dataclass
class Package:
    workbook_part: str
    workbook: ET.Element  # parsed workbook.xml
    workbook_rels: dict[str, tuple[str, str]]
    sheets: list[SheetInfo]

    def part_of_type(self, rel_suffix: str) -> str | None:
        return next((p for t, p in self.workbook_rels.values() if t.endswith(rel_suffix)), None)

    @property
    def date1904(self) -> bool:
        pr = next(children(self.workbook, "workbookPr"), None)
        return pr is not None and pr.get("date1904", "").lower() in ("1", "true")


def open_package(zf: zipfile.ZipFile) -> Package:
    root_rels = rels(zf, "")
    wb_part = next((p for t, p in root_rels.values() if t.endswith(REL_OFFICE_DOC)), "xl/workbook.xml")
    wb_rels = rels(zf, wb_part)
    wb_root = ET.fromstring(zf.read(wb_part))
    sheets: list[SheetInfo] = []
    for sheets_elem in children(wb_root, "sheets"):
        for s in children(sheets_elem, "sheet"):
            rid = next((v for k, v in s.attrib.items() if local(k) == "id"), "")
            rel_type, part = wb_rels.get(rid, ("", ""))
            sheets.append(
                SheetInfo(
                    name=s.get("name", ""),
                    hidden=s.get("state", "visible") != "visible",
                    part=part if rel_type.endswith(REL_SHEET) else None,
                )
            )
    return Package(wb_part, wb_root, wb_rels, sheets)


def sheet_parts(zf: zipfile.ZipFile, info: SheetInfo) -> tuple[int, int, int]:
    """(tables, pivot tables, charts) attached to a worksheet."""
    if info.part is None:
        return 0, 0, 0
    tables = pivots = charts = 0
    for rel_type, part in rels(zf, info.part).values():
        if rel_type.endswith(REL_TABLE):
            tables += 1
        elif rel_type.endswith(REL_PIVOT_TABLE):
            pivots += 1
        elif rel_type.endswith(REL_DRAWING):
            charts += sum(1 for t, _ in rels(zf, part).values() if t.endswith(REL_CHART))
    return tables, pivots, charts


@dataclass
class ExternalBook:
    """A workbook that formulas refer to as [n]: n-1 is its index in external_books()."""

    target: str  # path as stored: "file:///C:\\x\\Rates.xlsx", "../Budget.xlsx", "\\\\server\\x.xlsx"
    sheets: list[str]


def external_books(zf: zipfile.ZipFile, pkg: Package) -> list[ExternalBook | None]:
    """External workbooks in [n] order; None for DDE/OLE links (not workbooks)."""
    books: list[ExternalBook | None] = []
    for refs in children(pkg.workbook, "externalReferences"):
        for ref in children(refs, "externalReference"):
            rid = next((v for k, v in ref.attrib.items() if local(k) == "id"), "")
            _, part = pkg.workbook_rels.get(rid, ("", ""))
            book = None
            if part in zf.namelist():
                root = ET.fromstring(zf.read(part))
                elem = next(children(root, "externalBook"), None)
                if elem is not None:
                    target_id = next((v for k, v in elem.attrib.items() if local(k) == "id"), "")
                    target = rels(zf, part, external=True).get(target_id, ("", ""))[1]
                    sheets = [
                        s.get("val", "") for names in children(elem, "sheetNames") for s in children(names, "sheetName")
                    ]
                    book = ExternalBook(target, sheets)
            books.append(book)
    return books


@dataclass
class Connection:
    name: str
    kind: str  # "OLE DB", "ODBC", "web query", "text file", "Power Query", ...


# The "type" attribute of <connection>.
_CONNECTION_TYPES = {
    "1": "ODBC",
    "2": "DAO",
    "3": "file",
    "4": "web query",
    "5": "OLE DB",
    "6": "text file",
    "7": "ADO",
    "8": "DSP",
}


def connections(zf: zipfile.ZipFile, pkg: Package) -> list[Connection]:
    """Data connections (databases, web queries, Power Query). Connection strings are not
    read out: they can hold server names and passwords."""
    part = pkg.part_of_type(REL_CONNECTIONS)
    if not part or part not in zf.namelist():
        return []
    found = []
    for c in ET.fromstring(zf.read(part)):
        if local(c.tag) != "connection":
            continue
        kind = _CONNECTION_TYPES.get(c.get("type", ""), "other")
        db = next(children(c, "dbPr"), None)
        if db is not None and "Microsoft.Mashup" in db.get("connection", ""):
            kind = "Power Query"
        found.append(Connection(c.get("name", ""), kind))
    return found
