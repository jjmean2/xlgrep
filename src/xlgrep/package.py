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


def local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def children(elem: ET.Element, name: str) -> Iterator[ET.Element]:
    return (c for c in elem if local(c.tag) == name)


def rels(zf: zipfile.ZipFile, part: str) -> dict[str, tuple[str, str]]:
    """Relationship id -> (type, resolved zip path) for ``part`` (internal targets only)."""
    base, name = posixpath.split(part)
    rels_path = posixpath.join(base, "_rels", name + ".rels")
    if rels_path not in zf.namelist():
        return {}
    out = {}
    for rel in ET.fromstring(zf.read(rels_path)):
        target = rel.get("Target", "")
        if rel.get("TargetMode") == "External":
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
            sheets.append(SheetInfo(
                name=s.get("name", ""),
                hidden=s.get("state", "visible") != "visible",
                part=part if rel_type.endswith(REL_SHEET) else None,
            ))
    return Package(wb_part, wb_root, wb_rels, sheets)
