"""Formulas and text that live outside cells: defined names, conditional formats,
data validations and cell notes.

These are read straight from the package XML. openpyxl's read-only mode skips
conditional formats, validations and comments, its normal mode is slow on large
sheets, and neither reads the x14 extension blocks where Excel 2010+ stores rules
that reference other sheets.
"""

from __future__ import annotations

import posixpath
import zipfile
from collections.abc import Iterator
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

from .address import CellRange, parse_range, quote_sheet
from .text import normalize_formula

OBJECT_KINDS = ("name", "cf", "dv", "note")

_REL_SHEET = "/worksheet"
_REL_COMMENTS = "/comments"
_REL_OFFICE_DOC = "/officeDocument"


@dataclass
class SheetObject:
    object: str  # one of OBJECT_KINDS
    sheet: str | None  # None for workbook-scoped names
    ref: str  # defined name, cell (note) or comma-separated ranges (cf/dv)
    text: str  # formula with leading "=", or note text
    is_formula: bool
    detail: dict[str, str] = field(default_factory=dict)

    @property
    def location(self) -> str:
        """``TaxRate``, ``Summary!Local``, ``Summary!B2:B50`` …"""
        return self.ref if self.sheet is None else f"{quote_sheet(self.sheet)}!{self.ref}"

    def in_range(self, cell_range: CellRange) -> bool:
        if self.object == "name":
            return False
        return any(parse_range(part).intersects(cell_range) for part in self.ref.split(","))

    @property
    def areas(self) -> list[CellRange]:
        """The ranges a cf/dv rule applies to (or the note's cell)."""
        return [parse_range(part) for part in self.ref.split(",")] if self.object != "name" else []


@dataclass
class SheetInfo:
    name: str
    hidden: bool
    part: str | None  # zip path of the worksheet XML; None for chartsheets


@dataclass
class WorkbookObjects:
    sheets: list[SheetInfo]
    by_sheet: dict[str, list[SheetObject]]
    workbook_names: list[SheetObject]


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _children(elem: ET.Element, name: str) -> Iterator[ET.Element]:
    return (c for c in elem if _local(c.tag) == name)


def _rels(zf: zipfile.ZipFile, part: str) -> dict[str, tuple[str, str]]:
    """Relationship id -> (type, resolved zip path) for ``part``."""
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


def _formula(text: str | None, raw_formula: bool) -> str | None:
    if text is None or not text.strip():
        return None
    text = text if text.startswith("=") else "=" + text
    return text if raw_formula else normalize_formula(text)


def _sqref(text: str | None) -> str:
    return ",".join((text or "").split())


def read_objects(path: Path, kinds: set[str], raw_formula: bool = False) -> WorkbookObjects:
    with zipfile.ZipFile(path) as zf:
        root_rels = _rels(zf, "")
        wb_part = next((p for t, p in root_rels.values() if t.endswith(_REL_OFFICE_DOC)), "xl/workbook.xml")
        wb_rels = _rels(zf, wb_part)
        wb_root = ET.fromstring(zf.read(wb_part))

        sheets: list[SheetInfo] = []
        for sheets_elem in _children(wb_root, "sheets"):
            for s in _children(sheets_elem, "sheet"):
                rid = next((v for k, v in s.attrib.items() if _local(k) == "id"), "")
                rel_type, part = wb_rels.get(rid, ("", ""))
                sheets.append(SheetInfo(
                    name=s.get("name", ""),
                    hidden=s.get("state", "visible") != "visible",
                    part=part if rel_type.endswith(_REL_SHEET) else None,
                ))

        by_sheet: dict[str, list[SheetObject]] = {s.name: [] for s in sheets}
        workbook_names: list[SheetObject] = []

        if "name" in kinds:
            for names_elem in _children(wb_root, "definedNames"):
                for dn in _children(names_elem, "definedName"):
                    name = dn.get("name", "")
                    # Built-ins (_xlnm.Print_Area, _xlnm._FilterDatabase) and the hidden
                    # placeholders Excel adds for newer functions (_xlfn.XLOOKUP) are noise.
                    if name.lower().startswith(("_xlnm.", "_xlfn.")):
                        continue
                    text = _formula(dn.text, raw_formula)
                    if text is None:
                        continue
                    local_id = dn.get("localSheetId")
                    sheet = sheets[int(local_id)].name if local_id is not None and int(local_id) < len(sheets) else None
                    obj = SheetObject("name", sheet, name, text, True,
                                      {"hidden": "true"} if dn.get("hidden") in ("1", "true") else {})
                    (by_sheet[sheet] if sheet is not None else workbook_names).append(obj)

        for info in sheets:
            if info.part is None:
                continue
            found = by_sheet[info.name]
            if kinds & {"cf", "dv"}:
                found[:0] = _sheet_rules(zf, info, kinds, raw_formula)
            if "note" in kinds:
                found.extend(_sheet_notes(zf, info))

    return WorkbookObjects(sheets, by_sheet, workbook_names)


def _sheet_tail(zf: zipfile.ZipFile, part: str) -> ET.Element:
    """Parse the worksheet XML without its (potentially huge) <sheetData>.

    Conditional formats, validations and extensions all come after sheetData,
    so the root start tag (for namespace declarations) plus the tail is enough.
    """
    data = zf.read(part)
    start = data.find(b"<sheetData")
    if start == -1:
        return ET.fromstring(data)
    end_tag = data.find(b"</sheetData>", start)
    if end_tag != -1:
        end = end_tag + len(b"</sheetData>")
    else:  # self-closing <sheetData/>
        end = data.find(b">", start) + 1
    return ET.fromstring(data[:start] + data[end:])


def _sheet_rules(zf: zipfile.ZipFile, info: SheetInfo, kinds: set[str], raw_formula: bool) -> list[SheetObject]:
    root = _sheet_tail(zf, info.part)
    out: list[SheetObject] = []

    if "cf" in kinds:
        for cf in _children(root, "conditionalFormatting"):
            ref = _sqref(cf.get("sqref"))
            for rule in _children(cf, "cfRule"):
                texts = [f.text for f in _children(rule, "formula")]
                for scale in rule:  # colorScale / dataBar / iconSet thresholds
                    texts += [v.get("val") for v in _children(scale, "cfvo") if v.get("type") == "formula"]
                out += _rule_objects("cf", info.name, ref, texts, raw_formula, {"rule_type": rule.get("type", "")})

    if "dv" in kinds:
        for dvs in _children(root, "dataValidations"):
            for dv in _children(dvs, "dataValidation"):
                ref = _sqref(dv.get("sqref"))
                for part in ("formula1", "formula2"):
                    texts = [f.text for f in _children(dv, part)]
                    out += _rule_objects("dv", info.name, ref, texts, raw_formula,
                                         {"dv_type": dv.get("type", "any"), "part": part})

    # Excel 2010+ extensions (rules referring to other sheets, newer rule types).
    for ext_list in _children(root, "extLst"):
        for ext in _children(ext_list, "ext"):
            for block in ext:
                name = _local(block.tag)
                if name == "conditionalFormattings" and "cf" in kinds:
                    for cf in _children(block, "conditionalFormatting"):
                        ref = _sqref(next((s.text for s in _children(cf, "sqref")), ""))
                        for rule in _children(cf, "cfRule"):
                            texts = [f.text for f in rule.iter() if _local(f.tag) == "f"]
                            out += _rule_objects("cf", info.name, ref, texts, raw_formula,
                                                 {"rule_type": rule.get("type", "")})
                elif name == "dataValidations" and "dv" in kinds:
                    for dv in _children(block, "dataValidation"):
                        ref = _sqref(next((s.text for s in _children(dv, "sqref")), ""))
                        for part in ("formula1", "formula2"):
                            texts = [f.text for p in _children(dv, part) for f in _children(p, "f")]
                            out += _rule_objects("dv", info.name, ref, texts, raw_formula,
                                                 {"dv_type": dv.get("type", "any"), "part": part})
    return out


def _rule_objects(kind, sheet, ref, texts, raw_formula, detail) -> list[SheetObject]:
    out = []
    for text in texts:
        formula = _formula(text, raw_formula)
        if formula is not None:
            out.append(SheetObject(kind, sheet, ref, formula, True, dict(detail)))
    return out


def _sheet_notes(zf: zipfile.ZipFile, info: SheetInfo) -> list[SheetObject]:
    out = []
    for rel_type, part in _rels(zf, info.part).values():
        if not rel_type.endswith(_REL_COMMENTS) or part not in zf.namelist():
            continue
        root = ET.fromstring(zf.read(part))
        for comment_list in _children(root, "commentList"):
            for comment in _children(comment_list, "comment"):
                text = "".join(_note_text(comment))
                if text:
                    out.append(SheetObject("note", info.name, comment.get("ref", ""), text, False))
    return out


def _note_text(comment: ET.Element) -> Iterator[str]:
    for text_elem in _children(comment, "text"):
        for child in text_elem:
            name = _local(child.tag)
            if name == "t":
                yield child.text or ""
            elif name == "r":  # rich-text run
                yield from (t.text or "" for t in _children(child, "t"))
            # "rPh" (phonetic guides) is skipped
