"""Reading workbooks into per-sheet cell grids.

Cells are read straight from the sheet XML instead of through openpyxl, which is
several times slower and can't return formulas and their cached results in one
pass. Sheets in the regular layout Excel writes are scanned with regular
expressions; anything unusual (namespace prefixes, cells without coordinates,
CDATA, comments, single-quoted attributes) goes through the standard XML parser.
Both paths produce identical cells (tests/test_reader.py).
"""

from __future__ import annotations

import html
import re
import zipfile
from collections.abc import Iterator
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

from openpyxl.styles.numbers import BUILTIN_FORMATS, is_date_format, is_timedelta_format
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils.datetime import CALENDAR_MAC_1904, WINDOWS_EPOCH, from_excel, from_ISO8601

from .address import col_index
from .package import REL_SHARED_STRINGS, REL_STYLES, Package, SheetInfo, children, local, open_package
from .refs import SharedFormula
from .text import normalize_formula, value_text


@dataclass(slots=True)
class Cell:
    row: int
    col: int
    formula: str | None = None  # formula text including the leading "=", if any
    value: object = None  # literal value, or the cached result of a formula
    has_cached_value: bool = False
    shared: str | None = None  # shared-formula group (si); cells of a group have the same logic
    array: bool = False  # an array (CSE / dynamic array) formula

    @property
    def is_formula(self) -> bool:
        return self.formula is not None

    @property
    def value_str(self) -> str:
        return value_text(self.value)

    @property
    def display(self) -> str:
        """What the cell looks like in the formula bar."""
        return self.formula if self.formula is not None else self.value_str


@dataclass
class Sheet:
    name: str
    hidden: bool
    cells: dict[tuple[int, int], Cell] = field(default_factory=dict)
    dimension: str | None = None  # the used range Excel recorded, e.g. "A1:J100"

    def get(self, row: int, col: int) -> Cell | None:
        return self.cells.get((row, col))

    def sorted_cells(self) -> list[Cell]:
        return [self.cells[k] for k in sorted(self.cells)]


class CellReader:
    """Reads the cells of one workbook's sheets.

    Shared strings and the date styles are loaded once, on creation. Formula cells
    carry both the formula text and the value Excel cached when it last saved.
    """

    def __init__(self, zf: zipfile.ZipFile, pkg: Package, raw_formula: bool = False):
        self.zf = zf
        self.ctx = _context(zf, pkg, raw_formula)

    def read(self, info: SheetInfo, formulas_only: bool = False) -> Sheet:
        """``formulas_only`` skips other cells, for searches that can only match
        formulas and show no neighbouring cells."""
        sheet = Sheet(info.name, info.hidden)
        if info.part is not None:
            data = self.zf.read(info.part)
            dimension = _DIMENSION_RE.search(data, 0, 65536)
            sheet.dimension = dimension.group(1).decode() if dimension else None
            sheet.cells = read_cells(data, self.ctx, formulas_only=formulas_only)
        return sheet


def read_sheets(path: Path, raw_formula: bool = False) -> Iterator[Sheet]:
    """All worksheets of ``path`` (convenience for scripts and tests)."""
    with zipfile.ZipFile(path) as zf:
        pkg = open_package(zf)
        reader = CellReader(zf, pkg, raw_formula)
        for info in pkg.sheets:
            if info.part is not None:
                yield reader.read(info)


# ---------------------------------------------------------------- context


@dataclass
class _Context:
    strings: list[str]
    date_styles: set[int]
    timedelta_styles: set[int]
    epoch: object
    raw_formula: bool


def _context(zf: zipfile.ZipFile, pkg: Package, raw_formula: bool) -> _Context:
    strings: list[str] = []
    part = pkg.part_of_type(REL_SHARED_STRINGS)
    if part and part in zf.namelist():
        with zf.open(part) as f:
            for _, el in ET.iterparse(f):
                if local(el.tag) == "si":
                    strings.append(_rich_text(el))
                    el.clear()
    date_styles, timedelta_styles = _date_styles(zf, pkg.part_of_type(REL_STYLES))
    epoch = CALENDAR_MAC_1904 if pkg.date1904 else WINDOWS_EPOCH
    return _Context(strings, date_styles, timedelta_styles, epoch, raw_formula)


# East Asian locales show these built-in ids as dates; openpyxl doesn't list them.
_CJK_DATE_IDS = set(range(27, 37)) | set(range(50, 59))


def _date_styles(zf: zipfile.ZipFile, part: str | None) -> tuple[set[int], set[int]]:
    """Indexes of cell styles (the ``s`` attribute) whose number format is a date/time."""
    if not part or part not in zf.namelist():
        return set(), set()
    root = ET.fromstring(zf.read(part))
    custom = {}
    for fmts in children(root, "numFmts"):
        for fmt in children(fmts, "numFmt"):
            custom[int(fmt.get("numFmtId", "0"))] = fmt.get("formatCode", "")
    dates, timedeltas = set(), set()
    for xfs in children(root, "cellXfs"):
        for idx, xf in enumerate(children(xfs, "xf")):
            fmt_id = int(xf.get("numFmtId", "0"))
            code = custom.get(fmt_id, BUILTIN_FORMATS.get(fmt_id))
            if (fmt_id in _CJK_DATE_IDS and fmt_id not in custom) or (code and is_date_format(code)):
                dates.add(idx)
                if code and is_timedelta_format(code):
                    timedeltas.add(idx)
    return dates, timedeltas


# ---------------------------------------------------------------- values


_OOXML_ESCAPE_RE = re.compile(r"_x([0-9A-Fa-f]{4})_")


def _ooxml_unescape(text: str) -> str:
    """OOXML encodes control characters as _xHHHH_ (e.g. _x000D_ for CR)."""
    return _OOXML_ESCAPE_RE.sub(lambda m: chr(int(m.group(1), 16)), text) if "_x" in text else text


def _rich_text(si: ET.Element) -> str:
    """Text of an <si>/<is> element: plain <t> and rich-text runs, without phonetic guides."""
    parts = []
    for child in si:
        name = local(child.tag)
        if name == "t":
            parts.append(child.text or "")
        elif name == "r":
            parts.extend(t.text or "" for t in children(child, "t"))
    return _ooxml_unescape("".join(parts))


def _number(text: str) -> int | float:
    return float(text) if ("." in text or "e" in text or "E" in text) else int(text)


def _convert(raw: str, cell_type: str, style: int, ctx: _Context) -> object:
    """A stored <v> value as a Python value."""
    if cell_type == "s":
        return ctx.strings[int(raw)]
    if cell_type == "b":
        return raw.strip() in ("1", "true")
    if cell_type in ("str", "e", "inlineStr"):
        return _ooxml_unescape(raw)
    if cell_type == "d":
        return from_ISO8601(raw)
    number = _number(raw)
    if style in ctx.date_styles:
        try:
            return from_excel(number, ctx.epoch, timedelta=style in ctx.timedelta_styles)
        except (OverflowError, ValueError):
            return "#VALUE!"
    return number


class _Formulas:
    """Builds formula text, expanding shared formulas from their master cell."""

    def __init__(self, raw_formula: bool):
        self.raw_formula = raw_formula
        self.shared: dict[str, SharedFormula] = {}

    def _normal(self, text: str) -> str:
        return text if self.raw_formula else normalize_formula(text)

    def text(self, attrs: dict[str, str], body: str | None, row: int, col: int) -> str | None:
        kind = attrs.get("t")
        if kind == "dataTable":
            args = ",".join(attrs[a] for a in ("r1", "r2") if attrs.get(a))
            return self._normal(f"=TABLE({args})")
        text = self._normal("=" + (body or ""))
        if kind == "shared":
            si = attrs.get("si", "")
            if si in self.shared:
                if not body:
                    return self.dependent(si, row, col)
            elif body:
                # Normalising doesn't touch references, so the template is stored normalised
                # and dependents skip it.
                self.shared[si] = SharedFormula(text, row, col)
        return text

    def dependent(self, si: str, row: int, col: int) -> str:
        """Formula of a cell that only points at shared formula ``si``."""
        shared = self.shared.get(si)
        return shared.at(row, col) if shared is not None else "="


def _make_cell(
    row: int,
    col: int,
    cell_type: str,
    style: int,
    formula: tuple[dict[str, str], str | None] | None,
    raw_value: str | None,
    inline: str | None,
    ctx: _Context,
    formulas: _Formulas,
) -> Cell | None:
    value: object = None
    if cell_type == "inlineStr":
        value = inline
    elif raw_value is not None and raw_value != "":
        value = _convert(raw_value, cell_type, style, ctx)
    if formula is not None:
        attrs, body = formula
        text = formulas.text(attrs, body, row, col)
        kind = attrs.get("t")
        return Cell(
            row,
            col,
            formula=text,
            value=value,
            has_cached_value=value is not None,
            shared=attrs.get("si") if kind == "shared" else None,
            array=kind == "array",
        )
    if value is None:
        return None
    return Cell(row, col, value=value, has_cached_value=True)


# ---------------------------------------------------------------- regex path


class _Irregular(Exception):
    """The sheet isn't in the layout the regex scanner understands."""


_ROOT_PREFIX_RE = re.compile(rb"<\w+:worksheet\b")
_DIMENSION_RE = re.compile(rb'<(?:\w+:)?dimension\s+ref="([^"]+)"')
# One C-level pass pulls out coordinate, type, style and body; the lookaheads make
# attribute order irrelevant.
_CELL_RE = re.compile(
    rb'<c\b(?=[^>]*?\br="([A-Z]{1,3})(\d+)")(?:(?=[^>]*?\bt="(\w+)"))?(?:(?=[^>]*?\bs="(\d+)"))?'
    rb"[^>]*?(?:/>|>(.*?)</c>)",
    re.DOTALL,
)
_NO_COORD_RE = re.compile(rb'<c\b(?![^>]*?\br=")')
_SINGLE_QUOTED_RE = re.compile(rb"<[^>]*='")
_ATTR_RE = re.compile(rb'([\w:]+)="([^"]*)"')
_F_RE = re.compile(rb"<f\b([^>]*?)(?:/>|>(.*?)</f>)", re.DOTALL)
_SHARED_DEPENDENT_RE = re.compile(rb'<f t="shared" si="(\d+)"\s*/>(?:<v>([^<]*)</v>)?')
_V_RE = re.compile(rb"<v(?:\s[^>]*)?>(.*?)</v>", re.DOTALL)
_IS_RE = re.compile(rb"<is>(.*?)</is>", re.DOTALL)
_RPH_RE = re.compile(rb"<rPh\b.*?</rPh>", re.DOTALL)
_T_RE = re.compile(rb"<t\b[^>]*>(.*?)</t>", re.DOTALL)


def _xml_text(raw: bytes) -> str:
    text = raw.decode("utf-8")
    return html.unescape(text) if "&" in text else text


def _attrs(raw: bytes) -> dict[str, str]:
    return {k.decode(): _xml_text(v) for k, v in _ATTR_RE.findall(raw)} if raw.strip() else {}


def _sheet_data(data: bytes) -> bytes | None:
    """The <sheetData> part of a sheet, if it's in the layout the scanner handles."""
    # Check for a prefixed root first: then "<sheetData" isn't found at all.
    if _ROOT_PREFIX_RE.search(data, 0, 65536):
        raise _Irregular("namespace prefix")
    start = data.find(b"<sheetData")
    if start == -1:
        return None
    end = data.rfind(b"</sheetData>")
    region = data[start:end] if end != -1 else data[start:]
    if b"<![CDATA[" in region or b"<!--" in region:
        raise _Irregular("CDATA or comment")
    if _NO_COORD_RE.search(region):
        raise _Irregular("cell without coordinate")
    if _SINGLE_QUOTED_RE.search(region):
        raise _Irregular("single-quoted attribute")
    return region


def _scan_regex(data: bytes, ctx: _Context, formulas_only: bool = False) -> dict[tuple[int, int], Cell]:
    region = _sheet_data(data)
    if region is None:
        return {}
    cells: dict[tuple[int, int], Cell] = {}
    formulas = _Formulas(ctx.raw_formula)
    columns: dict[bytes, int] = {}  # "AB" -> 28, computed once per column
    for letters, digits, raw_type, raw_style, body in _CELL_RE.findall(region):
        if not body or (formulas_only and b"<f" not in body):
            continue  # styled but empty, or not wanted
        col = columns.get(letters)
        if col is None:
            col = columns[letters] = col_index(letters.decode())
        row = int(digits)
        style = int(raw_style) if raw_style else 0
        cell = _fast_cell(row, col, raw_type, style, body, ctx, formulas)
        if cell is None:
            cell = _general_cell(row, col, raw_type, style, body, ctx, formulas)
        if cell is not None:
            cells[(row, col)] = cell
    return cells


def _fast_cell(
    row: int, col: int, raw_type: bytes, style: int, body: bytes, ctx: _Context, formulas: _Formulas
) -> Cell | None:
    """The shapes most cells in an Excel file take, without general parsing.

    Returns None when the cell is something else; _general_cell handles it then.
    Measured on a 1M-cell sheet, these paths make reading about twice as fast.
    """
    # <v>123</v> (a number, not a date) or <v>5</v> with t="s" (shared string #5)
    if body.startswith(b"<v>") and body.endswith(b"</v>") and b"<" not in body[3:-4]:
        raw = body[3:-4]
        if raw_type in (b"", b"n") and style not in ctx.date_styles:
            return Cell(row, col, value=_number(raw.decode()), has_cached_value=True)
        if raw_type == b"s":
            return Cell(row, col, value=ctx.strings[int(raw)], has_cached_value=True)
        return None
    # <f t="shared" si="0"/><v>…</v>: a cell filled from a shared formula
    if body.startswith(b'<f t="shared"'):
        dep = _SHARED_DEPENDENT_RE.fullmatch(body)
        if dep is None or dep.group(1).decode() not in formulas.shared:
            return None
        raw = dep.group(2)
        value = _convert(_xml_text(raw), raw_type.decode() or "n", style, ctx) if raw else None
        si = dep.group(1).decode()
        return Cell(
            row,
            col,
            formula=formulas.dependent(si, row, col),
            value=value,
            has_cached_value=value is not None,
            shared=si,
        )
    return None


def _general_cell(
    row: int, col: int, raw_type: bytes, style: int, body: bytes, ctx: _Context, formulas: _Formulas
) -> Cell | None:
    """Any cell: pull <f>, <v> and <is> out of the body and build it."""
    cell_type = raw_type.decode() or "n"
    formula = None
    f = _F_RE.search(body) if b"<f" in body else None
    if f is not None:
        formula = (_attrs(f.group(1)), _xml_text(f.group(2)) if f.group(2) is not None else None)
    v = _V_RE.search(body) if b"<v" in body else None
    inline = None
    if cell_type == "inlineStr":
        i = _IS_RE.search(body)
        if i is not None:
            content = _RPH_RE.sub(b"", i.group(1))
            inline = _ooxml_unescape("".join(_xml_text(t) for t in _T_RE.findall(content)))
    raw_value = _xml_text(v.group(1)) if v is not None else None
    return _make_cell(row, col, cell_type, style, formula, raw_value, inline, ctx, formulas)


# ---------------------------------------------------------------- XML parser path


def _scan_etree(data: bytes, ctx: _Context, formulas_only: bool = False) -> dict[tuple[int, int], Cell]:
    """Standard-parser fallback; also handles rows/cells with implicit positions."""
    cells: dict[tuple[int, int], Cell] = {}
    formulas = _Formulas(ctx.raw_formula)
    root = ET.fromstring(data)
    for sheet_data in children(root, "sheetData"):
        row_no = 0
        for row in children(sheet_data, "row"):
            row_no = int(row.get("r")) if row.get("r") else row_no + 1
            col_no = 0
            for c in children(row, "c"):
                coord = c.get("r")
                if coord:
                    _, col_no = coordinate_to_tuple(coord)
                else:
                    col_no += 1
                formula = raw_value = inline = None
                for child in c:
                    name = local(child.tag)
                    if name == "f":
                        formula = ({local(k): v for k, v in child.attrib.items()}, child.text)
                    elif name == "v":
                        raw_value = child.text
                    elif name == "is":
                        inline = _rich_text(child)
                cell = _make_cell(
                    row_no,
                    col_no,
                    c.get("t", "n"),
                    int(c.get("s", "0") or 0),
                    formula,
                    raw_value,
                    inline,
                    ctx,
                    formulas,
                )
                if cell is not None and (cell.formula is not None or not formulas_only):
                    cells[(row_no, col_no)] = cell
    return cells


def read_cells(
    data: bytes, ctx: _Context, *, force_etree: bool = False, formulas_only: bool = False
) -> dict[tuple[int, int], Cell]:
    if not force_etree:
        try:
            return _scan_regex(data, ctx, formulas_only)
        except _Irregular:
            pass
    return _scan_etree(data, ctx, formulas_only)
