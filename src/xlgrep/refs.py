"""--ref: find formulas that reference given ranges.

References are scanned with regular expressions over the formula (string literals
masked) rather than openpyxl's Tokenizer, which rejects spill references (A1#) and
drops the start of ranges like A1:INDEX(...). Regex matches also carry offsets for
highlighting.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from .address import CellRange, col_index, col_letter
from .text import mask_string_literals

_SHEET_QUOTED = r"'(?:[^']|'')+'"
_SHEET_PLAIN = r"[^\W\d][\w.]*"
_PREFIX = rf"(?P<prefix>(?:{_SHEET_QUOTED}|(?:\[[^\]]+\])?{_SHEET_PLAIN}(?::{_SHEET_PLAIN})?)!)?"
_COL = r"\$?[A-Za-z]{1,3}"
_ROW = r"\$?\d+"
_AREA = rf"(?P<area>{_COL}{_ROW}(?::{_COL}{_ROW})?|{_COL}:{_COL}|{_ROW}:{_ROW})"
_BEFORE = r"(?<![\w.$'\[\]!:])"
_REF_RE = re.compile(rf"{_BEFORE}{_PREFIX}{_AREA}(?P<spill>#)?(?![\w(\[!])")
_NAME_RE = re.compile(rf"{_BEFORE}{_PREFIX}(?P<name>[^\W\d][\w.]*)(?![\w.(\[!])")
_ENDPOINT_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?$")
_BRACKET_RE = re.compile(r"\[[^\[\]]*\]")

MAX_ROW, MAX_COL = 1_048_576, 16_384


@dataclass(frozen=True)
class Endpoint:
    row: int | None
    col: int | None
    row_abs: bool
    col_abs: bool


@dataclass
class Ref:
    """A reference found in a formula."""

    start: int
    end: int
    sheets: tuple[str, ...] | None  # None: unqualified; two names for a 3D range
    external: bool
    first: Endpoint | None = None  # set for cell/range references
    last: Endpoint | None = None
    name: str | None = None  # set for defined-name references


@dataclass(frozen=True)
class Target:
    sheet: str | None  # None: any sheet
    area: CellRange | None  # None: whole sheet

    def hits(self, sheet: str, area: CellRange) -> bool:
        if self.sheet is not None and self.sheet.lower() != sheet.lower():
            return False
        return self.area is None or self.area.intersects(area)


@dataclass
class RefHit:
    start: int
    end: int
    text: str
    via: str | None = None  # defined name the hit goes through


def _unquote(sheet: str) -> str:
    return sheet[1:-1].replace("''", "'") if sheet.startswith("'") else sheet


def _parse_prefix(prefix: str | None) -> tuple[tuple[str, ...] | None, bool]:
    if not prefix:
        return None, False
    body = _unquote(prefix[:-1])
    external = body.startswith("[")
    if external:
        body = body[body.index("]") + 1 :]
    return tuple(body.split(":", 1)), external


def _split_endpoint(text: str) -> tuple[str, str | None, str, str | None] | None:
    """('$', 'A', '', '5') for '$A5'. In a row-only endpoint like '$5' the '$' is the row's."""
    m = _ENDPOINT_RE.match(text)
    if not m:
        return None
    col_abs, col, row_abs, row = m.groups()
    if col is None and col_abs:
        col_abs, row_abs = "", col_abs
    return col_abs, col, row_abs, row


def _endpoint(text: str) -> Endpoint:
    parts = _split_endpoint(text)
    if parts is None:
        raise ValueError(text)
    col_abs, col, row_abs, row = parts
    return Endpoint(int(row) if row else None, col_index(col) if col else None, bool(row_abs), bool(col_abs))


def _area(text: str) -> tuple[Endpoint, Endpoint]:
    parts = text.split(":")
    first = _endpoint(parts[0])
    return first, (_endpoint(parts[1]) if len(parts) > 1 else first)


def parse_target(text: str) -> Target:
    """'Data!A:D', "'Raw Data'!B2", 'Data!' (whole sheet) or 'A1:B5' (any sheet)."""
    text = text.strip()
    sheet = None
    if "!" in text:
        prefix, _, rest = text.rpartition("!")
        if not prefix:
            raise ValueError(f"invalid reference: {text!r}")
        sheet, rest = _unquote(prefix), rest.strip()
        if not rest:
            return Target(sheet, None)
    else:
        rest = text
    try:
        first, last = _area(rest)
    except ValueError:
        raise ValueError(f"invalid reference: {text!r}") from None
    return Target(sheet, _bounds(first, last))


def _bounds(first: Endpoint, last: Endpoint, dr: int = 0, dc: int = 0) -> CellRange:
    """Bounding range of an area, swept by (dr, dc) for relative parts (cf/dv rules)."""

    def span(a, b, a_abs, b_abs, delta, limit):
        if a is None or b is None:
            return None, None
        # Offsets only grow, so the low end stays put and the high end is the
        # furthest a relative endpoint travels.
        hi = max(a + (0 if a_abs else delta), b + (0 if b_abs else delta))
        return min(a, b), min(hi, limit)

    r1, r2 = span(first.row, last.row, first.row_abs, last.row_abs, dr, MAX_ROW)
    c1, c2 = span(first.col, last.col, first.col_abs, last.col_abs, dc, MAX_COL)
    return CellRange(r1, c1, r2, c2)


def mask_formula(formula: str) -> str:
    """Blank out string literals and bracket contents, keeping offsets.

    Bracket contents are table columns (Table1[[#This Row],[Qty]]) or external
    workbook names ([1]Ext!A1); neither holds references.
    """
    return _BRACKET_RE.sub(lambda m: "[" + " " * (len(m.group(0)) - 2) + "]", mask_string_literals(formula))


def scan_refs(formula: str) -> list[Ref]:
    """Cell/range references and candidate defined names in ``formula``."""
    masked = mask_formula(formula)
    refs: list[Ref] = []
    taken: list[tuple[int, int]] = []
    for m in _REF_RE.finditer(masked):
        sheets, external = _parse_prefix(m.group("prefix"))
        try:
            first, last = _area(m.group("area"))
        except ValueError:
            continue
        refs.append(Ref(m.start(), m.end(), sheets, external, first, last))
        taken.append(m.span())
    for m in _NAME_RE.finditer(masked):
        if any(s < m.end() and m.start() < e for s, e in taken):
            continue
        sheets, external = _parse_prefix(m.group("prefix"))
        if m.group("name").upper() in ("TRUE", "FALSE"):
            continue
        refs.append(Ref(m.start(), m.end(), sheets, external, name=m.group("name")))
    refs.sort(key=lambda r: r.start)
    return refs


class RefFinder:
    """Matches references against targets within one workbook."""

    def __init__(self, targets: list[Target], sheet_order: list[str],
                 names: dict[tuple[str | None, str], str]):
        """``names`` maps (scope sheet or None, NAME upper) to the name's formula."""
        self.targets = targets
        self.sheet_order = sheet_order
        self._lower_order = [s.lower() for s in sheet_order]
        self.names = {(s.lower() if s else None, n): f for (s, n), f in names.items()}
        self._name_memo: dict[tuple[str | None, str], bool] = {}

    def _sheets(self, ref: Ref, context: str | None) -> list[str]:
        if ref.sheets is None:
            return [context] if context is not None else []
        if len(ref.sheets) == 1:
            return [ref.sheets[0]]
        lo, hi = (self._lower_order.index(s.lower()) if s.lower() in self._lower_order else -1 for s in ref.sheets)
        if lo < 0 or hi < 0:
            return []
        lo, hi = min(lo, hi), max(lo, hi)
        return self.sheet_order[lo : hi + 1]

    def _resolve_name(self, ref: Ref, context: str | None) -> tuple[str | None, str] | None:
        key = ref.name.upper()
        if ref.sheets is not None:
            scoped = (ref.sheets[0].lower(), key)
            return scoped if scoped in self.names else None
        if context is not None and (context.lower(), key) in self.names:
            return context.lower(), key
        return (None, key) if (None, key) in self.names else None

    def _name_hits(self, name_key: tuple[str | None, str], stack: frozenset = frozenset()) -> bool:
        if name_key in self._name_memo:
            return self._name_memo[name_key]
        if name_key in stack:
            return False
        scope = name_key[0]
        context = next((s for s in self.sheet_order if s.lower() == scope), None) if scope else None
        hit = bool(self._hits(self.names[name_key], context, stack=stack | {name_key}))
        self._name_memo[name_key] = hit
        return hit

    def _hits(self, formula: str, context: str | None, sweep: tuple[int, int] = (0, 0),
              stack: frozenset = frozenset()) -> list[RefHit]:
        hits = []
        for ref in scan_refs(formula):
            if ref.external:
                continue
            if ref.name is not None:
                key = self._resolve_name(ref, context)
                if key is not None and self._name_hits(key, stack):
                    hits.append(RefHit(ref.start, ref.end, formula[ref.start : ref.end], via=ref.name))
                continue
            area = _bounds(ref.first, ref.last, *sweep)
            if any(t.hits(sheet, area) for sheet in self._sheets(ref, context) for t in self.targets):
                hits.append(RefHit(ref.start, ref.end, formula[ref.start : ref.end]))
        return hits

    def find(self, formula: str, sheet: str | None, applies_to: list[CellRange] | None = None) -> list[RefHit]:
        """References in ``formula`` (on ``sheet``) that hit a target.

        ``applies_to`` is the sqref of a conditional format / validation: relative
        references move across it, anchored at the top-left of its first range.
        """
        sweep = (0, 0)
        if applies_to:
            anchor = applies_to[0]
            rows = [r.max_row for r in applies_to if r.max_row is not None]
            cols = [r.max_col for r in applies_to if r.max_col is not None]
            if anchor.min_row is not None and anchor.min_col is not None and rows and cols:
                sweep = (max(rows) - anchor.min_row, max(cols) - anchor.min_col)
        return self._hits(formula, sheet, sweep)


class SharedFormula:
    """A shared formula's master text, re-anchored for the other cells of its range.

    Excel stores a formula filled across a range once, in the top-left cell; the
    others hold only a reference to it. Relative parts of each reference move with
    the cell, like openpyxl's Translator does, but the formula is parsed once and
    each cell only does arithmetic and joins.
    """

    def __init__(self, formula: str, row: int, col: int):
        self.formula = formula
        self.row, self.col = row, col
        # Literal text, or a reference endpoint: (col_abs, col, row_abs, row, original).
        self.parts: list[str | tuple[str, int | None, str, int | None, str]] = []
        pos = 0
        for m in _REF_RE.finditer(mask_formula(formula)):
            self.parts.append(formula[pos : m.start("area")])
            sides = formula[m.start("area") : m.end("area")].split(":")
            for i, side in enumerate(sides):
                if i:
                    self.parts.append(":")
                split = _split_endpoint(side)
                if split is None:
                    self.parts.append(side)
                    continue
                col_abs, col, row_abs, row = split
                self.parts.append((col_abs, col_index(col) if col else None, row_abs, int(row) if row else None, side))
            pos = m.end("area")
        self.parts.append(formula[pos:])
        self._letters: dict[int, str] = {}

    def at(self, row: int, col: int) -> str:
        dr, dc = row - self.row, col - self.col
        if not dr and not dc:
            return self.formula
        out = []
        for part in self.parts:
            if part.__class__ is str:
                out.append(part)
                continue
            col_abs, c, row_abs, r, original = part
            move_col = c is not None and not col_abs and dc
            move_row = r is not None and not row_abs and dr
            if not move_col and not move_row:
                out.append(original)
                continue
            if c is None:
                col_text = ""
            elif move_col:
                letters = self._letters.get(c + dc)
                if letters is None:
                    letters = self._letters[c + dc] = col_letter(c + dc)
                col_text = letters
            else:
                col_text = col_letter(c)
            row_text = "" if r is None else str(r + dr) if move_row else str(r)
            out.append(f"{col_abs}{col_text}{row_abs}{row_text}")
        return "".join(out)
